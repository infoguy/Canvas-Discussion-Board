"""
canvas_api.py — talk to Canvas over REST instead of driving the form in Chrome.

The desktop tool clicked "+ Discussion", typed into TinyMCE, and set every
option on the settings panel by hand. On a server there is no browser, so the
same job is done with one POST to /api/v1/courses/:id/discussion_topics.

The access token is passed in per run and never written to disk.
"""

import re
from urllib.parse import urlparse

import requests

TIMEOUT = 45


class CanvasError(Exception):
    pass


# ── URL parsing ──────────────────────────────────────────────────────────────
def split_course_url(url: str):
    """
    ('https://canvas.americancareercollege.edu', '8497') from any course URL.

    Accepts the course home page, the discussions index, or a discussion URL.
    """
    u = (url or "").strip()
    if not u.startswith("http"):
        raise CanvasError(f"Not a Canvas URL: {url}")
    parts = urlparse(u)
    base = f"{parts.scheme}://{parts.netloc}"
    m = re.search(r"/courses/(\d+)", parts.path)
    if not m:
        raise CanvasError(f"No /courses/<id> in this URL: {url}")
    return base, m.group(1)


def split_topic_url(url: str):
    """('base', 'course_id', 'topic_id') from a discussion URL."""
    base, course_id = split_course_url(url)
    m = re.search(r"/discussion_topics/(\d+)", urlparse(url).path)
    if not m:
        raise CanvasError("That template URL has no /discussion_topics/<id> in it.")
    return base, course_id, m.group(1)


# ── grade / thread option mapping ────────────────────────────────────────────
GRADING_TYPES = {
    "points": "points",
    "percentage": "percent",
    "percent": "percent",
    "complete/incomplete": "pass_fail",
    "complete / incomplete": "pass_fail",
    "letter grade": "letter_grade",
    "gpa scale": "gpa_scale",
    "not graded": "not_graded",
}


def _grading_type(display_grade_as: str) -> str:
    return GRADING_TYPES.get((display_grade_as or "").strip().lower(), "points")


class CanvasClient:
    def __init__(self, base_url: str, token: str):
        self.base = (base_url or "").rstrip("/")
        self.token = (token or "").strip()
        if not self.token:
            raise CanvasError("No Canvas access token was supplied.")
        self.session = requests.Session()
        self.session.headers.update({
            "Authorization": f"Bearer {self.token}",
            "Accept": "application/json",
        })

    # ── plumbing ─────────────────────────────────────────────────────────────
    def _url(self, path: str) -> str:
        return f"{self.base}/api/v1/{path.lstrip('/')}"

    def _request(self, method: str, path: str, **kwargs):
        try:
            r = self.session.request(method, self._url(path), timeout=TIMEOUT, **kwargs)
        except requests.RequestException as e:
            raise CanvasError(f"Could not reach Canvas: {e}") from e

        if r.status_code == 401:
            raise CanvasError("Canvas rejected the access token (401). "
                              "Generate a new one under Account → Settings.")
        if r.status_code == 403:
            raise CanvasError("That token does not have permission for this course (403).")
        if r.status_code == 404:
            raise CanvasError(f"Canvas has no {path} (404). Check the course URL.")
        if r.status_code >= 400:
            detail = ""
            try:
                data = r.json()
                detail = data.get("message") or str(data.get("errors") or data)[:300]
            except Exception:
                detail = (r.text or "")[:300]
            raise CanvasError(f"Canvas said {r.status_code}: {detail}")
        if not r.content:
            return None
        try:
            return r.json()
        except ValueError:
            return None

    def get(self, path, **kw):
        return self._request("GET", path, **kw)

    def post(self, path, payload):
        return self._request("POST", path, json=payload)

    # ── reads ────────────────────────────────────────────────────────────────
    def whoami(self) -> str:
        me = self.get("users/self") or {}
        return me.get("name") or me.get("short_name") or "this account"

    def course_name(self, course_id: str) -> str:
        c = self.get(f"courses/{course_id}") or {}
        return c.get("name") or c.get("course_code") or f"course {course_id}"

    def topic_message(self, course_id: str, topic_id: str) -> str:
        """The body HTML of an existing discussion, for use as the template."""
        t = self.get(f"courses/{course_id}/discussion_topics/{topic_id}") or {}
        return t.get("message") or ""

    def assignment_group_id(self, course_id: str, name: str):
        """Look up an assignment group by name; None when nothing matches."""
        want = (name or "").strip().lower()
        if not want:
            return None
        groups = self.get(f"courses/{course_id}/assignment_groups",
                          params={"per_page": 100}) or []
        for g in groups:
            if (g.get("name") or "").strip().lower() == want:
                return g.get("id")
        for g in groups:
            gname = (g.get("name") or "").strip().lower()
            if want in gname or gname in want:
                return g.get("id")
        return None

    # ── writes ───────────────────────────────────────────────────────────────
    def create_discussion(self, course_id: str, title: str, message: str,
                          published: bool = False, settings: dict = None,
                          group_category_id=None):
        """
        Create one discussion topic. Returns the Canvas topic dict.

        `settings` uses the same keys as the desktop tool's settings panel.
        """
        settings = dict(settings or {})
        payload = {
            "title": title,
            "message": message,
            "published": bool(published),
        }

        if settings.get("apply", True):
            payload["discussion_type"] = (
                "side_comment" if settings.get("disallow_threaded") else "threaded")
            payload["require_initial_post"] = bool(settings.get("require_initial_post"))
            payload["podcast_enabled"] = bool(settings.get("podcast_feed"))
            payload["allow_rating"] = bool(settings.get("allow_liking"))
            payload["sort_order"] = (
                "desc" if str(settings.get("sort_order", "")).lower().startswith("newest")
                else "asc")
            payload["expanded"] = str(settings.get("thread_state", "")).lower() == "expanded"

            if settings.get("group_discussion") and group_category_id:
                payload["group_category_id"] = group_category_id

            if settings.get("graded"):
                peer = str(settings.get("peer_reviews", "Off")).strip().lower()
                try:
                    points = float(settings.get("points_possible") or 0)
                except (TypeError, ValueError):
                    points = 0
                assignment = {
                    "points_possible": points,
                    "grading_type": _grading_type(settings.get("display_grade_as")),
                    "peer_reviews": peer != "off",
                    "automatic_peer_reviews": peer.startswith("automatic"),
                }
                group_id = settings.get("assignment_group_id")
                if group_id:
                    assignment["assignment_group_id"] = group_id
                payload["assignment"] = assignment

        return self.post(f"courses/{course_id}/discussion_topics", payload)
