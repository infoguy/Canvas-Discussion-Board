"""
app.py — Canvas Discussion Builder, the web version.

Same job as the desktop db_creator.py: take a Canvas template layout and a pile
of Word docs (or typed header/content pairs), fill each region of the template
with the matching part of the content, and create the discussions in Canvas.

What changed for the web:
  • No Selenium and no Chrome. Discussions are created through the Canvas REST
    API with an access token you paste in per run; the token is never stored.
  • No Tkinter. The form lives in templates/index.html.
  • The merge that used to run as JavaScript inside the browser now runs in
    merge.py against BeautifulSoup.

Run locally:  python app.py     →  http://127.0.0.1:5000
On Render:    gunicorn app:app  (see render.yaml)
"""

import os
import tempfile
import traceback

from flask import Flask, jsonify, render_template, request

import content as C
import merge as M
from canvas_api import CanvasClient, CanvasError, split_course_url, split_topic_url

app = Flask(__name__)
app.config["MAX_CONTENT_LENGTH"] = 25 * 1024 * 1024   # 25 MB of Word docs

DEFAULT_SETTINGS = {
    "apply": True,
    "disallow_threaded": False,
    "require_initial_post": False,
    "podcast_feed": False,
    "graded": True,
    "allow_liking": False,
    "group_discussion": False,
    "thread_state": "Collapsed",        # Expanded | Collapsed
    "sort_order": "Newest First",       # Oldest First | Newest First
    "points_possible": "0",
    "display_grade_as": "Points",
    "assignment_group": "Assignments",
    "peer_reviews": "Off",              # Off | Assign manually | Automatically assign
}


# ── helpers ──────────────────────────────────────────────────────────────────
def _course_urls(raw):
    """One Canvas URL per line, blanks and stray text ignored."""
    if isinstance(raw, list):
        lines = raw
    else:
        lines = (raw or "").splitlines()
    return [u.strip() for u in lines if u.strip().startswith("http")]


def _titles_from(raw):
    if isinstance(raw, list):
        lines = raw
    else:
        lines = (raw or "").splitlines()
    return [t.strip() for t in lines if t.strip()]


def _resolve_entry(entry, main_template, options):
    """
    Turn one entry from the form into (title, body_html, sections, template).

    A docx entry arrives already parsed by /api/parse-docx. A typed entry is
    parsed here, because routing its headers needs the template it will be
    built against — which may be the entry's own template override.
    """
    own = (entry.get("template") or "").strip()
    template = C.normalize_template_html(own) if own else main_template

    if entry.get("kind") == "typed":
        pairs = [(p.get("header", ""), p.get("content", ""))
                 for p in (entry.get("pairs") or [])]
        title, body_html, sections = C.sections_from_pairs(
            pairs, title=entry.get("title") or "", template_html=template)
    else:
        title = entry.get("title") or "Untitled Discussion"
        body_html = entry.get("body_html") or ""
        sections = entry.get("sections") or {}

    if entry.get("title_override"):
        title = entry["title_override"].strip() or title

    if options.get("week_prefix", True):
        title = C.apply_week_prefix(title, (sections or {}).get("_week"))

    return title, body_html, sections, template, bool(own)


def _build(entry, main_template, options):
    """Compose one finished discussion body. Returns (title, html, log)."""
    title, body_html, sections, template, own = _resolve_entry(
        entry, main_template, options)
    log = [f'   ➕ {title}' + ("   🧩 own template" if own else "")]
    if not template:
        log.append("   ❌ No template HTML for this entry.")
        return title, None, log

    page_title = C.title_for_page(title)
    html, merge_log = M.compose_body(
        template, body_html,
        title=page_title,
        anchor=options.get("anchor") or "Instructions",
        update_banner=options.get("update_banner", True),
        sections=sections,
        drop_leftovers=options.get("drop_leftovers", False),
    )
    log.extend(merge_log)
    if not html:
        log.append("   ❌ Nothing was built for this entry.")
    return title, html, log


def _settings_from(payload):
    settings = dict(DEFAULT_SETTINGS)
    for k, v in (payload.get("settings") or {}).items():
        if k in settings:
            settings[k] = v
    return settings


def _fail(message, code=400):
    return jsonify({"ok": False, "error": message}), code


# ── routes ───────────────────────────────────────────────────────────────────
@app.get("/")
def index():
    return render_template("index.html", defaults=DEFAULT_SETTINGS)


@app.get("/healthz")
def healthz():
    return {"ok": True}


@app.post("/api/parse-docx")
def parse_docx():
    """Upload one or more Word docs; hand back the parsed entries."""
    files = request.files.getlist("files")
    if not files:
        return _fail("No Word documents were uploaded.")

    entries, errors = [], []
    for f in files:
        name = f.filename or "document.docx"
        if not name.lower().endswith((".docx", ".docm")):
            errors.append(f"{name}: not a Word document.")
            continue
        suffix = ".docx"
        tmp = tempfile.NamedTemporaryFile(suffix=suffix, delete=False)
        try:
            f.save(tmp.name)
            tmp.close()
            title, body_html, sections = C.parse_discussion_docx(tmp.name)
            # The parser reads the week number off the filename when the doc
            # itself does not say, so keep the real name in play.
            if not sections.get("_week"):
                import re
                m = re.search(r"week\s*_?\s*#?\s*(\d+)", name, re.I)
                if m:
                    sections["_week"] = int(m.group(1))
            entries.append({
                "kind": "docx",
                "file": name,
                "title": title,
                "body_html": body_html,
                "sections": sections,
            })
        except Exception as e:
            errors.append(f"{name}: {e}")
        finally:
            try:
                os.unlink(tmp.name)
            except OSError:
                pass

    return jsonify({"ok": True, "entries": entries, "errors": errors})


@app.post("/api/template")
def load_template():
    """Read a template's layout HTML straight out of an existing discussion."""
    data = request.get_json(silent=True) or {}
    url = (data.get("url") or "").strip()
    token = (data.get("token") or "").strip()
    if not url:
        return _fail("Paste the URL of the discussion to copy the layout from.")
    if not token:
        return _fail("A Canvas access token is needed to read the template.")
    try:
        base, course_id, topic_id = split_topic_url(url)
        client = CanvasClient(base, token)
        html = client.topic_message(course_id, topic_id)
    except CanvasError as e:
        return _fail(str(e))
    if not html.strip():
        return _fail("That discussion has an empty body, so there is no layout to copy.")

    html = C.normalize_template_html(html)
    has_placeholder = "lorem ipsum" in html.lower()
    return jsonify({
        "ok": True,
        "html": html,
        "chars": len(html),
        "note": ("" if has_placeholder else
                 "No 'Lorem Ipsum' in this layout, so content will go into the "
                 "innermost shaded box instead."),
    })


@app.post("/api/route-headers")
def route_headers():
    """
    Where each typed header will send its content, for the ✓ next to the field.

    Same answer the desktop popup gave: a header the template has a spot for is
    matched to it; anything else is reported as a banner heading of its own, so
    nothing is a surprise at build time.
    """
    data = request.get_json(silent=True) or {}
    template = C.normalize_template_html(data.get("template_html") or "")
    out = []
    for raw in data.get("headers") or []:
        name = (raw or "").strip().rstrip(":")
        if not name:
            out.append({"header": raw, "route": "", "note": ""})
            continue
        route, match = C.classify_header(name, template)
        note = {
            "title": "Sets the discussion title.",
            "template": f'Goes under the template\'s "{match}" heading.',
            "prompt": "Goes into the prompt box.",
            "instructions": "Goes under the instructions heading.",
            "objective": "Goes into the objective spot.",
            "new": "The template has no spot with this name, so it is added as "
                   "a banner heading copied from the template.",
        }.get(route, "")
        out.append({"header": raw, "route": route, "match": match, "note": note,
                    "ok": route != "new"})
    return jsonify({"ok": True, "headers": out,
                    "template_headings": C.template_headings(template)})


@app.post("/api/preview")
def preview():
    """Check first: build every discussion body without touching Canvas."""
    data = request.get_json(silent=True) or {}
    options = data.get("options") or {}
    entries = data.get("entries") or []
    main_template = C.normalize_template_html(data.get("template_html") or "")
    courses = _course_urls(data.get("courses"))
    titles = _titles_from(data.get("titles"))

    if not entries:
        return _fail("Add content for at least one discussion: a Word document, "
                     "or a typed entry.")
    if not main_template and not any((e.get("template") or "").strip() for e in entries):
        return _fail("Paste the template's layout HTML, or load it from a "
                     "discussion URL.")
    if titles and len(titles) < len(entries):
        return _fail(f"{len(entries)} entries loaded but only {len(titles)} title(s). "
                     f"Add a title per line, or clear the box to use the entry titles.")

    log = []
    for i, url in enumerate(courses, 1):
        try:
            base, course_id = split_course_url(url)
            log.append(f"📍 Course {i}: {base} · course {course_id}")
        except CanvasError as e:
            log.append(f"⚠️ {e}")
    if not courses:
        log.append("⚠️ No course URLs yet. Nothing will be created until you add one.")

    built = []
    for n, entry in enumerate(entries):
        if titles:
            entry = dict(entry, title_override=titles[n])
        title, html, entry_log = _build(entry, main_template, options)
        log.extend(entry_log)
        built.append({
            "title": title,
            "ok": bool(html),
            "html": html or "",
            "chars": len(html or ""),
        })

    ready = sum(1 for b in built if b["ok"])
    log.append(f"\n🔍 Preview only, nothing was created. {ready} of {len(built)} "
               f"entr{'y' if len(built) == 1 else 'ies'} ready.")
    return jsonify({"ok": True, "entries": built, "log": log})


@app.post("/api/create")
def create():
    """Build every body, then create the discussions in each course."""
    data = request.get_json(silent=True) or {}
    token = (data.get("token") or "").strip()
    options = data.get("options") or {}
    entries = data.get("entries") or []
    settings = _settings_from(data)
    main_template = C.normalize_template_html(data.get("template_html") or "")
    courses = _course_urls(data.get("courses"))
    titles = _titles_from(data.get("titles"))
    publish = bool(options.get("publish"))

    if not token:
        return _fail("Paste your Canvas access token.")
    if not courses:
        return _fail("Enter at least one Canvas course URL, one per line.")
    if not entries:
        return _fail("Add content for at least one discussion.")
    if not main_template and not any((e.get("template") or "").strip() for e in entries):
        return _fail("Paste the template's layout HTML, or load it from a "
                     "discussion URL.")
    if titles and len(titles) < len(entries):
        return _fail(f"{len(entries)} entries loaded but only {len(titles)} title(s).")

    log, created, failed, links = [], 0, 0, []

    # Build everything first, so a bad template fails before anything is made
    built = []
    for n, entry in enumerate(entries):
        if titles:
            entry = dict(entry, title_override=titles[n])
        title, html, entry_log = _build(entry, main_template, options)
        log.extend(entry_log)
        built.append((title, html))
    if not any(html for _t, html in built):
        log.append("❌ Nothing could be built from the template, so nothing was created.")
        return jsonify({"ok": False, "log": log, "created": 0, "failed": len(built)})

    log.append(f"\n=== Creating {len(built)} discussion(s) in {len(courses)} course(s) "
               f"{'(published)' if publish else '(unpublished)'} ===")
    if settings.get("apply", True):
        log.append("   Settings per discussion: "
                   f"{'Graded' if settings.get('graded') else 'Ungraded'}, "
                   f"{settings.get('thread_state')} threads, "
                   f"{settings.get('sort_order')}, "
                   f"{settings.get('points_possible')} pts, "
                   f"{settings.get('display_grade_as')}, "
                   f"group '{settings.get('assignment_group')}', "
                   f"peer reviews {settings.get('peer_reviews')}")
    else:
        log.append("   Settings panel is off — Canvas defaults will be used.")

    for i, url in enumerate(courses, 1):
        try:
            base, course_id = split_course_url(url)
            client = CanvasClient(base, token)
            log.append(f"\n[{i}/{len(courses)}] {client.course_name(course_id)} "
                       f"({base}/courses/{course_id})")
        except CanvasError as e:
            log.append(f"\n[{i}/{len(courses)}] ❌ {e}")
            failed += sum(1 for _t, html in built if html)
            continue

        course_settings = dict(settings)
        if settings.get("apply", True) and settings.get("graded"):
            try:
                gid = client.assignment_group_id(course_id, settings.get("assignment_group"))
            except CanvasError as e:
                gid = None
                log.append(f"   ⚠️ Could not read the assignment groups: {e}")
            if gid:
                course_settings["assignment_group_id"] = gid
            elif settings.get("assignment_group"):
                log.append(f"   ⚠️ No assignment group named "
                           f"\"{settings.get('assignment_group')}\" in this course — "
                           f"Canvas will use its default group.")

        for title, html in built:
            if not html:
                failed += 1
                continue
            try:
                topic = client.create_discussion(
                    course_id, title, html, published=publish, settings=course_settings)
                created += 1
                link = (topic or {}).get("html_url") or ""
                if link:
                    links.append({"title": title, "url": link})
                log.append(f"   ✅ {title}" + (f" → {link}" if link else ""))
            except CanvasError as e:
                failed += 1
                log.append(f"   ❌ {title}: {e}")
            except Exception as e:                       # pragma: no cover
                failed += 1
                log.append(f"   ❌ {title}: unexpected error: {e}")

    log.append(f"\n{'=' * 50}")
    log.append(f"✅ Created: {created}   ❌ Failed: {failed}")
    if created:
        log.append("🎉 All done.")
    return jsonify({"ok": failed == 0, "log": log, "created": created,
                    "failed": failed, "links": links})


@app.errorhandler(413)
def too_large(_e):
    return _fail("Those files are larger than the 25 MB upload limit.", 413)


@app.errorhandler(500)
def server_error(_e):                                    # pragma: no cover
    traceback.print_exc()
    return _fail("Something went wrong on the server. Check the Render logs.", 500)


if __name__ == "__main__":
    app.run(host="0.0.0.0", port=int(os.environ.get("PORT", 5000)), debug=True)
