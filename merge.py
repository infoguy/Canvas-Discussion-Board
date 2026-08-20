"""
merge.py — fill a Canvas template with content, region by region.

This is a straight port of the two chunks of browser JavaScript the desktop
tool ran through Selenium (_MERGE_JS and _SWAP_BOX_JS), rewritten against
BeautifulSoup so it runs on a server with no Chrome anywhere in sight.

Each region is found by its own heading ("Discussion Prompt", "Instructions",
"Helpful Tips") and only the content under that heading is replaced, so the
banner table, the dividers, and the tips box all stay exactly where they are.
Regions with no heading fall back to their [ ... ] placeholder. Anything left
in brackets is reported so nothing ships with template filler still in it.
"""

import re
from bs4 import BeautifulSoup, NavigableString, Tag

from content import (
    banner_heading_html,
    html_to_lines,
)

# The section headings a template is allowed to have. These double as the
# boundaries between regions and as the headings the banner rename must skip.
TEMPLATE_SECTION_LABELS = [
    "topic introduction", "submissions", "this week's prompt",
    "discussion prompt", "prompt", "scenario", "overview", "background",
    "summary", "discussion instructions", "instructions", "directions",
    "task", "requirements", "questions", "discussion questions",
    "initial post", "response requirements", "step 1", "step 2",
    "helpful tips", "helpful tip", "tips", "resources", "rubric",
    "evaluation", "grading", "objective", "objectives",
]

HEADINGS = ["h1", "h2", "h3", "h4", "h5", "h6"]
CELL = {"td", "th", "tr", "table", "tbody", "thead", "tfoot"}
BLOCK = {"p", "div", "li", "section", "header", "h1", "h2", "h3", "h4", "h5", "h6"}
TEXT_TAGS = {"p", "li", "td"}

_BRACKET_RE = re.compile(r"\[[^\[\]]{0,240}\]")
# An author instruction ("[Insert overview here]") is filler and its line goes.
# A bracketed sentence that already reads as finished content is kept, with
# just the brackets stripped off.
AUTHOR_FILLER = re.compile(
    r"^\[\s*(insert|list|enter|type|add|replace|choose|select|provide|specify|"
    r"describe|write|fill|numeral|number of|tbd|n/a)\b", re.I)
# "Step 1", "Part 2", "Week 3" and friends number a layout, they are not a spot
# for the title. Renaming one wrecks the template.
STRUCTURAL = re.compile(
    r"^(step|part|section|phase|week|day|module|unit|topic|question)\s*#?\s*\d+", re.I)


# ── small helpers ────────────────────────────────────────────────────────────
def _wrap(html):
    """Parse a fragment and hand back its wrapper div, the way the JS did."""
    soup = BeautifulSoup(f'<div id="__wrap__">{html or ""}</div>', "html.parser")
    return soup.find(id="__wrap__")


def _text(el):
    if el is None:
        return ""
    if isinstance(el, NavigableString):
        return str(el)
    return el.get_text()


def norm(t):
    return re.sub(r"\s+", " ", (t or "").replace("\u00a0", " ")).strip().lower()


def label(t):
    return re.sub(r"[\s:.\-\u2010-\u2015]+$", "", norm(t))


def words(t):
    n = norm(t)
    return len(n.split(" ")) if n else 0


def is_heading(el):
    return isinstance(el, Tag) and el.name in HEADINGS


def level(el):
    return int(el.name[1]) if is_heading(el) else 99


def _contains(box, node):
    """True when node sits somewhere under box."""
    cur = node.parent if node is not None else None
    while cur is not None:
        if cur is box:
            return True
        cur = cur.parent
    return False


def _child_tags(el):
    return [c for c in el.children if isinstance(c, Tag)]


def block_of(el, root):
    """
    Climb from an inline node to the block that owns the line.

    Never climbs out into a table cell, which is what let content spill across
    the banner columns.
    """
    cur = el
    while cur is not None and cur is not root and cur.parent is not None:
        if isinstance(cur, Tag):
            if cur.name in BLOCK:
                return cur
            if cur.name in CELL:
                return el
        cur = cur.parent
    return el


def matches_name(text, names):
    """
    "Assignment Instructions" matches a template heading of just
    "Instructions", and the other way round. Only short lines are eligible so a
    paragraph that merely mentions the word is never mistaken for a heading.
    """
    t = label(text)
    if not t:
        return False
    for n in names:
        w = label(n)
        if not w:
            continue
        if t == w:
            return True
        if len(t) > 3 and len(w) > 3 and (w in t or t in w):
            return True
    return False


def _attrs_of(el):
    if not isinstance(el, Tag):
        return None
    cls = el.get("class") or []
    if isinstance(cls, list):
        cls = " ".join(cls)
    return {"style": el.get("style") or "", "cls": cls}


def _donors_from(nodes):
    """Pick up the styling of the template blocks being replaced."""
    d = {"p": None, "li": None, "ul": None, "ol": None}

    def grab(el, tag):
        if not isinstance(el, Tag):
            return None
        return _attrs_of(el if el.name == tag else el.find(tag))

    for n in nodes:
        for tag in ("p", "li", "ul", "ol"):
            d[tag] = d[tag] or grab(n, tag)
    d["ul"] = d["ul"] or d["ol"]
    d["ol"] = d["ol"] or d["ul"]
    return d


def _apply_attrs(el, a):
    if not a:
        return
    if a.get("style"):
        el["style"] = a["style"]
    if a.get("cls"):
        el["class"] = a["cls"]


def _style_frag(frag, donors):
    for tag in ("p", "li", "ul", "ol"):
        for el in frag.find_all(tag):
            _apply_attrs(el, donors.get(tag))


def _set_deep_text(el, txt):
    """Write text into the innermost single wrapper, keeping icon spans intact."""
    cur = el
    while True:
        kids = [k for k in cur.children
                if isinstance(k, Tag) or (isinstance(k, NavigableString) and str(k).strip())]
        if len(kids) == 1 and isinstance(kids[0], Tag):
            cur = kids[0]
            continue
        break
    cur.clear()
    cur.append(NavigableString(txt))


def _move_into(frag, insert_before=None, parent=None):
    """Move a fragment's children into the document, in order."""
    moved = []
    for kid in list(frag.contents):
        kid.extract()
        if insert_before is not None:
            insert_before.insert_before(kid)
        else:
            parent.append(kid)
        moved.append(kid)
    return moved


def prune_empty(el, root):
    """Take away an emptied line's leftover wrapper so no bare bullet is left."""
    cur = el
    while cur is not None and cur is not root and cur.parent is not None:
        if not isinstance(cur, Tag) or cur.name not in ("li", "p", "ul", "ol"):
            break
        if norm(cur.get_text()) != "":
            break
        if cur.find(["img", "iframe", "hr", "table", "input", "video", "br"]):
            break
        par = cur.parent
        cur.extract()
        cur = par


# ── finding the spot ─────────────────────────────────────────────────────────
def _headings(root):
    return root.find_all(HEADINGS)


def find_label(root, names):
    hs = _headings(root)
    for n in names:
        for h in hs:
            if label(h.get_text()) == label(n):
                return h
    for h in hs:
        if words(h.get_text()) <= 6 and matches_name(h.get_text(), names):
            return h
    # Non-heading candidates must match a section name exactly. Loose matching
    # here would latch onto any short line that merely contains the word, such
    # as a placeholder reading "Lorem Ipsum placeholder prompt goes here".
    cands = root.find_all(["p", "div", "strong", "b", "span", "td", "th"])
    for n in names:
        for c in cands:
            t = c.get_text()
            if words(t) == 0 or words(t) > 6:
                continue
            if label(t) == label(n):
                return block_of(c, root)
    return None


def is_boundary(el, lvl, section_labels):
    if not isinstance(el, Tag):
        return False
    if is_heading(el):
        return level(el) <= lvl or lvl == 99
    if el.name == "hr":
        return True
    if el.find(HEADINGS + ["hr"]):
        return True
    if words(el.get_text()) <= 6:
        t = label(el.get_text())
        for b in section_labels:
            if t == label(b):
                return True
    return False


def is_filler_line(t):
    """
    A template's real content spot is the filler it ships with: a Lorem Ipsum
    box, a [ ... ] placeholder, an empty line.
    """
    s = norm(t)
    if not s:
        return True
    if re.match(r"^\[[^\]]*\]$", s):
        return True
    if "lorem ipsum" in s:
        return True
    if re.match(r"^(lorem|loprem|ipsum)\b", s):
        return True
    if re.match(r"^(placeholder|sample text|text here|your text here|insert |type |add your)", s):
        return True
    return False


def is_leaf_block(el):
    if not isinstance(el, Tag):
        return False
    if el.find(["p", "li", "ul", "ol", "div", "table"] + HEADINGS):
        return False
    if el.find(["img", "iframe", "video", "input"]):
        return False
    return True


def leaf_blocks(box):
    return [el for el in box.find_all(["p", "li", "td"] + HEADINGS) if is_leaf_block(el)]


def section_after(target):
    """
    Everything after the heading that still belongs to it. Only another heading
    of the same or higher rank ends the section here; a decorative <hr> does not.
    """
    lvl = level(target)
    out = []
    for n in target.next_siblings:
        if not isinstance(n, Tag):
            continue
        if is_heading(n) and (level(n) <= lvl or lvl == 99):
            break
        stop = any(level(h) <= lvl for h in n.find_all(HEADINGS))
        if stop:
            break
        out.append(n)
    return out


def top_ancestor(el, root):
    cur = el
    while cur is not None and cur.parent is not None and cur.parent is not root:
        cur = cur.parent
    return cur


def holds_heading_up_to(el, lvl):
    if is_heading(el) and level(el) <= lvl:
        return True
    if isinstance(el, Tag):
        return any(level(h) <= lvl for h in el.find_all(HEADINGS))
    return False


def holds_section_label(el, section_labels):
    """
    A row that opens with a section label of its own ("Instructions", "Helpful
    Tips") belongs to that section, not to the heading above it, even when the
    label is a bold line rather than a real heading tag.
    """
    if not isinstance(el, Tag):
        return False
    cands = [el] + el.find_all(["p", "li", "td", "strong", "b", "span"])
    for c in cands:
        if not 0 < words(c.get_text()) <= 6:
            continue
        t = label(c.get_text())
        if any(t == label(s) for s in section_labels):
            return True
    return False


def wide_section_after(target, root, section_labels):
    """
    Loree layouts wrap the heading in one row <div> and the content box that
    belongs to it in the NEXT row, so the box is not a sibling of the heading at
    all. When the heading's own siblings hold no filler, widen the search to the
    rows that follow the heading's row and stop at the next heading of the same
    or higher rank.
    """
    lvl = level(target)
    top = top_ancestor(target, root)
    if top is None:
        return []
    out = []
    for n in top.next_siblings:
        if not isinstance(n, Tag):
            continue
        if holds_heading_up_to(n, lvl):
            break
        if holds_section_label(n, section_labels):
            break
        out.append(n)
    return out


def fillers_in(scope):
    out = []
    for el in scope:
        cands = el.find_all(["p", "li", "td"]) if isinstance(el, Tag) else []
        # A layout <div> is never filler itself, only the text lines inside it.
        # Counting empty wrappers (a divider, a spacer) would drag the shared
        # container up to the whole section and fill the wrong place.
        if isinstance(el, Tag) and el.name in TEXT_TAGS:
            cands = [el] + list(cands)
        for c in cands:
            if is_heading(c):
                continue          # a heading is never filler
            if not is_leaf_block(c):
                continue
            if any(c is o for o in out):
                continue
            if is_filler_line(c.get_text()):
                out.append(c)
    return out


def insert_point(el, box):
    """
    Where to drop content in place of a filler line: outside any list wrapper it
    sits in, so bullets do not end up nested inside a leftover <li>.
    """
    cur = el
    while cur is not None:
        par = cur.parent
        if par is None or par is box or par.name not in ("li", "ul", "ol"):
            break
        cur = par
    return cur


def box_of(fillers):
    """The single container that holds every filler line handed to it."""
    box = fillers[0].parent if fillers else None
    for f in fillers[1:]:
        while box is not None and not _contains(box, f):
            box = box.parent
    return box


def mostly_filler(box, fillers, root, target=None):
    """
    Only wipe a box that is mostly sample content, never one where the filler
    is a minor part of real template copy.
    """
    if box is None or box is root:
        return False
    if target is not None and (_contains(box, target) or box is target):
        return False
    leaves = leaf_blocks(box)
    return bool(leaves and len(fillers) * 2 >= len(leaves))


def box_donors(box, state):
    """
    The template's own list and paragraph styling, remembered from the sample
    content the first time a box is emptied. Without this a section added after
    the wipe has no white type or bullet styling left to copy.
    """
    live = _donors_from(_child_tags(box))
    saved = None
    for entry in state.get("donors", []):
        if entry["box"] is box:
            saved = entry["donors"]
            break
    if saved:
        for tag in ("p", "li", "ul", "ol"):
            live[tag] = live[tag] or saved[tag]
    return live


def wipe_into(box, html, state):
    donors = box_donors(box, state)
    if not any(e["box"] is box for e in state.setdefault("donors", [])):
        state["donors"].append({"box": box, "donors": donors})
    frag = _wrap(html)
    _style_frag(frag, donors)
    box.clear()
    return _move_into(frag, parent=box)


def put_into(box, html, before, state):
    """Add a section to a box that already has content, keeping its colours."""
    frag = _wrap(html)
    _style_frag(frag, box_donors(box, state))
    if before is not None and before.parent is box:
        return _move_into(frag, insert_before=before)
    return _move_into(frag, parent=box)


def header_line(box, text, state):
    """A bold header line in the box's own type."""
    donors = box_donors(box, state)
    soup = BeautifulSoup("<p><strong></strong></p>", "html.parser")
    p = soup.find("p")
    _apply_attrs(p, donors.get("p"))
    p.find("strong").append(NavigableString(text))
    p.extract()
    return p


def record_placement(state, box, node):
    if box is None:
        return
    if node is not None:
        state["placements"].append({"box": box, "node": node,
                                    "order": state["order"],
                                    "label": state["label"]})
    state["last_box"] = box


def anchor_in(state, box):
    """
    The earliest thing already in this box that the doc puts AFTER the section
    being placed now. New content goes in front of it.
    """
    best = None
    for pl in state["placements"]:
        if pl["box"] is not box or pl["order"] <= state["order"]:
            continue
        if pl["node"] is None or pl["node"].parent is not box:
            continue
        if best is None or pl["order"] < best["order"]:
            best = pl
    return best["node"] if best else None


def filler_fill(target, html, root, state, section_labels):
    fillers = fillers_in(section_after(target))
    # Loree layouts keep the heading in one row and its box in the next, so a
    # sibling-only search finds nothing and the content used to be dropped loose
    # under the heading with the coloured box left full of Lorem Ipsum.
    if not fillers:
        fillers = fillers_in(wide_section_after(target, root, section_labels))
    if not fillers:
        return -1

    box = box_of(fillers)

    if mostly_filler(box, fillers, root, target):
        # When the filler was a sample bullet list and the new content is not
        # all bullets, the list wrapper goes with it. Filling it instead would
        # leave paragraphs sitting inside a <ul>, which Canvas then mangles.
        frag = _wrap(html)
        _style_frag(frag, box_donors(box, state))
        incoming = [c for c in frag.children if getattr(c, "name", None)]
        if (box.name or "").lower() in {"ul", "ol"} and box.parent is not None \
                and not all((c.name or "").lower() == "li" for c in incoming):
            host = box.parent
            moved = _move_into(frag, insert_before=box)
            box.extract()
            record_placement(state, host, moved[0] if moved else None)
            return len(fillers)
        moved = wipe_into(box, html, state)
        record_placement(state, box, moved[0] if moved else None)
        return len(fillers)

    anchor = insert_point(fillers[0], box or root)
    if anchor is None or anchor.parent is None:
        return -1
    host = anchor.parent
    frag = _wrap(html)
    _style_frag(frag, _donors_from(fillers))
    moved = _move_into(frag, insert_before=anchor)
    for f in fillers:
        fp = f.parent
        if fp is not None:
            f.extract()
            prune_empty(fp, root)
    landed = box if (box is not None and box is not root) else host
    record_placement(state, landed, moved[0] if moved else None)
    return len(fillers)


def place_loose(html, root, state, section_labels):
    """
    A section the template has no heading and no [ ] spot for still belongs
    inside the layout: it joins the box the section before it filled, or takes
    over whatever sample box is still untouched. Appending to the bottom of the
    page is the last resort, never the first move.
    """
    box = state.get("last_box")
    if box is not None and box.parent is not None:
        moved = put_into(box, html, anchor_in(state, box), state)
        record_placement(state, box, moved[0] if moved else None)
        return "box"
    fillers = fillers_in([root])
    if not fillers:
        return None
    box = box_of(fillers)
    if not mostly_filler(box, fillers, root):
        return None
    moved = wipe_into(box, html, state)
    record_placement(state, box, moved[0] if moved else None)
    return "free"


def heading_before(root, box):
    """The nearest heading that comes before this box in the page."""
    best = ""
    for h in _headings(root):
        if h is box or _contains(h, box):
            continue
        if _precedes(h, box, root):
            best = h.get_text()
    return best


def _doc_index(root, node):
    """Position of a node in document order, for before/after comparisons."""
    for i, el in enumerate(root.descendants):
        if el is node:
            return i
    return -1


def _precedes(a, b, root):
    ia, ib = _doc_index(root, a), _doc_index(root, b)
    return ia >= 0 and ib >= 0 and ia < ib


def restore_headers(root, state):
    """
    The header the doc gave each section. A section at the top of a box needs
    none when the template's own heading right above the box already says it;
    every other section gets its header back, which is how the Word doc reads.
    """
    groups = []
    for pl in state["placements"]:
        if pl["node"] is None or pl["node"].parent is not pl["box"]:
            continue
        slot = next((g for g in groups if g["box"] is pl["box"]), None)
        if slot is None:
            slot = {"box": pl["box"], "list": []}
            groups.append(slot)
        slot["list"].append(pl)

    for g in groups:
        box, items = g["box"], g["list"]
        kids = list(box.children)

        def pos(pl):
            for i, k in enumerate(kids):
                if k is pl["node"]:
                    return i
            return len(kids)

        items.sort(key=pos)
        above = heading_before(root, box)
        for i, pl in enumerate(items):
            text = (pl.get("label") or "").strip()
            if not text:
                continue
            # The template's own heading just above the box already says it
            if i == 0 and above and label(above) == label(text):
                continue
            prev = pl["node"].find_previous_sibling()
            if prev is not None and label(prev.get_text()) == label(text):
                continue
            pl["node"].insert_before(header_line(box, text, state))


def replace_under(target, html, root, section_labels, state):
    """
    Swap everything between a heading and the next section boundary. Sibling
    walking keeps this inside the heading's own container, so a heading that
    lives in a banner cell can never eat the cell beside it.
    """
    if target.name in CELL:
        return -1
    # Fill the template's own placeholder first; the sibling sweep below is the
    # fallback for templates that ship no filler at all.
    by_filler = filler_fill(target, html, root, state, section_labels)
    if by_filler >= 0:
        return by_filler

    lvl = level(target)
    removed, stop = [], None
    for n in list(target.next_siblings):
        if isinstance(n, Tag) and is_boundary(n, lvl, section_labels):
            stop = n
            break
        if isinstance(n, Tag) or (isinstance(n, NavigableString) and str(n).strip()):
            removed.append(n)

    frag = _wrap(html)
    _style_frag(frag, _donors_from([r for r in removed if isinstance(r, Tag)]))
    if stop is not None:
        moved = _move_into(frag, insert_before=stop)
    else:
        moved = _move_into(frag, parent=target.parent)
    for r in removed:
        r.extract()
    record_placement(state, target.parent, moved[0] if moved else None)
    return len(removed)


# ── [ ... ] placeholders ─────────────────────────────────────────────────────
def brackets_in(t):
    return [m.group(0) for m in _BRACKET_RE.finditer(t or "")]


def bracket_nodes(root):
    out = []
    for n in root.descendants:
        if isinstance(n, NavigableString) and n.parent is not None and brackets_in(str(n)):
            out.append(n)
    return out


def find_bracket(root, keys):
    for node in bracket_nodes(root):
        for b in brackets_in(str(node)):
            t = norm(b)
            for k in keys:
                if norm(k) and norm(k) in t:
                    return {"node": node, "text": b}
    return None


def is_filler_bracket(b):
    if AUTHOR_FILLER.match(b):
        return True
    if re.search(r"here\s*[.!?]?\s*\]$", b, re.I):
        return True
    inner = re.sub(r"^\[|\]$", "", b).strip()
    if re.match(r"^[#xX?_\-\s]+$", inner):
        return True
    return False


def only_brackets(el):
    t = (el.get_text() or "").replace("\u00a0", " ")
    return norm(_BRACKET_RE.sub("", t)) == ""


def _replace_in_node(node, old, new):
    """Swap text inside one text node, keeping it in place."""
    fresh = NavigableString(str(node).replace(old, new))
    node.replace_with(fresh)
    return fresh


# ── banner ───────────────────────────────────────────────────────────────────
def _is_banner_box(el):
    st = ((el.get("style") or "") if isinstance(el, Tag) else "").lower().replace(" ", "")
    return "background-color:#b20000" in st or "background-color:rgb(178,0,0)" in st


def banner_target(root, used_targets, section_labels):
    hs = _headings(root)
    colored = plain = None
    has_box = False
    for h in hs:
        t = label(h.get_text())
        if not t or STRUCTURAL.match(t):
            continue
        in_cell, box, up = False, None, h.parent
        while up is not None and up is not root:
            if isinstance(up, Tag) and up.name in CELL:
                in_cell = True
                break
            if box is None and _is_banner_box(up):
                box = up
            up = up.parent
        if box is not None:
            has_box = True
        if in_cell:
            continue
        # Already carrying content, so it is a section heading, not the banner
        if any(h is u for u in used_targets):
            continue
        # Never rename a heading that labels a fixed part of the template
        if any(t == label(s) for s in section_labels):
            continue
        if box is not None:
            colored = colored or h
        # Only the page's very first heading can stand in for a banner. Any
        # later one labels a section of the layout, and renaming it destroys
        # part of the template.
        elif plain is None and hs and h is hs[0]:
            plain = h
    # Once the template has a coloured banner, only that banner is fair game.
    return colored or (None if has_box else plain)


# ── the merge itself ─────────────────────────────────────────────────────────
def merge_template(raw_html, cfg):
    """
    Fill a template. Mirrors the JS return shape:
      {ok, html, notes, warnings, filled, removed, kept, dropped, unwrapped}
    """
    root = _wrap(raw_html)
    notes, warnings, filled = [], [], []
    leftovers, kept, dropped, unwrapped, used_targets = [], [], [], [], []
    section_labels = cfg.get("sectionLabels") or []
    # Sections are not always placed in the order the doc had them: one that
    # matches a template heading is filled first, and one with no heading joins
    # a box afterwards. Each placement is recorded with the position its section
    # held in the doc, so a later arrival is still slotted into the right spot.
    state = {"placements": [], "donors": [], "last_box": None,
             "order": 0, "label": ""}

    # ---- 1. Fill each region from its own heading ---------------------------
    pending = []
    for i, r in enumerate(cfg.get("regions") or []):
        if not r.get("html") and not r.get("text"):
            continue
        r["__order"] = r["order"] if isinstance(r.get("order"), int) else i
        state["order"] = r["__order"]
        state["label"] = r.get("label") or ""
        target = find_label(root, r.get("names") or [])
        if target is not None and not r.get("inline"):
            count = replace_under(target, r["html"], root, section_labels, state)
            if count >= 0:
                used_targets.append(target)
                filled.append(r["key"])
                notes.append(f'"{norm(target.get_text())}" filled from the doc '
                             f'({count} template block(s) replaced).')
                continue
        pending.append(r)

    # ---- 2. Anything still unfilled: use its [ ... ] placeholder ------------
    for i, r in enumerate(pending):
        state["order"] = r["__order"] if isinstance(r.get("__order"), int) else i
        state["label"] = r.get("label") or ""
        hit = find_bracket(root, r.get("brackets") or [])
        if not hit:
            # No heading and no [ ] spot of its own. The content still belongs
            # inside the layout, so it goes into the box the section above it
            # filled rather than being dropped or tacked on to the end.
            where = None if r.get("inline") else place_loose(r["html"], root, state,
                                                             section_labels)
            if where:
                filled.append(r["key"])
                notes.append(
                    f'The template has no "{r["key"]}" heading, so that content '
                    + ("went into the same box as the section above it."
                       if where == "box" else
                       "went into the placeholder box in the template."))
            else:
                warnings.append(f'Nowhere in the template to put the "{r["key"]}" '
                                f'content (no heading and no matching [ ] placeholder).')
            continue
        if r.get("inline"):
            plain = re.sub(r"<[^>]+>", "", r.get("text") or "").replace("&amp;", "&")
            _replace_in_node(hit["node"], hit["text"], plain)
            filled.append(r["key"])
            notes.append(f'Placeholder {hit["text"]} filled in place.')
            continue
        parent_el = hit["node"].parent
        block = block_of(parent_el, root) if parent_el is not None else None
        if block is None or block.parent is None:
            warnings.append(f'Could not replace the {hit["text"]} placeholder.')
            continue
        frag = _wrap(r["html"])
        _style_frag(frag, _donors_from([block]))
        host = block.parent
        moved = _move_into(frag, insert_before=block)
        block.extract()
        record_placement(state, host, moved[0] if moved else None)
        filled.append(r["key"])
        notes.append(f'Placeholder {hit["text"]} replaced with the "{r["key"]}" content.')

    # ---- 2b. Banner: its heading text and the copy directly under it --------
    banner_used = False
    banner = cfg.get("banner")
    if banner and (banner.get("heading") or banner.get("html")):
        state["order"] = banner["order"] if isinstance(banner.get("order"), int) else 9999
        state["label"] = banner.get("heading") or ""
        bt = banner_target(root, used_targets, section_labels)
        if bt is not None:
            if banner.get("html"):
                bn = replace_under(bt, banner["html"], root, section_labels, state)
                if bn >= 0:
                    notes.append(f"Banner copy replaced ({bn} template block(s)).")
            if banner.get("heading"):
                _set_deep_text(bt, banner["heading"])
            used_targets.append(bt)
            filled.append(banner.get("key") or "Banner")
            banner_used = True
            notes.append(f'Banner heading set to "{banner.get("heading") or ""}".')
        elif banner.get("html") or banner.get("fallbackHtml"):
            # No banner heading was free. Put the content in the layout's own
            # content box before ever appending a new banner to the bottom of
            # the page, which is where this used to land.
            b_where = (place_loose(banner["html"], root, state, section_labels)
                       if banner.get("html") else None)
            if b_where:
                filled.append(banner.get("key") or "Banner")
                notes.append(f'No banner heading was free, so the '
                             f'"{banner.get("key") or "banner"}" content went into the '
                             f'template box, not the bottom of the page.')
            elif banner.get("fallbackHtml"):
                frag = _wrap(banner["fallbackHtml"])
                _move_into(frag, parent=root)
                filled.append(banner.get("key") or "Banner")
                warnings.append(f'The banner was already taken, so the '
                                f'"{banner.get("key") or "banner"}" section was added as '
                                f'a banner of its own at the end.')
        else:
            warnings.append(f'No banner heading in the template for the '
                            f'"{banner.get("key") or "banner"}" section.')

    # ---- 2c. Put each section's own header back above it --------------------
    restore_headers(root, state)

    # ---- 3. Discussion title -----------------------------------------------
    title_done = False
    if cfg.get("title"):
        th = find_bracket(root, cfg.get("titleBrackets") or ["title"])
        if th:
            _replace_in_node(th["node"], th["text"], cfg["title"])
            notes.append(f'Title placeholder {th["text"]} set to "{cfg["title"]}".')
            title_done = True
    if not title_done and not banner_used and cfg.get("updateBanner") and cfg.get("title"):
        # The same search the banner step uses, so the title can only ever land
        # on the template's banner heading. When the banner already carries a
        # real label of its own, nothing is renamed and the title stays in the
        # Canvas name field, which is better than overwriting the layout.
        bt = banner_target(root, used_targets, section_labels)
        if bt is not None:
            _set_deep_text(bt, cfg["title"])
            notes.append(f'Banner heading renamed to "{cfg["title"]}".')
            title_done = True
        else:
            warnings.append("No banner heading or [title] placeholder found, so the "
                            "title was only set in the Canvas name field.")

    # ---- 4. Deal with the [ ... ] text the doc never filled -----------------
    # By default nothing here is edited, only reported, because template lines
    # like "[insert number] scholarly resources" are filled in by hand after the
    # build. With the clean-up option on, author instructions take their whole
    # line while bracketed finished content keeps its wording and loses the
    # brackets.
    seen = []
    for node in bracket_nodes(root):
        if node.parent is None:
            continue
        parent_el = node.parent
        bs = brackets_in(str(node))
        fillers = []
        if not cfg.get("dropLeftovers"):
            kept.extend(bs)
            continue
        cur = node
        for b in bs:
            if is_filler_bracket(b):
                fillers.append(b)
            else:
                cur = _replace_in_node(cur, b, b[1:-1])
                unwrapped.append(b)
        if not fillers:
            continue
        blk = block_of(parent_el, root)
        if blk is None or blk is root or blk.parent is None:
            kept.extend(fillers)
            continue
        leftovers.extend(fillers)
        if any(blk is s for s in seen):
            continue
        seen.append(blk)
        if not only_brackets(blk):
            dropped.append(norm(blk.get_text())[:120])
        par = blk.parent
        blk.extract()
        prune_empty(par, root)

    if not filled:
        return {"ok": False, "reason": "no template region matched the doc content",
                "notes": notes, "warnings": warnings, "removed": leftovers,
                "kept": kept, "dropped": dropped, "unwrapped": unwrapped}
    return {"ok": True, "html": root.decode_contents(), "notes": notes,
            "warnings": warnings, "filled": filled, "removed": leftovers,
            "kept": kept, "dropped": dropped, "unwrapped": unwrapped}


def swap_prompt_box(raw_html, new_inner):
    """
    The old single-box path, for templates with no headings at all: find the
    smallest element that still holds every Lorem Ipsum placeholder and replace
    its contents, keeping the box's own paragraph and list styling.
    """
    root = _wrap(raw_html)
    all_els = root.find_all(True)

    def depth(el):
        n, cur = 0, el
        while cur is not None and cur is not root:
            n += 1
            cur = cur.parent
        return n

    def ph_count(el):
        return len(re.findall(r"lorem\s+ipsum", el.get_text() or "", re.I))

    def deepest(lst):
        return max(lst, key=depth) if lst else None

    total = ph_count(root)
    if total > 0:
        box = deepest([el for el in all_els if ph_count(el) == total])
        leaf = {"p", "li", "span", "em", "strong", "b", "i", "u"}
        while (box is not None and box.parent is not None and box.parent is not root
               and box.name in leaf):
            box = box.parent
    else:
        # No placeholder text — fall back to the innermost background-styled block
        box = deepest([el for el in all_els if "background" in (el.get("style") or "").lower()])

    if box is None:
        return {"ok": False, "reason": "prompt box not found"}

    donors = {
        "p": _attrs_of(box.find("p")),
        "li": _attrs_of(box.find("li")),
        "ul": _attrs_of(box.find("ul")) or _attrs_of(box.find("ol")),
        "ol": _attrs_of(box.find("ol")) or _attrs_of(box.find("ul")),
    }
    frag = _wrap(new_inner)
    _style_frag(frag, donors)
    box.clear()
    _move_into(frag, parent=box)
    return {"ok": True, "html": root.decode_contents()}


def compose_body(template_html, content_html, title="", anchor="Instructions",
                 update_banner=True, sections=None, drop_leftovers=False):
    """
    Build the finished discussion body from the template.

    Returns (body_html_or_None, log_lines). The log lines are the same messages
    the desktop tool wrote into its run window.
    """
    log = []
    sections = sections or {}
    order_map = sections.get("_order") or {}
    label_map = sections.get("_labels") or {}

    def _ord(key, default):
        value = order_map.get(key)
        return value if isinstance(value, int) else default

    prompt = (sections.get("prompt") or "").strip()
    objective = (sections.get("objective") or "").strip()
    instructions = (sections.get("instructions") or "").strip()
    banner = sections.get("_banner") or None

    if not (prompt or objective or instructions or banner):
        instructions = content_html

    regions = []
    if prompt:
        regions.append({
            "key": "Discussion Prompt",
            "names": ["Discussion Prompt", "Prompt", "Scenario", "Overview",
                      "Background", "Case Study"],
            "brackets": ["prompt", "scenario", "overview", "background"],
            "html": prompt,
            "order": _ord("prompt", 0),
            "label": label_map.get("prompt", ""),
        })
    if objective:
        regions.append({
            "key": "Objective",
            "names": ["Objective", "Objectives", "Purpose"],
            "brackets": ["objective", "purpose"],
            "html": objective,
            "order": _ord("objective", 1),
            "label": label_map.get("objective", ""),
        })
    if instructions:
        names = [n for n in [anchor, "Discussion Instructions", "Instructions",
                             "Directions", "Questions", "Task", "Requirements"] if n]
        regions.append({
            "key": "Instructions",
            "names": names,
            "brackets": ["requirement", "instruction", "question",
                         "insert instructions"],
            "html": instructions,
            "order": _ord("instructions", 2),
            "label": label_map.get("instructions", ""),
        })

    # Sections the doc names itself: "Response Requirements" routes to the
    # template heading of that name, "[Prompt]" to its [ ... ] spot.
    for _ei, ex in enumerate(sections.get("_regions") or []):
        if not (ex.get("html") or "").strip():
            continue
        regions.append({
            "key": ex["label"],
            "names": [] if ex.get("bracket") else [ex["label"], ex["keyword"]],
            "brackets": [ex["keyword"]],
            "html": ex["html"],
            "order": ex["order"] if isinstance(ex.get("order"), int) else 3 + _ei,
            "label": ex.get("display") or ex["label"],
        })

    # Doc order decides where each section ends up in the box, so the list the
    # merge works through is sorted that way too.
    regions.sort(key=lambda reg: reg.get("order", 0))

    banner_cfg = None
    if banner and ((banner.get("label") or "").strip() or (banner.get("html") or "").strip()):
        banner_cfg = {
            "order": banner["order"] if isinstance(banner.get("order"), int)
            else _ord("_banner", -1),
            "key": (banner.get("label") or "Banner").strip(),
            "heading": (banner.get("label") or "").strip(),
            "html": (banner.get("html") or "").strip(),
            # Last resort if the banner turns out to be spoken for: a copy of it
            # is appended so the content is never silently dropped.
            "fallbackHtml": (
                banner_heading_html(template_html, banner.get("label") or "")
                + "\n" + (banner.get("html") or "")
            ),
        }

    cfg = {
        "regions": regions,
        "banner": banner_cfg,
        "title": title or "",
        "titleBrackets": ["title", "discussion name", "assignment name"],
        "updateBanner": bool(update_banner) and banner_cfg is None,
        "dropLeftovers": bool(drop_leftovers),
        "sectionLabels": TEMPLATE_SECTION_LABELS,
    }

    result = merge_template(template_html, cfg)
    if result.get("ok"):
        for note in result.get("notes") or []:
            log.append(f"   ⚙️ {note}")
        for warn in result.get("warnings") or []:
            if "Nowhere in the template" in warn:
                continue
            log.append(f"   ⚠️ {warn}")

        placed = set(result.get("filled") or [])
        for r in regions:
            if r["key"] in placed:
                continue
            log.append(f'   📋 No spot in the template for "{r["key"]}". '
                       f'Add a heading with that name, or copy this in by hand:')
            for line in html_to_lines(r["html"]):
                log.append(f"        {line}")

        unwrapped = list(dict.fromkeys(result.get("unwrapped") or []))
        if unwrapped:
            log.append(f"   ✏️ Brackets stripped, wording kept: {', '.join(unwrapped)[:300]}")
        removed = list(dict.fromkeys(result.get("removed") or []))
        if removed:
            log.append(f"   🧹 Empty template filler removed: {', '.join(removed)[:300]}")
        for line in list(dict.fromkeys(result.get("dropped") or [])):
            log.append(f'   ✂️ Whole line removed, the doc had nothing for it: "{line}"')
        kept = list(dict.fromkeys(result.get("kept") or []))
        if kept:
            log.append(f"   📝 Left for you to fill in by hand: {', '.join(kept)[:300]}")
        return result.get("html"), log

    log.append(f'   ℹ️ No template heading matched ({result.get("reason", "unknown")}) '
               f'— using the prompt box instead.')

    # Older templates have one shaded prompt box and no headings at all
    fallback = swap_prompt_box(template_html, content_html)
    if fallback.get("ok"):
        return fallback.get("html"), log
    log.append(f'   ❌ Could not locate the prompt box in the template '
               f'({fallback.get("reason", "unknown")}).')
    return None, log
