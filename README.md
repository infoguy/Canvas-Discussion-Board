# Canvas Discussion Builder

The web version of `db_creator.py`. Takes a Canvas template layout and a set of
Word docs (or typed header/content pairs), fills each region of the template
with the matching part of the content, and creates the discussions in Canvas.

Runs on Render with no Chrome, no Selenium, and no desktop install.

## What changed from the desktop script

| Desktop `db_creator.py` | This app |
| --- | --- |
| Selenium drives Chrome, logs in with a saved username and password | Canvas REST API with an access token you paste in per run |
| The merge runs as JavaScript inside the browser page | `merge.py`, the same logic ported to BeautifulSoup |
| Tkinter window | `templates/index.html` |
| Settings panel clicked field by field | One POST with the same settings as parameters |
| Template read by opening its edit page | Template read from the discussion's API record, or pasted in |
| `~/.canvas_automation_profile.json` holds credentials | Nothing is stored: the token lives only in the request |

Everything about how content is parsed and placed is unchanged: the same docx
parser, the same header routing, the same region matching, the same banner
clone, the same `[ ... ]` placeholder handling.

## Files

```
app.py                Flask routes: parse docs, preview, create
content.py            Word doc and typed-header parsing (from the desktop script)
merge.py              The template merge, ported from the browser JavaScript
canvas_api.py         Canvas REST client
templates/index.html  The form, the run log, and the preview
requirements.txt      Python dependencies
render.yaml           Render Blueprint
Procfile              Start command
runtime.txt           Python version
```

## Put it on GitHub

From the folder containing these files:

```bash
git init
git add .
git commit -m "Canvas Discussion Builder web app"
git branch -M main
git remote add origin https://github.com/infoguy/canvas-discussion-builder.git
git push -u origin main
```

If the repo already exists and has the older version in it, the same commands
work; use `git push -f origin main` only if you want to overwrite its history.

## Deploy on Render

**With the Blueprint (easiest):** in Render, choose New → Blueprint, point it at
the repo, and accept what `render.yaml` describes. Nothing else to fill in.

**Manually:** New → Web Service, connect the repo, then set

- Runtime: Python 3
- Build command: `pip install -r requirements.txt`
- Start command: `gunicorn app:app --bind 0.0.0.0:$PORT --workers 2 --threads 4 --timeout 180`
- Health check path: `/healthz`

No environment variables are required. The free plan sleeps after inactivity, so
the first request after a quiet spell takes a few seconds to wake up.

## Get a Canvas access token

In Canvas: Account → Settings → Approved Integrations → **+ New Access Token**.
Copy it when it is shown, because Canvas will not show it again. The token acts
as you, so treat it like a password. This app keeps it only for the length of
the request; it is never written to disk or into the log.

## Using it

1. Paste the token and one Canvas course URL per line.
2. Load the template layout: paste the HTML from the template discussion's HTML
   editor, load a saved `.html` file, or give a discussion URL and press
   **Load layout**.
3. Add content, either way, mixed freely:
   - **Add Word documents** for the usual `MBCS700_Week_1_Discussion.docx` files.
   - **Type headers & paste** to type a header such as "This Week's Prompt" and
     paste the content that belongs under it.
4. Any entry can build against a layout of its own: **Template for this entry…**
   opens a paste box with **Load from file**, **Use main template**, and **Save**.
   Entries with no template of their own use the main one.
5. **Check first** builds every discussion and renders it, without touching
   Canvas. Read the run log, then press **Create discussions**.

### Build options

- **Instructions heading** names the template heading the questions go under.
- **Week N Discussion:** prefixes titles that do not already carry a week.
- **Rename the banner heading** puts the title on the template's banner.
- **Remove leftover [ ] placeholders** clears author instructions the content
  never filled. Left off, they are only reported, since lines like
  "[insert number] scholarly resources" are usually finished by hand.

## Two things fixed on the way over

- **Hyperlinks in Word docs survive.** The desktop file defines `_run_to_html`
  and `_para_to_html` twice, and the second pair wins, which is the pair that
  walks `p.runs` only. Runs nested inside a `<w:hyperlink>` are not in `p.runs`,
  so every linked phrase was being dropped. Only the hyperlink-aware pair was
  carried over here.
- **Bulleted filler no longer swallows paragraphs.** When the template's
  placeholder was a sample `<ul>` and the new content was not all bullets, the
  old logic filled the `<ul>` itself, leaving `<p>` tags inside a list for
  Canvas to mangle. The list wrapper is now replaced instead.

## Known limits

- Group discussions need a group category, which the API cannot invent. The
  checkbox is honoured only when the course already has one; otherwise Canvas
  creates an ordinary discussion.
- The assignment group is matched by name. When no group of that name exists in
  a course, Canvas files the discussion under its default group and the run log
  says so.
- Uploads are capped at 25 MB per run.
