# Hmoob Keeb Kwm — local setup

This folder holds the source for the Apps Script web app. Claude writes the three
source files here directly; `clasp` pushes them to Google; `git` keeps the history.

**Files pushed to Apps Script:** `Code.gs`, `Index.html`, `Styles.html`, `appsscript.json`
Everything else here (this file, git, notes) stays on your PC — see `.claspignore`.

---

## One-time setup

### 1. Install clasp

```
npm install -g @google/clasp
```

Current version is v3. If you have an old v2 installed, the same command upgrades it.

### 2. Turn on the Apps Script API

Open <https://script.google.com/home/usersettings> and switch **Google Apps Script API**
to **On**. Nothing works without this.

### 3. Log in

```
clasp login
```

A browser window opens — sign in with the Google account that owns the spreadsheet.

### 4. Fill in your Script ID

Open `.clasp.json` in this folder and replace `PASTE_YOUR_SCRIPT_ID_HERE`.

To find it: open the Apps Script editor → **⚙ Project Settings** → copy **Script ID**.
(It is *not* the `AKfycb...` string in the web-app URL — that's the deployment ID.)

### 5. Pull once, to fetch your real manifest

```
cd "D:\Dev\Hmoob Keeb Kwm"
clasp pull
```

This downloads `appsscript.json` — your project's real settings, including who the web
app runs as and who may access it.

> **Why pull instead of letting Claude write it:** a wrong `appsscript.json` would be
> pushed over your live one and could change your web app's permissions or break the
> deployment. Better to take the real one from Google than to guess it.

`clasp pull` also overwrites `Code.gs` / `Index.html` / `Styles.html` with whatever is
currently live — which may be older than what's in this folder. That's fine: tell Claude
"pull is done" and it will write the current versions back over them before you push.

### 6. Start the git history

```
git init
git add .
git commit -m "Initial import of Hmoob Keeb Kwm"
```

---

## Everyday use

After Claude makes a change (the files here update automatically):

```
cd "D:\Dev\Hmoob Keeb Kwm"
clasp push
git add . && git commit -m "what changed"
```

Then reload the web app. No copy-pasting.

| Command | What it does |
|---|---|
| `clasp push` | Send this folder's files up to Apps Script |
| `clasp pull` | Bring Google's version down (overwrites local files) |
| `clasp show-file-status` | List which files would be pushed |
| `git log --oneline` | See the change history |
| `git diff` | See what changed since the last commit |

### Two things to watch

- **`clasp push` deletes remote files that don't exist locally.** Your project should
  contain exactly `Code.gs`, `Index.html`, `Styles.html` and the manifest. If you ever add
  a file in the Apps Script editor, `clasp pull` it down before your next push, or the push
  will remove it.
- **Never commit `.clasprc.json`** — it holds your Google login token. `.gitignore`
  already excludes it.

---

## Rolling back

```
git log --oneline           # find the commit you want
git checkout <commit> -- .  # restore those file versions
clasp push                  # send them live
```
