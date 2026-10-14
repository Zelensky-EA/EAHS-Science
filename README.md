# EAHS Science Department web apps

This folder (GitHub repo `Zelensky-EA/EAHS-Science`) is the **single source of truth** for three things:

| What | Where in this folder | Where it runs | Link people use |
|---|---|---|---|
| Public website (home redirect, Room 203 reservations page) | top level: `index.html`, `inventory.html`, `reservations.html`, `lab-checkout.html`, `styles.css`, `hub.css`, `app.js`, `config.js` | GitHub Pages | `https://zelensky-ea.github.io/EAHS-Science/…` |
| Inventory + Department Hub (Apps Script) | `apps-script/inventory-hub/` | Google Apps Script | `…/AKfycbx5wsf3…/exec` (inventory) and `…/exec?page=hub` (hub) |
| Room 203 booking backend (Apps Script) | `apps-script/lab-reservations/` | Google Apps Script | `…/AKfycbx-a_9B…/exec` (set in `config.js`) |

`_archive/` holds old copies kept for reference only. **Never edit or deploy from `_archive/`.**

## Rules

1. **Change code here, then send it to Google.** Do not edit code in the Apps Script editor, except to deploy. If you ever do, run `clasp pull` right after so this folder catches up.
2. **Always `clasp pull` before the first `clasp push` on any computer.** `clasp push` replaces the *entire* Apps Script project with this folder's files. A file missing here (for example `Catalog.js`) is **deleted from the live project**.
3. **Deploy as a new version of the existing deployment** (Deploy → Manage deployments → ✏️ Edit → New version). Never create a new deployment, because that changes the /exec link.
4. **Commit to git after every deploy**, with a message saying what changed. The commit history is your record of what is live.
5. **Edit spreadsheets outside school hours.** Do not sort, insert, or delete rows in the live inventory Sheet while teachers are using it. Use filter views to look at the data.

## One-time setup on your Mac (about 15 minutes)

1. **Node.js:** open Terminal and run `node -v`. If you see "command not found," install the LTS version from nodejs.org.
2. **clasp** (Google's command-line tool for Apps Script): `npm install -g @google/clasp`
3. **Turn on the Apps Script API** for your account: open https://script.google.com/home/usersettings and switch it **On**.
   If the district blocks this setting, clasp cannot connect. In that case keep using copy-paste, but follow rules 3–5 anyway.
4. `clasp login`, then sign in with your school account.
5. **Fill in the Script IDs.** In each Apps Script project, open ⚙️ Project Settings → IDs → **Script ID**, and paste it into that project's `.clasp.json`, replacing `PASTE_SCRIPT_ID_HERE`:
   - `apps-script/inventory-hub/.clasp.json` → the inventory/hub project
   - `apps-script/lab-reservations/.clasp.json` → the Room 203 booking project
6. **Pull the live code down** (done 2026-10-13). Note: clasp saves Apps Script `.gs` files as `.js` here; that is normal, and they become `.gs` again in the editor:
   ```
   cd ~/"EAHS WebPages/EAHS-203/apps-script/inventory-hub"   && clasp pull
   cd ~/"EAHS WebPages/EAHS-203/apps-script/lab-reservations" && clasp pull
   ```
   Then in the repo folder run `git status` and `git diff` to see exactly how the folder differed from what is live. Commit the result: this is your first verified snapshot.

## Making a change (every time)

1. `cd` into the project folder and run `clasp pull`. This picks up anything changed in the editor.
2. Edit the files here.
3. `clasp push`, which uploads to the Apps Script project. The live /exec link does **not** change yet.
4. Test: Apps Script editor → Deploy → **Test deployments**. Open the `/dev` link and try the checklist in `apps-script/inventory-hub/CHANGES-2026-10-13.md`.
5. Deploy: Deploy → Manage deployments → ✏️ Edit → **New version** → Deploy.
6. Commit: `git add -A && git commit -m "Describe the change" && git push`.

**Rollback:** Deploy → Manage deployments → ✏️ Edit → choose the previous version → Deploy. To roll back the files, `git log` to find the old commit, check out its files, then `clasp push`.

## Website (GitHub Pages)

Pushing to `main` publishes **only** the files listed in `.github/workflows/static.yml`. If you add a page, stylesheet, or script to the website, add its name to the `cp` line there.

Note: the published *site* no longer includes Apps Script code or CSVs. If the GitHub *repository* itself is public, its files can still be read on github.com. Check under Settings → General → Danger Zone → visibility. A private repo needs a paid GitHub plan to keep using Pages.
