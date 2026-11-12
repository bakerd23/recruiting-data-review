# Recruiting Data Review

The review site for Virginia Tech's recruiting database interns. Reviewers sign in with Google, pick
one of three tools, and decide items one at a time:

- **High schools:** which school does this spelling refer to?
- **Duplicates in one source:** are these two profiles from one site the same player?
- **Cross-source pairs:** are these two profiles from different sites the same player?

Anything a reviewer can't decide goes to **Needs review**, where a senior reviewer makes the call.

**This repository holds no player data.** It contains only the page and the Apps Script code. The review
lists live in a private Google Sheet. The script hands an item to a reviewer only after checking their
Google sign-in against the approved list. Phone numbers, emails and birth dates are never in the lists,
only whether they match.

## How it works
- One reviewer per item: an item is reserved for 30 minutes while someone has it open.
- Every decision is saved as a new row (who, when, what, note). Undo adds a row that cancels the
  previous one, so nothing is ever erased.
- Reviewers see only their own decisions.
- **Roles:**
  - `intern` reviews
  - `senior` also decides items marked Needs review
  - `admin` also manages the reviewer list, in the site under **Reviewers**
- Two-reviewer agreement is built into `Code.gs` but switched off (`REQUIRED_AGREEMENTS = 1`).

## One-time setup (owner)
1. **Sign-in client ID** (free). Open the Google Cloud console in your own project.
   - Go to **APIs & Services → OAuth consent screen**. Choose *External*, fill in the app name and your
     email, and add the reviewers as *Test users* (or publish the app later).
   - Go to **Credentials → Create credentials → OAuth client ID** and pick type **Web application**.
   - Under **Authorized JavaScript origins**, add `https://<your-github-user>.github.io` (and
     `http://localhost:8770` for local tests). Copy the **Client ID**.
2. **The Sheet.** Create a blank Google Sheet, e.g. "Recruiting Data Review", and keep it private.
3. **The script.** In the Sheet, go to **Extensions → Apps Script** and paste [`apps_script/Code.gs`](apps_script/Code.gs).
   Then open **Project Settings → Script Properties** and add:
   - `CLIENT_ID`: the client ID from step 1
   - `CEEB_FILE_ID`: the Drive file id of `ceeb.json` (upload it privately, then **Share → Copy link**;
     the id is between `/d/` and `/view`)
4. **Deploy.** Choose **Deploy → New deployment → Web app**, with *Execute as: Me* and *Who has access:
   Anyone*. Copy the URL, which ends in `/exec`. The script itself still refuses anyone who isn't signed in
   and on the list.
5. **Fill the tabs.** Run any request once (open the site), and the tabs `Reviewers`, `Queue`, `Claims`,
   `Decisions` and `AdminLog` are created.
   - In `Reviewers`, add a row for each person: `email | role | TRUE | name`. Start with yourself as `admin`.
   - In `Queue`, choose **File → Import → Upload `queue.csv` → Replace current sheet**. The file is built
     privately by the project's `export_queues.py`.
6. **Connect the page.** Put the `/exec` URL and the client ID into `config.js`, then commit and push.
7. **GitHub Pages.** Go to **Settings → Pages**, set *Deploy from a branch*, choose `main` and `/ (root)`.

If you change `Code.gs` later, use **Deploy → Manage deployments → Edit → New version** so the URL stays the same.

## Moving to Google Cloud later
The site talks to one set of actions (`me`, `next`, `release`, `decide`, `undo`, `mine`, `ceeb`, `hs_index`,
`escalated`, `resolve`, `reviewers`, `set_reviewer`). A Cloud Run service answering the same actions can
replace the Apps Script. Only `endpoint` in `config.js` changes. The Sheet tabs map one-to-one to BigQuery
tables.
