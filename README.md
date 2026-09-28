# Diya Bridal Studio & Academy CRM

Static frontend (GitHub Pages) + Google Apps Script backend + Google Sheets/Drive database.

## Step 1: Backend (Google Apps Script)
1. Go to https://script.google.com > New project.
2. Delete the default code and paste all of `Code.gs`.
3. Select `setup` in the function dropdown > Run > allow permissions. This creates the spreadsheet and Drive folders.
4. Deploy > New deployment > type **Web app**. Execute as: **Me**. Who has access: **Anyone**. Deploy.
5. Copy the Web app URL (ends with `/exec`).

## Step 2: Connect the frontend
Open `index.html`, find `API_URL: ""` near the top of the script, and paste the URL between the quotes.

## Step 3: Publish on GitHub Pages
1. Create a new GitHub repository (e.g. `diya-crm`).
2. Upload `index.html`, `.nojekyll` and `README.md` (Code.gs is optional; it is a backup copy).
3. Repository > Settings > Pages > Source: **Deploy from a branch** > Branch: `main` / root > Save.
4. After about a minute your site is live at `https://YOUR-USERNAME.github.io/diya-crm/`.

Login: `admin` / `admin123`. Change this right away (edit the Users sheet: passHash is SHA-256 base64).

## Updating
- Frontend change: edit `index.html` on GitHub and commit.
- Backend change: edit in Apps Script, then Deploy > Manage deployments > Edit > New version.
- Leave `API_URL` empty to run in demo mode (nothing is saved).

## Note
The repository is public on the free plan. Do not put real customer data or passwords in it. Data lives in your Google Sheet, not in GitHub.
