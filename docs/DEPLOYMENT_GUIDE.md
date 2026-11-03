# Deployment guide — Streamlit Community Cloud from GitHub

## 1. Push the repository to GitHub
1. Create a **private** repository (for example `bess-auxiliary-engineering-app`).
2. Push the contents of this folder. `.gitignore` already excludes `.streamlit/secrets.toml` and saved project JSON files.
   **Never commit real secrets or project data.**

```bash
git init && git add . && git commit -m "BESS auxiliary engineering app v1.0.0"
git branch -M main
git remote add origin https://github.com/<org>/bess-auxiliary-engineering-app.git
git push -u origin main
```

## 2. Create the app
1. Go to share.streamlit.io and sign in with GitHub.
2. Choose **Create app** → *Deploy a public app from GitHub*, then select the repository, branch `main` and main file `app.py`.
3. Under **Advanced settings**:
   - Choose **Python 3.12**.
   - Paste the secrets (see step 3).
4. Deploy. The dependencies in `requirements.txt` are installed automatically. LibreOffice is **not** needed at runtime.

## 3. Secrets (App → Settings → Secrets)
```toml
[access]
allowed_emails = ["engineer1@company.com", "engineer2@company.com"]
session_timeout_minutes = 60

[branding]
company_name = "SunStripe"
logo_path = "assets/logo.svg"
```
- **Updating access:** editing the secrets updates the allowlist without redeploying. Addresses are compared after trimming and lower-casing.
- **Missing or empty `[access]`:** the login page blocks every sign-in and shows a configuration message.

## 4. Access-control disclosure (mandatory reading)
The email allowlist gate is **not identity authentication**. It does not prove that the person typing an address owns it. Anyone who
knows an approved address can enter.

Built-in mitigations:
- Generic denial messages; the allowlist is never shown.
- A 60-second lockout after 5 failed attempts in a session.
- An inactivity timeout.
- Logout clears the session.

Further steps you can take:
- Make the Streamlit Cloud app **private** and restrict viewers in the Cloud sharing settings. This adds account-based access.
- Treat the app as suitable only for controlled preliminary engineering use.
- For stronger access control, a later release can switch to Streamlit's native OIDC login (Microsoft Entra ID or Google). That change affects only `authentication/` and `app.py`.

## 5. Data protection
- Projects live only in each user's session (`st.session_state`, isolated per browser session) and in the JSON files they download.
- Nothing is written to a database or the server disk, and no shared cache (`st.cache_data`) holds project data.
- `client.showErrorDetails = false` and the log level is `warning`. The application does not log project content.
- Streamlit Community Cloud restarts apps and does not keep local files. Users must download their project JSON before leaving.

## 6. Post-deployment checks (to be performed and recorded by the deploying engineer)
1. An unapproved address is rejected with a generic message; an approved address in mixed case with spaces is accepted.
2. Logout returns to the login page. Opening a page URL directly without signing in shows "not authorised".
3. Create a project, save the JSON, start a new session, upload it, and confirm the inputs and results are identical.
4. Import `tests/reference/BASELINE_v1.1_worked_example.xlsx` and compare the peak demand with the worked example (≈ 3,850.9 kVA, MAXSIM).
5. Download the PDF, the Excel results and the handoff JSON (the handoff page shows "validates against schema").
6. Open the exported v1.1 workbook in **Microsoft Excel**, recalculate (Ctrl+Alt+F9) and confirm there is no repair prompt. This also starts closing dependency D-1.

## 7. Updating
Push to `main` and the app redeploys automatically. Bump `APP_VERSION` in `data_models/project_schema.py` and
`ENGINE_VERSION` in `calculations/constants.py` when the methodology changes. Run `python -m pytest` before every release.
