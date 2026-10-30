# Deploy the Method Statement Builder — Step by Step

This is a generic Streamlit app for building and infrastructure activities. The user enters an activity name; the app uses the Groq API to draft a detailed method statement and offers a Word download. It includes a sequential work method, two generated conceptual sketches, activity-relevant ITP, field checklist, safety risk assessment, environmental and emergency controls, and review fields. It does not read or follow project specifications or drawings. A competent engineering, QA/QC and HSE team must review every draft before use.

## Files already prepared

Use the files in this folder. The full application source is [app.py](app.py); you do not need to write that long file from scratch.

- `app.py` — complete Streamlit app and Word report generator.
- `requirements.txt` — packages needed by the app.
- `.gitignore` — keeps local secrets and Python environment files out of GitHub.
- `.streamlit/secrets.toml.example` — sample API key setting. Replace with the real key only in the deployment dashboard or your ignored local secrets file.

## Step 1 — Get a Groq API key

1. Open [Groq Console API Keys](https://console.groq.com/keys) and sign in.
2. Create a secret key. Copy it when it is displayed and store it privately; the complete key is only shown at creation.
3. API requests may incur charges. Check the API account's billing and usage settings.

Never paste the key into `app.py`, a public GitHub file, a chat, or a screenshot. Keep it in Streamlit Secrets or the ignored local secrets file. The app uses Groq's OpenAI-compatible API endpoint with the Groq base URL and API key, as described in [Groq's compatibility guide](https://console.groq.com/docs/openai).

## Step 2 — Create the GitHub repository

1. Sign in at [GitHub](https://github.com).
2. Select **New repository**.
3. Use this repository name: `infrastructure-method-statement-builder`.
4. Choose **Private** or **Public**. This app has no project files, but Private limits who can view the source.
5. Do not initialize it with a README, license, or `.gitignore` if you plan to push the prepared folder using the commands below.
6. Select **Create repository** and leave that page open; GitHub shows your repository URL there.

## Step 3 — Push the prepared app files to GitHub

On Windows, open PowerShell and copy/paste these commands one at a time. The folder path below is where this prepared app was saved:

```powershell
cd "C:\Users\waqas\Documents\Codex\2026-10-30\while-working-on-infrastructure-and-buildings\outputs\infrastructure-method-statement-app"
git init
git add app.py requirements.txt README.md .gitignore .streamlit/secrets.toml.example
git commit -m "Add generic method statement builder"
git branch -M main
git remote add origin https://github.com/YOUR-GITHUB-USERNAME/infrastructure-method-statement-builder.git
git push -u origin main
```

Before running the `git remote` line, replace `YOUR-GITHUB-USERNAME` with the username shown in your GitHub profile. If your repository belongs to an organization, use the organization name instead. For example:

```powershell
git remote add origin https://github.com/acme-construction/infrastructure-method-statement-builder.git
```

Git may open a browser sign-in window. Complete the GitHub sign-in and return to PowerShell. If PowerShell says `git` is not recognized, install Git for Windows, reopen PowerShell, and repeat.

## Step 4 — Create the Streamlit app

1. Open [Streamlit Community Cloud](https://share.streamlit.io) and sign in with the GitHub account that owns or can access the repository.
2. Select **Create app**.
3. Select the repository `infrastructure-method-statement-builder`.
4. Select branch `main`.
5. Set the app file path to `app.py`.
6. Open **Advanced settings** (or the **Secrets** field in the deployment form).
7. Paste this into **Secrets**, replacing the quoted example with your real API key:

   ```toml
   GROQ_API_KEY = "paste-your-real-groq-key-here"
   ```

8. Select **Save**, then **Deploy**. Streamlit installs the libraries listed in `requirements.txt` and builds the app. The official deployment steps are in [Streamlit's Community Cloud guide](https://docs.streamlit.io/deploy/streamlit-community-cloud/deploy-your-app/deploy).

## Step 5 — Use the deployed app

1. Wait for Streamlit to finish deploying and open the app URL it provides.
2. Enter one activity, for example `Concrete block masonry wall construction`.
3. Select **Generate method statement**.
4. Review the draft and the information gaps shown on screen.
5. Select **Download Word method statement**.
6. Have the competent project team review, complete, and approve the document before work begins.

## Optional — Run the app on your computer first

Open PowerShell and run:

```powershell
cd "C:\Users\waqas\Documents\Codex\2026-10-30\while-working-on-infrastructure-and-buildings\outputs\infrastructure-method-statement-app"
py -m venv .venv
.\.venv\Scripts\Activate.ps1
python -m pip install -r requirements.txt
Copy-Item .streamlit\secrets.toml.example .streamlit\secrets.toml
notepad .streamlit\secrets.toml
```

In Notepad, replace the example value with your Groq key and save. Then run:

```powershell
streamlit run app.py
```

This opens the app in a browser on your computer. The real `.streamlit/secrets.toml` is excluded by `.gitignore`; never add it to GitHub.

## Optional — Copy-ready dependency file

If you need to recreate `requirements.txt`, its full contents are:

```text
streamlit>=1.39,<2
openai>=1.60,<3
python-docx>=1.1,<2
matplotlib>=3.9,<4
```

## Update the app later

After editing `app.py` or `README.md`, open PowerShell in the app folder and run:

```powershell
git add app.py README.md requirements.txt
git commit -m "Update method statement builder"
git push
```

Streamlit Community Cloud rebuilds the deployed app from the updated GitHub files. To change the API key, update the app's **Settings → Secrets** in Streamlit; do not edit the key into the code.

## Common deployment errors

- **App cannot find a package:** confirm `requirements.txt` is in the repository root and wait for Streamlit to finish reinstalling dependencies.
- **GROQ_API_KEY missing:** add the TOML setting under Streamlit app Secrets, then reboot or redeploy.
- **Incorrect API key:** create a replacement key and update Streamlit Secrets.
- **GitHub authentication failed:** sign in to GitHub in the browser and confirm the selected account has access to the repository.
- **Generation fails after deploy:** view the Streamlit app logs and check the API key, model access and API account usage.

## Important limits

- Output is a generic AI draft based only on the activity name. It does not use project specifications, drawings, or site files.
- The sketch is an indicative process flow, not an engineered or scale drawing.
- Competent personnel must review site-specific methods, acceptance criteria, hold points, frequencies, permits, risk ratings and controls.
- The default Groq model is `openai/gpt-oss-20b`. The app requests JSON mode, normalizes common array wrappers, allows up to 8,192 completion tokens, and uses low reasoning effort. To select a different compatible model, configure `GROQ_MODEL` in the runtime environment. See [Groq's JSON mode and structured output documentation](https://console.groq.com/docs/structured-outputs) and [API parameter reference](https://console.groq.com/docs/api-reference).

The expanded detail level requires `matplotlib` for the two Word sketches. Make sure the updated `requirements.txt` is pushed with the app; Streamlit will reinstall dependencies during its next build.
