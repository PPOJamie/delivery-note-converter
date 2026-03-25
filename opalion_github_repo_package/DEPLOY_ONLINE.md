# Deploying the app online

## Option 1: Streamlit Community Cloud
1. Push this repo to GitHub.
2. Sign in to Streamlit Community Cloud and connect your GitHub account.
3. Create a new app and select:
   - Repository: this repo
   - Branch: `main`
   - Entry point: `streamlit_app.py`

Streamlit Community Cloud deploys apps from GitHub repositories, and updates show up when you push changes to GitHub. It also provides a unique `streamlit.app` URL for the deployed app. 

## Option 2: GitHub Releases
Use GitHub Releases when you want a downloadable package. Releases are based on Git tags and are intended to package software and binary files for users to download.
