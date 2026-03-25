# Opalion Delivery Note Converter

This repository contains a Streamlit app for converting Opalion delivery notes into one combined order import CSV.

## What it does
- Accepts multiple `.xlsx` delivery notes
- Extracts the order lines automatically
- Applies the required rules:
  - `Owner = Opalion`
  - `Order Type = Sales`
  - product codes with a leading `0` get an apostrophe
- Shows a preview
- Downloads one combined CSV

## Online deployment
The app is designed to be deployed from GitHub to Streamlit Community Cloud.

### Deploy steps
1. Push this repository to GitHub.
2. Create a Streamlit Community Cloud account and connect GitHub.
3. In Streamlit Community Cloud, create a new app and point it at `streamlit_app.py`.
4. After that, updates pushed to GitHub will appear in the deployed app.

## GitHub release
GitHub Releases are a good way to package software and downloadable files for others to use. This repo includes a release workflow so you can attach a zip of the source code to a tagged release.

## Local run
```bash
pip install -r requirements.txt
streamlit run streamlit_app.py
```
