# Partner Center API Explorer

Static React + Vite website for exploring Microsoft Partner Center REST APIs from GitHub Pages.

## Features

- Office-style Entra ID sign-in with MSAL browser authentication.
- Signed-in profile display and sign-out from the Microsoft 365-style header.
- Internal build-time tenant, client ID, scopes, and Partner Center base URL configuration.
- GET, POST, PATCH, PUT, and DELETE request composer with JSON body support.
- Scenario templates based on the Microsoft Learn Partner Center API scenarios page.
- Local browser history for the last 25 actions.
- GitHub Pages-friendly static build with relative asset paths.

## Run locally

```powershell
npm install
npm run dev
```

## Build for GitHub Pages

```powershell
npm run build
```

Publish the generated `dist` folder to GitHub Pages.

## Automatic GitHub Pages deployment

This repository includes `.github/workflows/deploy-github-pages.yml`. On every push to `main`, GitHub Actions installs dependencies, builds the Vite app, uploads `dist`, and deploys it to GitHub Pages.

In the GitHub repository, go to **Settings > Pages** and set **Build and deployment > Source** to **GitHub Actions**.

Add these repository variables in **Settings > Secrets and variables > Actions > Variables**:

```text
VITE_PARTNER_CENTER_CLIENT_ID=<application-client-id>
VITE_PARTNER_CENTER_TENANT_ID=organizations
VITE_PARTNER_CENTER_SCOPE=https://api.partnercenter.microsoft.com/user_impersonation
VITE_PARTNER_CENTER_BASE_URL=https://api.partnercenter.microsoft.com/v1
```

For the Entra app registration, add the deployed GitHub Pages URL as a **Single-page application** redirect URI.

## Entra ID setup

Create or update an application registration in Microsoft Entra ID.

1. Go to **Authentication**.
2. Add a platform and choose **Single-page application**. Do not use the **Web** platform for this React app.
3. Add redirect URIs for local development and GitHub Pages:

```text
http://localhost:5173/
https://<github-user>.github.io/<repository-name>/
```

4. Grant delegated Partner Center API permissions such as:

```text
https://api.partnercenter.microsoft.com/user_impersonation
```

If you see `AADSTS9002326`, the redirect URI was added under **Web** instead of **Single-page application**, or the exact local origin such as `http://localhost:5173/` is missing.

Configure the internal app settings in `.env` before running or publishing:

```text
VITE_PARTNER_CENTER_CLIENT_ID=replace-with-application-client-id
VITE_PARTNER_CENTER_TENANT_ID=organizations
VITE_PARTNER_CENTER_SCOPE=https://api.partnercenter.microsoft.com/user_impersonation
VITE_PARTNER_CENTER_BASE_URL=https://api.partnercenter.microsoft.com/v1
```

The app stores the MSAL token cache and request history in the user's browser. If Partner Center rejects browser-origin requests because of CORS, keep this UI static and route API calls through your own authenticated backend proxy.
