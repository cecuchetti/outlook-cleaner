# Outlook Cleaner

> Automate unwanted email removal from your Outlook/Hotmail account with a secure web interface and API.

[![Python Version](https://img.shields.io/badge/python-3.7%2B-blue.svg)](https://www.python.org/downloads/)
[![License: MIT](https://img.shields.io/badge/License-MIT-yellow.svg)](https://opensource.org/licenses/MIT)

## Why This Exists

Manually unsubscribing from marketing emails is a recurring chore. Outlook Cleaner eliminates this friction by searching for specific sender names and automatically moving matched emails to the deleted folder. 

With its new web-based architecture, you can:
- **Host in the cloud**: Run it 24/7 on platforms like Render or Railway.
- **Trigger remotely**: Use the external API to start cleanups from other web apps or automation scripts.
- **Set and forget**: OAuth2 refresh tokens ensure persistent access without repeated manual logins.

## Quick Start

```bash
# 1. Clone and install
git clone https://github.com/ecuchetti/outlook-cleaner
cd outlook-cleaner
pip install -r requirements.txt

# 2. Configure environment
cp .env.example .env
# Edit .env with your Azure credentials (see Configuration)

# 3. Run locally
python server.py
```

Open [http://localhost:8000](http://localhost:8000) to login and start cleaning.

## Features

- **Premium Web Dashboard**: Manage rules and trigger cleanups via a modern Glassmorphism UI.
- **Direct API Integration**: Protected endpoint (`/api/v1/trigger-clean`) for third-party application control.
- **Secure Authentication**: Uses official Microsoft OAuth2 flow (no app passwords or IMAP secrets stored).
- **Infinite Lifecycle**: Automatically refreshes access tokens for continuous cloud operation.
- **High Performance**: Uses server-side IMAP search to process thousands of emails in seconds.

## Installation

### Prerequisites
- Python 3.7+
- A Microsoft Account (Hotmail/Outlook)
- An Azure AD Application registration (see below)

### 1. Azure App Registration
1. Go to [Azure Portal](https://portal.azure.com) -> **App registrations**.
2. Create a new **Web** registration.
3. Set **Redirect URI** to `http://localhost:8000/callback` (or your production URL).
4. Under **API permissions**, add `IMAP.AccessAsUser.All` (Microsoft Graph -> Delegated).
5. Generate a **Client Secret** under "Certificates & secrets".

### 2. Environment Variables
Create a `.env` file in the root directory:
```env
AZURE_CLIENT_ID=your_id_here
AZURE_CLIENT_SECRET=your_secret_here
TENANT_ID=consumers
REDIRECT_URI=http://localhost:8000/callback
SESSION_SECRET=your_random_secret
X_API_KEY=your_api_key_for_external_apps
USER_EMAIL=your_email@hotmail.com
```
```

## Usage

### Web Interface
Navigate to the root URL and click **Login with Microsoft**. Once authenticated, your token cache is initialized. You can trigger cleanups manually from the dashboard.

### External API
Trigger a cleanup from any other application using a simple POST request:

```bash
curl -X POST https://your-app.onrender.com/api/v1/trigger-clean \
     -H "X-API-KEY: your_api_key_here"
```

| Header | Value | Description |
|--------|-------|-------------|
| `X-API-KEY` | `string` | Your unique API key from `.env` |

## Deployment

### Render.com (Recommended)
1. **Web Service**: Create a new Web Service from your GitHub repo.
2. **Start Command**: `uvicorn server:app --host 0.0.0.0 --port $PORT`
3. **Persistence**: To keep your session active across redeploys, mount a **Persistent Disk** on `/data` and set the environment variable `CACHE_FILE=/data/token_cache.json`.

## License

MIT © [Enrique Cuchetti](https://github.com/ecuchetti)
