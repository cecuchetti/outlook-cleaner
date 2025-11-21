# Outlook Cleaner

Tool to automatically clean unwanted emails from your Outlook/Hotmail account using OAuth2 authentication.

## Features

- Secure OAuth2 authentication (no App Passwords needed)
- Search by sender name (not just email)
- Automatically moves unwanted emails to deleted folder
- External configuration via JSON file
- Support for multiple senders to filter

## Requirements

- Python 3.7+
- Microsoft account (Hotmail/Outlook)
- Application registered in Azure AD (see configuration)

## Installation

1. Clone the repository:
```bash
git clone <repository-url>
cd outlook-cleaner
```

2. Install dependencies:
```bash
pip install -r requirements.txt
```

3. Configure the application:
```bash
cp config.json.example config.json
```

4. Edit `config.json` with your data (see configuration section)

## Configuration

### 1. Register application in Azure AD

1. Go to [Azure Portal](https://portal.azure.com)
2. **Azure Active Directory** → **App registrations** → **New registration**
3. Name: "Outlook Cleaner" (or your preferred name)
4. **Supported account types**: "Personal Microsoft accounts only"
5. **Redirect URI**: `http://localhost` (type: Public client/native)
6. Click **Register**

### 2. Configure permissions

1. In the registered app, go to **API permissions**
2. Click **Add a permission**
3. Select **Microsoft Graph** → **Delegated permissions**
4. Search and select: `IMAP.AccessAsUser.All`
5. Click **Add permissions**
6. (Optional) If "Grant admin consent" appears, click it

### 3. Get Client ID

1. In **Overview** of your app, copy the **Application (client) ID**
2. Paste it in `config.json` in `oauth2.client_id`

### 4. Configure config.json

Edit `config.json` with your data:

```json
{
  "email": "your-email@hotmail.com",
  "oauth2": {
    "client_id": "your-azure-client-id",
    "tenant_id": "consumers",
    "force_interactive_login": true
  },
  "cleaning": {
    "sender_names_to_search": [
      "Banco Galicia",
      "Claro Video",
      "Farmacity",
      "Carrefour"
    ],
    "move_to_deleted": true
  }
}
```

**Note**: `tenant_id` must be `"consumers"` for personal Microsoft accounts.

## Usage

Run the script:

```bash
python main.py
```

The first time, the browser will open for authentication. After that, the token is cached.

### Read-only mode

To see which emails would be found without moving them, configure:

```json
{
  "cleaning": {
    "move_to_deleted": false
  }
}
```

## Project Structure

```
outlook-cleaner/
├── main.py                 # Main entry point
├── config.py               # Configuration management
├── auth.py                 # OAuth2 authentication
├── filters.py              # Email filtering strategies (Strategy Pattern)
├── imap_service.py         # IMAP service facade (Facade Pattern)
├── config.json.example     # Configuration template
├── config.json             # Your configuration (not uploaded to Git)
├── requirements.txt        # Dependencies
├── .gitignore             # Files to ignore
└── README.md              # This file
```

### Architecture

The codebase follows clean architecture principles with design patterns:

- **Strategy Pattern** (`filters.py`): Flexible email filtering rules
- **Facade Pattern** (`imap_service.py`): Simplifies IMAP complexity
- **Separation of Concerns**: Each module has a single responsibility

## Security

- **WARNING**: Never upload `config.json` to GitHub (it's in `.gitignore`)
- The `config.json.example` file is just a template without sensitive data
- OAuth2 tokens are stored locally in cache by MSAL

## License

[MIT License](LICENSE)

## Contributing

Contributions are welcome. Please:

1. Fork the project
2. Create a branch for your feature (`git checkout -b feature/AmazingFeature`)
3. Commit your changes (`git commit -m 'Add some AmazingFeature'`)
4. Push to the branch (`git push origin feature/AmazingFeature`)
5. Open a Pull Request

## Author

Enrique Cuchetti

## Acknowledgments

- Microsoft for the OAuth2 API
- The Python community

