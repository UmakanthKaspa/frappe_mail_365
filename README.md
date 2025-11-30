# Frappe Mail 365

Send and receive emails in Frappe using Microsoft 365 Graph API instead of SMTP/IMAP.

## Features

- Send/Receive via Graph API
- **Email Threading** - Replies from Frappe appear in the same thread in Outlook

## Installation

```bash
bench get-app https://github.com/UmakanthKaspa/frappe_mail_365
bench --site your-site install-app frappe_mail_365
```

## Setup

### Azure App
Add these **Microsoft Graph** permissions (Delegated):
- `Mail.Read`
- `Mail.Send`
- `Mail.ReadWrite`
- `offline_access`

### Connected App Scopes
```
https://graph.microsoft.com/Mail.Read
https://graph.microsoft.com/Mail.ReadWrite
offline_access
```

### Email Account
Enable **"Use Microsoft 365 Graph API"** checkbox.

## License

MIT
