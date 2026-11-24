# LeadConnect Pro

Local, Egypt-first prospect manager and WhatsApp outreach assistant. Python 3.10+ with Tkinter and SQLite; no third-party dependencies.

## Windows installation

In Command Prompt:
```cmd
cd /d "%USERPROFILE%\Desktop"
git clone https://github.com/ebroboooo/LeadConnect-Pro.git
cd LeadConnect-Pro
python leadconnect.py --self-test
python leadconnect.py
```

If `python` is not recognized, try `py -3` instead. The self-test checks the GUI and an in-memory database without accessing your contacts.

## Workflow

Import MapHarvest CSV, JSON, XLSX (first numbered worksheet, headers in row 1), or paste numbers. The Egyptian mobile filter excludes landlines and other formats from the default view; a valid mobile format is not proof of WhatsApp registration. Verify each listing and record recipient permission. Edit the message template, review an individual recipient, open WhatsApp and send manually. Confirm sent only after actually sending. Match incoming replies manually using the sender number; unrelated conversations are not recorded.

No automatic bulk messaging, account probing or inbox monitoring is provided. Use normal WhatsApp notifications. Official integration is necessary for reliable automated reply detection.

Contacts and history stay in `%USERPROFILE%\LeadConnectPro\leadconnect.sqlite3`. Back up this file privately after closing the program. Never commit contact exports, databases, credentials or message logs. No external API or paid dependency is required.

The repository contains source code, not an executable. Windows security protections should remain enabled. Windows-specific behavior requires local testing.
