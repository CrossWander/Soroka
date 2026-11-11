# Security policy

Soroka works with your Microsoft Teams session and your work chats, so security reports matter here.

## Supported versions

Only the latest release gets fixes. Soroka tells you when a new version is out, so please update before reporting.

## Reporting a vulnerability

Please report it privately. Do not open a public issue.

Use **Security → Report a vulnerability** on this repository: [open a private report](https://github.com/CrossWander/Soroka/security/advisories/new). Only the maintainer sees it.

If that does not work for you, email dikalopavel@gmail.com.

Please include:
- the Soroka version (Settings → footer) and the Windows version;
- what an attacker can do, and what they need for it: a message in a chat, a web page, a program on the same computer;
- the steps to reproduce it.

Do not include real chat contents, `token.txt`, the Jira token or your Teams session. Made-up data is enough to show the problem.

This is a personal project maintained in spare time. You will get an answer as soon as possible. A confirmed issue is fixed in a new release, and the report is published after that.

## What counts

For example:
- a chat message, a web page or a link that makes Soroka or its AI run a program, read a file or send something without your click;
- reaching the local API (`127.0.0.1:8765`) or the app's internal calls from a web page or without the token;
- a way to make the update check install something that is not the installer from this repository's releases;
- your Teams session, the local API token or the Jira token leaving the computer or being stored in plain text.

Bugs in Microsoft Teams itself, or in the AI programs Soroka starts, belong to their makers.
