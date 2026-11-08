# Security

Please **do not open a public issue** for vulnerabilities.
Use GitHub's private reporting: **Security** tab → **Report a vulnerability**.

Include what you found, how to reproduce it, and the impact. You'll get a reply within a few days.

## Scope notes

- outlook-calendar-lite-mcp uses delegated Microsoft Graph permissions only (`Calendars.ReadWrite`, `MailboxSettings.Read`, `User.Read`) and stores tokens locally (`~/.outlook-calendar-lite-mcp`, mode 600).
- Event descriptions and invitations are untrusted input. Anything that notifies other people (invitations, updates, cancellations, replies) requires user confirmation;
  a bypass of that confirmation is in scope and treated as high severity.
