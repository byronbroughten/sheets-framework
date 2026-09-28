# How should a user outside this repo authenticate the harness?

Research for [#26](https://github.com/byronbroughten/sheets-framework/issues/26), part of map #24. Checked 2026-09-27 against Google's own docs, `google-auth-library-nodejs`, gspread and `mcp-google-sheets`.

## Answer

- **Default: a bring-your-own OAuth Desktop client**, the same shape the framework already uses through clasp (`scripts/nodeHost.ts` refreshes a `client_id`/`client_secret`/`refresh_token` triple; `docs/how-it-runs.md` "publish to production" note). The user creates a Desktop OAuth client in their own GCP project, sets the consent screen to External + **In production**, and leaves it unverified. That's allowed for personal use under 100 users. They click past one "unverified app" warning once, and the refresh token then doesn't expire after 7 days. The harness runs a loopback browser flow once and stores `{client_id, client_secret, refresh_token}` in `~/.config/<harness>/` at `chmod 600`. After that, agents run it headless.
- **Fallback: a service account**, for CI, sandboxes and any machine with no browser. The user shares the spreadsheet with the service account's `client_email` as Editor and points the harness at the key JSON (`GOOGLE_APPLICATION_CREDENTIALS` or a flag). It has two limits: it **can't make Drive copies into My Drive**, and key creation may be blocked by org policy.
- **Accept ADC as an input, but don't document it as a setup path.** `gcloud auth application-default login` with Sheets/Drive scopes gets "This app is blocked" unless you also pass `--client-id-file`, which means creating the same OAuth client as the default path. Using `GoogleAuth` means ADC comes free: `GOOGLE_APPLICATION_CREDENTIALS`, a gcloud ADC file, or GCE metadata.
- **Scope choice drives snapshot design.** `spreadsheets` covers every `batchUpdate`. A **Drive-copy snapshot** of a spreadsheet the user made in the browser needs the full `drive` scope, which is *restricted*, because `drive.file` only reaches files the app created or the user picked. If snapshots can stay inside the Sheets API (`duplicateSheet`, or a local JSON dump of `spreadsheets.get` with grid data), the harness needs only `spreadsheets`, and service accounts work fully. Recommendation: ask for `spreadsheets` by default and make `drive` an opt-in scope that `auth login --with-drive` adds.

## Comparison

| | OAuth Desktop client (own project) | Service account | ADC via gcloud |
| --- | --- | --- | --- |
| First-time setup | Create project, enable Sheets (+Drive) API, consent screen (External, publish), Desktop client, download JSON, one browser login. About 10 minutes, all in Console. | Create project, enable APIs, create SA, create key JSON, share each sheet with SA email. About 10 minutes, plus a share per sheet. | Install gcloud, *and still* create a Desktop OAuth client for `--client-id-file`. Most steps of all. |
| Whose identity | The user; edits show as them | A robot; edits show as the SA | The user |
| Sees which sheets | Everything the user can open | Only sheets shared with it | Everything the user can open |
| Consent / verification | Unverified is fine under 100 users (warning screen). Testing status means tokens expire in 7 days, so publish. | None: it accesses its own data | Same as the OAuth client supplied |
| Drive copy (snapshot) | Works with `drive` scope | Fails with 403 in My Drive (no storage); needs a shared drive | Works with `drive` scope |
| Token storage | Refresh token file, `chmod 600` | Long-lived private key JSON | `application_default_credentials.json` (gcloud-managed) |
| Headless / agent | After the one login, yes | Fully | After login, yes |
| Blockers | Workspace admins can block unverified apps | Orgs created on/after 2024-05-03 block SA key creation by default | Default gcloud client blocks Drive/Sheets scopes |

## Findings, with sources

### Scopes and verification

- Sheets scopes: `spreadsheets` and `spreadsheets.readonly` are **Sensitive**; `drive.file` is **Non-sensitive** ("Recommended"); `drive` and `drive.readonly` are **Restricted**. [Sheets API scopes](https://developers.google.com/workspace/sheets/api/scopes)
- `drive.file` gives per-file access: files the app creates, or files the user opens with the app or picks through the Google Picker. It does not reach arbitrary existing files. Restricted scopes that store or send data from servers need a security assessment. [Drive API scopes](https://developers.google.com/workspace/drive/api/guides/api-specific-auth)
- No verification is needed for personal use: "If the app is for your personal use (fewer than 100 users), you and your limited number of users can continue using the app without going through verification." None is needed for service-account-only access either, or for Internal (Workspace org) apps. Unverified apps show a warning screen and have a 100-user cap. [When verification is not needed](https://support.google.com/cloud/answer/13464323)
- Testing status allows up to 100 listed test users, and "Authorizations by a test user will expire seven days from the time of consent." Unverified sensitive or restricted scopes show the "Unverified apps" warning. [Manage app audience](https://support.google.com/cloud/answer/15549945)
- Here is the 7-day rule verbatim: "A Google Cloud Platform project with an OAuth consent screen configured for an external user type and a publishing status of 'Testing' is issued a refresh token expiring in 7 days…" There is also a limit of 100 refresh tokens per Google Account per client ID; the oldest is revoked silently. [Using OAuth 2.0](https://developers.google.com/identity/protocols/oauth2)
- Consequence: a published, unverified personal client is the only user-OAuth setup with no expiry and no review. This repo already relies on it (`packages/framework/docs/how-it-runs.md`, auth section).

### ADC

- `gcloud auth application-default login` defaults to `openid`, `userinfo.email`, `cloud-platform` and `sqlservice.login`. "To add scopes for applications outside of Google Cloud Platform, such as Google Drive, create an OAuth Client ID and provide it by using the --client-id-file flag." [gcloud reference](https://docs.cloud.google.com/sdk/gcloud/reference/auth/application-default/login)
- A "This app is blocked" or "Access blocked" error comes from scopes outside Google Cloud, such as Drive. The fixes are `--client-id-file`, or `--impersonate-service-account`. [Troubleshoot ADC](https://docs.cloud.google.com/docs/authentication/troubleshoot-adc)
- `google-auth-library` ADC order: `GOOGLE_APPLICATION_CREDENTIALS` key file, then environment metadata (GCE, Cloud Run and so on), then the gcloud user ADC file. It refreshes access tokens itself when it has a refresh token. A `refresh_token` "is only returned on the first authorization", so persist it from the `tokens` event. [google-auth-library-nodejs](https://github.com/googleapis/google-auth-library-nodejs)

### Service accounts

- "Standalone service accounts that use a client_email identity don't have personal Drive storage." Saving to My Drive returns 403; the fix is a shared drive (`supportsAllDrives=true`) or domain-wide delegation. That means a `files.copy` snapshot run as an SA fails for a personal Gmail user. [Workspace storage limits](https://knowledge.workspace.google.com/admin/drive/storage-and-upload-limits-for-google-workspace)
- `iam.disableServiceAccountKeyCreation` is enforced by default for organizations created on or after 2024-05-03 (secure-by-default baseline). Personal Gmail projects have no organization, so this doesn't apply to them. [Security baseline constraints](https://docs.cloud.google.com/resource-manager/docs/secure-by-default-organizations)

### How existing tools do it

- **Google's Node quickstart**: enable the API, configure the consent screen, create a Desktop client, save `credentials.json`, then `@google-cloud/local-auth` + `googleapis`; "Authorization information is stored in the file system". [Sheets Node quickstart](https://developers.google.com/workspace/sheets/api/quickstart/nodejs)
- **gspread**: a service account ("for bots", share the sheet with `client_email`, key at `~/.config/gspread/service_account.json`), or an OAuth Desktop client ("for end users", `~/.config/gspread/credentials.json` + `authorized_user.json`, browser on first run), or an API key for public sheets only. Both auth modes ask for `spreadsheets` + `drive`. [gspread auth](https://docs.gspread.org/en/latest/oauth2.html)
- **mcp-google-sheets** (xing5): it recommends a service account for headless and server use, with OAuth for personal or local use. It also accepts base64 `CREDENTIALS_CONFIG` for containers, then ADC. Precedence is `CREDENTIALS_CONFIG` > `SERVICE_ACCOUNT_PATH` > `CREDENTIALS_PATH` (OAuth, `TOKEN_PATH` default `token.json`) > ADC. Scopes are `spreadsheets` + `drive`. [xing5/mcp-google-sheets](https://github.com/xing5/mcp-google-sheets)

## Suggested harness surface (for the spec, not decided)

- A `harness auth login [--client <desktop.json>] [--with-drive]` command runs the loopback flow, writes `~/.config/<harness>/credentials.json` at 0600, and warns if the refresh token looks Testing-scoped (it dies in 7 days).
- Credential precedence: explicit flag > `<HARNESS>_CREDENTIALS` env > stored OAuth token > `GoogleAuth` ADC. A service-account key reaches it through the flag, the env var or `GOOGLE_APPLICATION_CREDENTIALS`.
- Snapshots should default to Sheets-only (`duplicateSheet` or a local JSON dump), with a Drive copy only when `drive` was granted and the identity has storage.
- Inside this repo, a small adapter can keep reading clasp's `~/.clasprc.json`, so nothing here has to change.

## Open questions

- Does the pre-send snapshot really need a Drive copy? This decides whether `drive` (restricted) is ever requested. Belongs to the snapshot or preview ticket.
- Should the harness ship its own verified OAuth client so users skip the Console? It would need Google verification plus a restricted-scope security assessment if it uses `drive`. Out of scope until the public launch (#24 "Out of scope").
