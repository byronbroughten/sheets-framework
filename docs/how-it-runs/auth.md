# When the Node host fails to authenticate

`gen-configs` and the chore runner both refresh the named credential `desktop-clasp-run` out of `~/.clasprc.json`, or the credential in the file `clasp_config_auth` names when that is set. Both are minted by `npx sheets-framework setup-auth`, a wizard of browser steps the developer runs.

## Re-minting the developer's credential

If `gen-configs` or a chore fails with an auth error, the token needs re-minting — run `npx sheets-framework setup-auth` from any package (it runs the framework's `scripts/setup-clasp-run-auth.sh`), which walks through it. The consent screen of the GCP project behind that credential should be published to production; left in "Testing" it would issue refresh tokens that expire in 7 days ([documented](../sheets-behaviour.md#refresh-tokens-under-a-testing-consent-screen)). Publishing alone doesn't fix an existing token, since one minted under "Testing" keeps its expiry — the re-authorization is the part that matters.

Google no longer lets you view or download a client secret after creating it, but you don't need to: clasp stores `client_id` and `client_secret` in `~/.clasprc.json`, and `--creds` reads only those two plus a localhost `redirect_uris` entry, so the file is always rebuildable. The script does that for you. Keep `~/.clasprc.json` at `chmod 600` — the setup assumes a refresh token minted under a published consent screen doesn't expire in 7 days; documented only for "Testing": [the line](../sheets-behaviour.md#refresh-tokens-under-a-testing-consent-screen).

## Minting the agent account's credential

The wizard's last stage, which asks before it runs, mints a credential for the agent account rather than the developer's, so code that skips the framework reaches only what is shared with that account. It runs the OAuth flow against the GCP project's OAuth client for only the scopes in the script's `AGENT_SCOPES`, and writes a `.clasprc.json`-format file at `AGENT_AUTH_FILE`, mode `600`. In the browser, sign in as the agent account. The wizard unsets `clasp_config_auth` for its own run, so re-running it with the variable exported still re-mints the developer's credential.

To have the Node host run as the agent account, set `clasp_config_auth` to that file's full path in the environment that runs it, as the stage prints it, since Node doesn't expand a `~`; unset, the host goes back to the developer's `desktop-clasp-run` credential. To revoke the credential, sign in to Google as the agent account, open its third-party access page at <https://myaccount.google.com/connections>, pick the app named on the project's consent screen (its Branding tab) and remove its access; the file's refresh token then stops working, and re-running the stage mints a new one.
