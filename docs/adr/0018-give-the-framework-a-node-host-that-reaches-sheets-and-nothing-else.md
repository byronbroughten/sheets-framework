# Give the framework a Node host that reaches Sheets and nothing else

A Node host runs the framework from the terminal against the live spreadsheet by installing the three Sheets calls and the one script property that are the framework's whole contact with Google, so chores and `gen:configs` run the working tree's code with no push to the deployment. It reaches Sheets only: triggers, Gmail and Docs stay in Apps Script, and the boundary is recorded as one so a missing global is not read as an unfinished adapter.

Source: byronbroughten/byro-repo#25
