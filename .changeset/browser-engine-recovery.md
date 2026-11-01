---
"@consultchimps/messages": patch
---

Give `DB_BROWSER_ENGINE_UNAVAILABLE` its own recovery steps. The browser
database tool reports this code when its engine cannot start, most often because
another tab or window of the same browser still holds the working database
files. The generic database advice told the reader to check schema and import
options; the steps now say to close the other tab, reload, and check the
connection and site storage.
