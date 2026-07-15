# RiegenCRM Classic ASP Protection Project

## Critical rules

1. Never modify files under src/ unless explicitly instructed.
2. All transformed output must go to dist/.
3. Never deploy readable source code.
4. Preserve Classic ASP include directives.
5. Preserve VBScript semantics.
6. Never rename database table names automatically.
7. Never rename database column names automatically.
8. Never rename Request.Form keys automatically.
9. Never rename Request.QueryString keys automatically.
10. Never rename Session keys automatically unless whole-repository proof exists.
11. Never rename Cookie keys automatically.
12. Never rename COM ProgIDs or method names.
13. Never transform third-party libraries.
14. Every transformation must produce a report.
15. Build must fail on missing include targets.
16. Source must remain byte-for-byte unchanged.