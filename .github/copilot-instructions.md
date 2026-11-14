# Project Context Maintenance

- Whenever you modify application code or configuration, review `AI_CONTEXT.md` before finishing and update it in the same task so it describes the resulting behavior, architecture, configuration, and limitations accurately.
- Read `PROJECT_REQUIREMENTS.md` before implementing product features. Keep its coverage statuses factual, update affected rows after implementation, and distinguish implemented behavior from target requirements and blocked acceptance checks.
- Do not claim Excel extraction or Consolidated reconciliation until the actual workbook path and sheet structure have been inspected and tested.
- Keep `AI_CONTEXT.md` concise and factual. Update relevant sections rather than adding a running diff log; do not record changes that have no lasting effect on project context.
- When behavior changes, update `PROJECT_REQUIREMENTS.md` as needed so its coverage matrix stays aligned with the code.
- Verify context statements against the current source and `README.md`. Do not copy secrets, OAuth material, spreadsheet contents, local database data, or logs into the context file.
- If a code change does not alter any durable project behavior or guidance, leave the context unchanged and state that you reviewed it.