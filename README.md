# ITTool

Lightweight Windows diagnostic and reporting tool for support engineers.

## Run

From a local checkout:

```powershell
powershell.exe -NoProfile -ExecutionPolicy Bypass -File .\Invoke-ITTool.ps1
```

With ticket metadata:

```powershell
powershell.exe -NoProfile -ExecutionPolicy Bypass -File .\Invoke-ITTool.ps1 -Ticket "INC12345" -Tech "jsmith" -Issue "VPN disconnects"
```

Remote `irm` / `iex` style launch after publishing the script:

```powershell
irm "https://raw.githubusercontent.com/slitzer/PulseCheck/main/Invoke-ITTool.ps1" | iex
```

## Output

Each launch creates a run folder under:

```text
C:\Temp\ITTool\Runs\COMPUTERNAME_yyyy-MM-dd_HHmmss
```

The run folder contains:

- `report.html`
- `summary.json`
- `task-log.jsonl`
- `transcript.log`
- `run-info.txt`
- `Logs\` task stdout/stderr files
- `Reports\` generated HTML reports such as battery and WLAN reports
- `Exports\` CSV and JSON exports

## Notes

- Recommended diagnostics are non-admin by default.
- Admin-only checks are visible and clearly marked.
- Admin-only checks are skipped with a reason when the tool is not elevated.
- The tool does not store credentials, email settings, or client-specific recipients.
- Repair/reset actions are not included in the recommended diagnostic run.
