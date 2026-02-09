---
name: broker-powershell
description: Use PersistentPowerShellBroker for all PowerShell and Excel automation via named pipe; prefer Invoke-PSBroker.ps1 and only fall back when broker is unavailable.
---

# Broker PowerShell Skill

## Purpose
Use `PersistentPowerShellBroker` as the default execution engine for PowerShell and Excel automation.

This skill is mandatory when broker support exists. Do not invent alternate direct PowerShell or Excel approaches first.

## Hard Rules
- Do use broker first for any PowerShell execution.
- Do use `Invoke-PSBroker.ps1` first; raw JSON pipe calls are fallback only.
- Do keep one chosen pipe name for the whole work session.
- Do tell the user the chosen pipe/channel key when you pick it.
- Do reuse the same pipe consistently in later commands.
- Do check broker availability with `broker.info` before running work.
- Do use broker-native Excel commands (`broker.excel.*`) before any non-broker Excel automation.
- Do call `broker.stop` only when the user explicitly asks, or when this session started that broker instance.
- Do not switch to direct PowerShell silently.
- Do not switch to non-broker Excel silently.
- Do not create alternative TCP/HTTP command channels.

## Environment and Discovery
Broker folder is provided by environment variable `%PERSISTENTPSBROKER%`.

In PowerShell:
- Resolve broker root: `$brokerRoot = $env:PERSISTENTPSBROKER`
- Broker exe path: `$brokerExe = Join-Path $brokerRoot 'PersistentPowerShellBroker.exe'`
- Helper path: `$helper = Join-Path $brokerRoot 'Invoke-PSBroker.ps1'`

Both files are expected in the same folder.

If `%PERSISTENTPSBROKER%` is missing or invalid:
- Stop and tell the user exactly what is missing.
- Ask for the correct broker folder.

## Session Pipe Policy
Choose one pipe key per session, for example:
- `psbroker-<random8>`

Immediately tell the user:
- which pipe you selected
- that you will reuse it for this session

Keep it in a session variable and reuse it.

## Start Broker If Not Running
1. Dot-source the helper:
   - `. $helper`
2. Probe broker:
   - `Invoke-PSBroker -PipeName $pipe -Command 'broker.info' -PassThru`
3. If probe fails, start broker:
   - `$brokerProc = Start-Process -FilePath $brokerExe -ArgumentList @('--pipe', $pipe, '--log-option', 'silent') -PassThru`
   - `$brokerStartedBySession = $true`
4. Poll `broker.info` with fixed policy:
   - total timeout: `30s`
   - retry interval: `500ms`
   - stop polling on first successful response
   - if timeout reached, treat as broker unavailable

If broker still cannot start/connect:
- Explicitly tell the user you are falling back to direct PowerShell and why.
- Use direct PowerShell only as needed to unblock the task.

## Preferred Helper Usage
Always prefer helper calls:
- Native command:
  - `Invoke-PSBroker -PipeName $pipe -Command 'broker.info' -PassThru`
- Native with args:
  - `Invoke-PSBroker -PipeName $pipe -Command 'broker.excel.get_workbook_handle' -Args @{ path = 'C:\Temp\Book1.xlsx' } -PassThru`
- Free-form PowerShell:
  - `Invoke-PSBroker -PipeName $pipe -Script 'Get-Date' -PassThru`

Use `-PassThru` in automation so failures return structured results without immediate throw.

## Raw JSON Pipe Fallback (Only If Helper Unavailable)
Request envelope (one line JSON, one request per connection):
- `id` (string)
- `kind` (`native` or `powershell`)
- `command` (string)
- `args` (object, optional)
- `timeoutMs` (int, optional)
- `clientName` (string, optional)
- `clientPid` (int, optional)

Response envelope (one line JSON):
- `id` (string)
- `success` (bool)
- `stdout` (string)
- `stderr` (string)
- `error` (string|null)
- `durationMs` (int)

Protocol behavior:
- one request per connection
- one JSON line response
- UTF-8
- disconnect after response

## Current Native Commands (from source)

### `broker.info`
Returns broker metadata.
- Inputs: none
- Output fields: `version`, `pipeName`, `startedAtUtc`, `pid`

### `broker.stop`
Requests graceful shutdown after replying.
- Inputs: none
- Output: `stdout` usually indicates stopping

### `broker.help`
Returns command index or per-command schema.
- Args:
  - `command` (string, optional)
  - `format` (`Json|Text|Both`, optional, default `Json`)
- Status values: `Success`, `NotFound`

### `broker.excel.get_workbook_handle`
Finds or opens workbook and stores global handle bundle.
- Required args:
  - `path` (local path or SharePoint/OneDrive URL)
- Optional args:
  - `readOnly` (bool, default `false`)
  - `openPassword` (string|null)
  - `modifyPassword` (string|null)
  - `timeoutSeconds` (int, default `15`)
  - `instancePolicy` (`ReuseIfRunning|AlwaysNew`, default `ReuseIfRunning`)
  - `displayAlerts` (bool, default `false`)
  - `forceVisible` (bool, default `true`)
- Key output fields:
  - `ok`, `status`, `psVariableName`, `workbookFullName`, `requestedTarget`
  - `attachedExisting`, `openedWorkbook`, `isReadOnly`

### `broker.excel.release_handle`
Releases stored handle; can close workbook and quit Excel.
- Required args:
  - `psVariableName` (string)
- Optional args:
  - `closeWorkbook` (bool, default `false`)
  - `saveChanges` (bool|null)
  - `quitExcel` (bool, default `false`)
  - `onlyIfNoOtherWorkbooks` (bool, default `true`)
  - `timeoutSeconds` (int, default `10`)
  - `displayAlerts` (bool, default `false`)
- Key output fields:
  - `ok`, `status`, `closedWorkbook`, `quitExcelAttempted`, `quitExcelSucceeded`, `released`

## Excel Fallback Rules (Explicit)
- First choice: `broker.excel.get_workbook_handle` and `broker.excel.release_handle`.
- If broker-native Excel command fails/unavailable:
  - Tell the user exactly what failed and why.
  - Then use non-broker Excel automation only if necessary.
- After fallback, state clearly that execution is outside broker and session persistence differs.

## Operational Notes
- Broker runspace is persistent across requests.
- Keep commands deterministic and short when possible.
- For command discovery, call `broker.help` rather than guessing.
- Stop policy:
  - if user asks to stop broker, call `broker.stop`
  - if this session started broker (`$brokerStartedBySession = $true`), it may stop it when work is complete
  - otherwise do not stop broker
