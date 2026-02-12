---
name: broker-powershell
description: Use PersistentPowerShellBroker for all PowerShell and Excel automation via named pipe; prefer Invoke-PSBroker.ps1 and only fall back when broker is unavailable.
---

# Broker PowerShell Skill

## Purpose
Use `PersistentPowerShellBroker` as the default execution engine for PowerShell and Excel automation.

This skill is mandatory when broker support exists. Do not invent alternate direct PowerShell or Excel approaches first.

## Hard Rules
- Do use the broker first for any PowerShell execution.
- Do use `Invoke-PSBroker.ps1` first; raw JSON pipe calls are fallback only.
- Do keep one chosen pipe name for the whole work session.
- Do tell the user the chosen pipe/channel key when you pick it.
- Do reuse the same pipe consistently in later commands.
- Do check broker availability with `broker.info` before running work.
- Do use broker-native Excel commands (`broker.excel.*`) before any non-broker Excel automation.
- Do open Excel workbooks in visible mode for interactive work; avoid headless Excel automation.
- Do not switch to direct PowerShell silently.
- Do not switch to non-broker Excel silently.
- Do not create alternative TCP/HTTP command channels.
- Do not iterate through different broker start commands/flags. Use exactly the start command in this skill.

Operational clarification:
- If a chosen pipe becomes unresponsive (hang/timeouts), declare that broker session failed, tell the user, and start a new broker session with a new pipe name.
- This is not considered pipe thrashing; it is failure recovery.

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

## Working Directory and Path Assumptions
- Do not assume the broker is in the current working directory.
- Always resolve broker paths from `%PERSISTENTPSBROKER%` and use `$brokerRoot` as the working directory when starting the broker.
- Always dot-source the helper using its full path from `$brokerRoot` (never `./Invoke-PSBroker.ps1`).

## Session Pipe Policy
Choose one pipe key per session, for example `psbroker-<random8>`.

Immediately tell the user:
- which pipe you selected
- that you will reuse it for this session

Keep it in a session variable and reuse it.

Failure recovery:
- When a pipe fails, explicitly announce:
  - old pipe name,
  - failure symptom (timeout/hang/no response),
  - new pipe name for the recovery session.

## Start Broker If Not Running (robust, no iteration)

1. Dot-source the helper:
   - `. $helper`

2. Probe broker:
   - `Invoke-PSBroker -PipeName $pipe -Command 'broker.info' -PassThru`

3. If probe fails, start broker **once** (detached) and remember you started it:
   - `$brokerProc = Start-Process -FilePath $brokerExe -WorkingDirectory $brokerRoot -ArgumentList @('--pipe', $pipe) -PassThru`
   - `$brokerStartedBySession = $true`
   - Tell the user: “Started broker (silent) on pipe <pipe>.”

4. Poll for readiness (do NOT restart broker during this wait):
   - Retry `broker.info` every 500ms, up to 30s total.
   - Stop polling on first successful response.
   - If the current shell command times out but the broker process exists (PID is running), assume startup is still in progress and continue probing in a new shell command.
   - If 30s passes without success, tell the user and only then fall back if necessary.
   - If polling times out, surface: `brokerExe`, `pipe`, and whether the process is still running (e.g. `$brokerProc.HasExited` if available).

If broker still cannot start/connect:
- Explicitly tell the user you are falling back to direct PowerShell and why.

## Preferred Helper Usage
Always prefer helper calls (use `-PassThru` so failures return structured results):

- Native command:
  - `Invoke-PSBroker -PipeName $pipe -Command 'broker.info' -PassThru`

- Native with args:
  - `Invoke-PSBroker -PipeName $pipe -Command 'broker.excel.get_workbook_handle' -Args @{ path = 'C:\Temp\Book1.xlsx' } -PassThru`

- Free-form PowerShell:
  - `Invoke-PSBroker -PipeName $pipe -Script 'Get-Date' -PassThru`

Scope clarification:
- For model/excel operations, run PowerShell logic in broker runspace via `Invoke-PSBroker -Script`.
- Local shell PowerShell is allowed for broker bootstrap/discovery, process diagnostics, and non-model file edits.

## Excel Policy
- Always try broker-native Excel commands (`broker.excel.*`) first.
- When calling `broker.excel.get_workbook_handle`, always pass an **absolute local file path** (never relative paths).
- When calling `broker.excel.get_workbook_handle`, always pass `forceVisible = $true` unless the user explicitly asks for hidden/background behavior.
- After opening/attaching a workbook handle, always verify visibility in broker runspace:
  - `$app = $handle.Workbook.Application`
  - `$win = $handle.Workbook.Windows.Item(1)`
  - confirm `$app.Visible -eq $true` and `$win.Visible -eq $true`
- If visibility check fails, immediately force visibility via handle before continuing:
  - set `$app.Visible = $true`
  - set `$app.UserControl = $true`
  - set `$win.Visible = $true`
  - activate workbook/window
- If workbook still cannot be made visible, stop and report the exact reason to the user before continuing with more files.
- Rationale: Excel COM automation is unreliable in fully headless/background mode; visible interactive mode is the default expectation.
- If broker-native Excel fails or Excel is unavailable:
  - Tell the user exactly why.
  - Then fall back to non-broker Excel automation only if necessary.

Handle lifecycle policy:
- Prefer short-lived deterministic handle usage:
  1. `get_workbook_handle`
  2. perform one logical batch
  3. verify expected result
  4. save/close/release handle when done
- If a `psVariableName` release fails, treat handle as stale; reacquire by workbook path before retrying close/release.

Refresh reliability playbook:
- After `RefreshAll`, do not tight-loop COM status checks.
- Use delayed sparse polling:
  - initial wait 20-30s,
  - poll every 10-15s,
  - typically 3-4 polls.
- If still refreshing, report exact connection names (`Workbook.Connections`) as stuck.
- If refresh request causes broker call timeout/hang, stop current session and recover on a new pipe.

## Raw JSON Pipe Fallback (Only If Helper Unavailable)
One request per connection, one JSON line response, UTF-8.

Request envelope fields:
- `id` (string)
- `kind` (`native` or `powershell`)
- `command` (string)
- `args` (object, optional)
- `timeoutMs` (int, optional)

Response envelope fields:
- `id` (string)
- `success` (bool)
- `stdout` (string)
- `stderr` (string)
- `error` (string|null)
- `durationMs` (int)

## Broker Lifecycle
- If this session started the broker (`$brokerStartedBySession = $true`), it may stop it when work is complete.
- Otherwise, do not stop the broker.
