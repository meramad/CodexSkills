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
- Do not switch to direct PowerShell silently.
- Do not switch to non-broker Excel silently.
- Do not create alternative TCP/HTTP command channels.
- Do not iterate through different broker start commands/flags. Use exactly the start command in this skill.

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

## Start Broker If Not Running (robust, no iteration)

1. Dot-source the helper:
   - `. $helper`

2. Probe broker:
   - `Invoke-PSBroker -PipeName $pipe -Command 'broker.info' -PassThru`

3. If probe fails, start broker **once** (detached) and remember you started it:
   - `$brokerProc = Start-Process -FilePath $brokerExe -WorkingDirectory $brokerRoot -ArgumentList @('--pipe', $pipe, '--log-option', 'silent') -PassThru`
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

## Excel Policy
- Always try broker-native Excel commands (`broker.excel.*`) first.
- When calling `broker.excel.get_workbook_handle`, always pass an **absolute local file path** (never relative paths).
- If broker-native Excel fails or Excel is unavailable:
  - Tell the user exactly why.
  - Then fall back to non-broker Excel automation only if necessary.

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
