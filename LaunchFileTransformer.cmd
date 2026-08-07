@echo off
rem FALLBACK ONLY - NOT the launcher in use. LaunchFileTransformer.vbs is, run
rem by wscript.exe, because cmd.exe flashes a console window and wscript does
rem not. Keep this file in case VBScript is removed from Windows: publish it via
rem SIDECARS in cmse_report.py and point the filetransformer: command in
rem Register-CMSE-Tools.reg back at `cmd.exe /d /c "<this file>"`.
rem
rem Argument-ignoring launcher for the File Transformer ClickOnce deployment,
rem used by the "filetransformer:" URL protocol on the CMSE Dashboard.
rem
rem Three separate traps are handled here - all three were confirmed by testing
rem on 2026-08-07, and each one fails SILENTLY, so don't "simplify" them away:
rem
rem  1. dfshim does NOT strip surrounding quotes. The deployment URL must be
rem     passed UNQUOTED or ShOpenVerbApplication just does nothing. (Safe here:
rem     the percent-encoded URL contains no spaces.)
rem  2. A URL protocol handler whose shell\open\command has no %%1 gets the
rem     clicked URL APPENDED by Windows anyway, and dfshim takes exactly one
rem     argument. This file never references %%1, so every argument is
rem     discarded and dfshim always receives exactly one URL.
rem  3. Inside a batch file "%%20" is required to emit a literal "%20"; a bare
rem     %%20 is parsed as batch parameter %%2 followed by "0". Keeping the URL
rem     in here rather than in the registry value also avoids the shell doing
rem     its own %%2 substitution on the command template.
rem
rem dfshim rejects a plain UNC path, so the percent-encoded file:// form is
rem mandatory. The URL is hard-coded and can never be influenced by the caller,
rem which keeps the protocol from becoming an argument-injection hole.
rundll32.exe dfshim.dll,ShOpenVerbApplication file://trgfile1/Operations/Software/_Source/ToolBox/File%%20Transformer/File%%20Transformer.application
