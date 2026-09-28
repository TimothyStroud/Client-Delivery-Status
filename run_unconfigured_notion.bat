@echo off
cd /d H:\
echo ==== %DATE% %TIME% ==== >> "C:\Users\tls2\.claude\projects\H--\unconfigured_notion_run.log"
type "C:\Users\tls2\.claude\projects\H--\unconfigured_notion_prompt.txt" | "C:\Users\tls2\.local\bin\claude.exe" -p --model sonnet --allowedTools "Bash(python:*),mcp__plugin_Notion_notion__notion-fetch,mcp__plugin_Notion_notion__notion-update-page" >> "C:\Users\tls2\.claude\projects\H--\unconfigured_notion_run.log" 2>&1
