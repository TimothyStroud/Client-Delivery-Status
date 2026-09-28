@echo off
cd /d H:\
type "C:\Users\tls2\.claude\projects\H--\_test_notion_prompt.txt" | "C:\Users\tls2\.local\bin\claude.exe" -p --model sonnet --allowedTools "Bash(python:*),mcp__plugin_Notion_notion__notion-fetch,mcp__plugin_Notion_notion__notion-update-page"
