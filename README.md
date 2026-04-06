# Teams PST Transcript Extractor

Extract Microsoft Teams chat data from PST files (eDiscovery export) using PowerShell.

## Features
- No Outlook manual work
- Clean transcript output
- Sender + timestamp extraction
- Removes metadata noise

## Usage
1. Update PST path in script
2. Run:
   Set-ExecutionPolicy RemoteSigned -Scope CurrentUser
   .\Export-TeamsTranscript.ps1

## Output
Generates clean .txt transcript for compliance/legal review
