# Windows 11 Bootstrap — First Draft (Interactive)

**What this does**
- Installs: Python 3, Windows Terminal, Notepad++, VS Code (system), 7-Zip, Starship, Git (via winget).
- Enables WSL components and sets WSL2 default; attempts Ubuntu install (may fail if Store content is blocked).
- Sets Git global identity to **Dallas Bleak / dbleak42@gmail.com**.
- Shows file extensions and hidden files in Explorer.
- Adds safe context-menu verbs (instead of forced defaults):
  - .ps1 → **Run with PowerShell 7**
  - .md  → **Open with VS Code**
- Optional local installers for **SecureCRT** and **RSAT AD tools** (offline FoD).

**Log location:** `C:\Temp\Logs`

---

## Usage

1. Copy `win11-bootstrap.ps1` and `win11-apps.json` to a folder on the new machine.
2. Open **PowerShell as Administrator** and run:
```powershell
Set-ExecutionPolicy Bypass -Scope Process -Force
cd <folder>
.\win11-bootstrap.ps1
```
3. If you want SecureCRT and RSAT AD tools:
   - Edit `win11-bootstrap.ps1` and set:
     - `$SecureCRTInstaller` → path to your SecureCRT installer (EXE/MSI)
     - `$SecureCRTSilent` → silent args (often `/S` or `/qn`)
     - `$RsatFoDPath` → folder containing RSAT FoD CABs/MUMs for AD tools
   - Re-run the script.

---

## Notes
- **Store blocked:** winget still works for community repos; Ubuntu install via `wsl --install` may need Store content. If blocked, use an offline Ubuntu `.appx` and I can extend the script to sideload it.
- **File associations:** Windows 11 protects “UserChoice” defaults with a hash; this draft adds context menu verbs rather than forcing defaults. I can add a sanctioned default-apps policy if desired.
- **Unattended:** This draft is interactive by design (no auto reboot).

---

### Next iteration (once you confirm):
- Confirm whether to include **PowerShell 7** install + profile (modules were left unchecked).
- Provide the SecureCRT installer path and RSAT FoD source folder (or allow online RSAT capability install).
- Decide if you want me to add an **offline Ubuntu** sideload step if Store is blocked for WSL.
