# Installation Scripts

This folder contains PowerShell scripts for installing software on Windows systems, both locally and remotely.

## 📁 Contents

- **Install-ApplicationWinget.ps1** - Local winget-based application installer
- **install_software.ps1** - Remote software deployment tool using PowerShell Remoting

---

## 🔹 Install-ApplicationWinget.ps1

### Overview
Installs a predefined list of applications sequentially using Windows Package Manager (winget). The script waits for each installation to complete before proceeding to the next, ensuring a controlled and logged installation process.

### Features
- ✅ Sequential installation (one app at a time)
- ✅ Automatic log directory creation
- ✅ Detailed logging with timestamps
- ✅ Exit code verification
- ✅ Silent installation mode
- ✅ Automatic acceptance of package/source agreements

### Default Applications
The script is pre-configured to install:
- Microsoft Visual Studio Code
- Yubico Authenticator
- Microsoft Power Automate Desktop
- DisplayLink Graphics Driver

### Prerequisites
- Windows 10 1809+ or Windows 11
- Windows Package Manager (winget) installed
- Administrator privileges (recommended)

### Configuration
Edit the `$appList` array to customize applications:
```powershell
$appList = @(
  "Microsoft.VisualStudioCode",
  "YourApp.PackageId"
)
```

### Logging
- **Log Directory**: `C:\temp\logs\`
- **Log File**: `winget_install_log.txt`
- Logs include timestamps, success/failure status, and exit codes

### Usage
```powershell
# Run the script
.\Install-ApplicationWinget.ps1

# Or run with execution policy bypass
powershell -ExecutionPolicy Bypass -File .\Install-ApplicationWinget.ps1
```

### Exit Codes
- `0` - Successful installation
- Non-zero - Installation failed (logged with exit code)

---

## 🔹 install_software.ps1

### Overview
A comprehensive remote software deployment tool that allows technicians to install software packages on single or multiple Windows computers via PowerShell Remoting (WinRM). The script provides an interactive menu system for flexible deployment scenarios.

### Features
- 🖥️ Single or multi-computer deployment
- 📦 Multiple software package selection via GUI
- 🔐 Optional credential authentication
- 📝 CSV logging for unreachable computers
- 🧪 Dry-run mode for testing
- 🔄 Automatic WinRM service management
- 📂 Automatic remote directory creation

### Architecture
This script acts as a wrapper/menu system that delegates to a refactored implementation:
- Main script: `install_software.ps1` (UI/menu system)
- Worker script: `installSoftware.refactored.ps1` (deployment logic)

### Prerequisites
- **PowerShell 5.1+** or PowerShell Core (pwsh)
- **Administrator privileges** on local and remote machines
- **PowerShell Remoting (WinRM)** enabled on target computers
- **Network share** with software packages at relative path `..\\Software`
- **Software structure**: Each app in a folder with `Install.cmd`

### Software Package Structure
```
Software/
├── ApplicationName1/
│   └── Install.cmd
├── ApplicationName2/
│   └── Install.cmd
└── ApplicationName3/
    └── Install.cmd
```

### Computer List Files
The script creates/uses text files for multi-computer deployments:
- **Location**: `fileNames\<username>.txt`
- **Format**: One computer name per line
- **Auto-created**: If file doesn't exist, script creates it and opens it for editing

### Usage

#### Interactive Menu
```powershell
.\install_software.ps1
```

**Menu Options:**
1. **Install on a single computer**
   - Enter computer name
   - Select applications via GUI
   - Deploy immediately

2. **Install on multiple computers (from file)**
   - Edit/review computer list file
   - Select applications via GUI
   - Deploy to all listed computers

3. **Dry-run (show what will run)**
   - Preview deployment without executing
   - Verify computer list and applications
   - Test configuration

Q. **Quit**

#### Credential Prompt
During execution, you'll be asked:
```
Provide credentials? (Y/N)
```
- **Y**: Enter alternate credentials for remote connections
- **N**: Use current user credentials

### Workflow

1. **Launch Script** → Interactive menu appears
2. **Select Mode** → Single/Multiple/Dry-run
3. **Choose Target(s)** → Enter computer name or edit file list
4. **Select Software** → GUI picker shows available applications
5. **Credentials** → Optional alternate credentials
6. **Deploy** → Script handles:
   - Remote session creation
   - C:\Temp directory creation
   - Software copy to remote C:\Temp\<App>
   - Silent installation execution via Install.cmd

### Error Handling
- **Unreachable computers**: Logged to CSV file
- **Missing software**: Warning displayed
- **WinRM failures**: Detailed error messages
- **Service preservation**: Doesn't disable services that were already running

### Logging
- Unreachable computers logged to CSV format (not plain text)
- Real-time console feedback with color coding
- Installation status per computer

### Limitations
- Remote installer **must** be named `Install.cmd`
- Current user needs **admin rights** on remote machines
- **WinRM** must be accessible between machines
- Software packages must be in expected directory structure

### Version History
- **Version**: 1.4.1
- **Date**: 2025-08-19
- **Author**: Dallas Bleak

---

## 🚀 Quick Start

### Local Installation (Winget)
```powershell
# Install default applications locally
.\Install-ApplicationWinget.ps1
```

### Remote Deployment (Single Computer)
```powershell
# Launch interactive menu
.\install_software.ps1

# Select option 1
# Enter: COMPUTER01
# Select applications from GUI
# Proceed with installation
```

### Remote Deployment (Multiple Computers)
```powershell
# Launch interactive menu
.\install_software.ps1

# Select option 2
# Edit computer list file (opens automatically)
# Add computer names (one per line)
# Select applications from GUI
# Proceed with installation
```

---

## 🔒 Security Considerations

- Both scripts require elevated privileges
- Remote deployments need admin rights on target machines
- Credentials transmitted via WinRM (encrypted by default)
- Consider using secure credential storage for production environments

---

## 📝 Notes

- **Install-ApplicationWinget.ps1**: Self-contained, no external dependencies
- **install_software.ps1**: Requires refactored script and Software directory structure
- Both scripts include comprehensive error handling and logging
- Scripts can be customized for organizational needs

---

## 🆘 Troubleshooting

### Winget Script Issues
- **Winget not found**: Install App Installer from Microsoft Store
- **Permission denied**: Run as Administrator
- **Package not found**: Verify package IDs with `winget search <app>`

### Remote Deployment Issues
- **WinRM errors**: Enable WinRM with `Enable-PSRemoting -Force`
- **Access denied**: Verify admin rights on remote machines
- **Computer not found**: Check DNS/NetBIOS name resolution
- **Refactored script missing**: Ensure `installSoftware.refactored.ps1` exists

### Network Issues
- Verify firewall allows WinRM (port 5985/5986)
- Check network share accessibility
- Confirm remote computer is powered on and accessible

---

## 📚 Additional Resources

- [Windows Package Manager (winget) Documentation](https://learn.microsoft.com/en-us/windows/package-manager/)
- [PowerShell Remoting (WinRM) Guide](https://learn.microsoft.com/en-us/powershell/scripting/learn/remoting/running-remote-commands)
- [about_Remote Help](https://learn.microsoft.com/en-us/powershell/module/microsoft.powershell.core/about/about_remote)
