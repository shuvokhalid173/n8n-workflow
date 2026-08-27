The `Cannot GET /` error occurs because n8n is an oclif CLI application. Executing the `bin\n8n` script directly without passing the **`start`** argument boots the CLI entry point rather than launching the web editor and static file routes.

### Resolution Steps

**1. Delete the misconfigured PM2 process**

```powershell
pm2 delete n8n-server

```

**2. Start n8n with the `start` argument passed after `--**`

```powershell
pm2 start "$env:APPDATA\npm\node_modules\n8n\bin\n8n" --name "n8n-server" -- start

```

**3. Save the process configuration**

```powershell
pm2 save

```

**4. Access the Editor**
Refresh `http://localhost:5678` in the browser. The n8n setup UI will now render correctly.



PM2 failed because it tried to execute `N8N.CMD` (a Windows batch file) directly using Node.js, which threw a `SyntaxError` when it encountered the `@ECHO off` batch syntax.

### Solution: Target the JavaScript Binary Directly

Run these commands in your PowerShell window to clean up the errored process and start n8n directly via its JavaScript entry point:

```powershell
# 1. Delete the failed process entry
pm2 delete n8n-server

# 2. Start n8n using its JS file path
pm2 start "$env:APPDATA\npm\node_modules\n8n\bin\n8n" --name "n8n-server"

# 3. Save the process configuration for system reboots
pm2 save

```

### Verification

Check the process status and startup logs:

```powershell
pm2 status
pm2 logs n8n-server --lines 20

```

Once the log displays `Editor is now accessible at http://localhost:5678/`, open `http://localhost:5678` in the browser to access n8n.

On Windows machines with dual-core processors, running `n8n start` directly in PowerShell hangs because n8n defaults to spawning separate child worker processes (`EXECUTIONS_PROCESS=own`), causing the Node event loop to lock before binding to port `5678`.

**Step 1: Terminate the Stuck Process**
Press `Ctrl + C` in PowerShell to stop the hanging shell process.

---

**Step 2: Set Environment Variables to Fix the Hang**
Force n8n to run in single-process mode and bypass Windows file permission checks by running these commands in PowerShell:

```powershell
[System.Environment]::SetEnvironmentVariable('EXECUTIONS_PROCESS', 'main', 'User')
[System.Environment]::SetEnvironmentVariable('N8N_ENFORCE_SETTINGS_FILE_PERMISSIONS', 'false', 'User')

```

*Close PowerShell and open a new PowerShell window as Administrator for these variables to apply.*

---

**Step 3: Start n8n using PM2**
Do not call `n8n start` interactively in the terminal. Pass n8n to PM2 so it executes cleanly in the background:

```powershell
pm2 start n8n --name "n8n-server"

```

---

**Step 4: Verify Status and Logs**
Check if the service is online and inspect the startup logs:

```powershell
pm2 status
pm2 logs n8n-server --lines 20

```

Once the log shows `Editor is now accessible at http://localhost:5678/`, open `http://localhost:5678` in the browser.

---

**Step 5: Persist Across System Reboots**
Save the running PM2 process configuration so n8n launches automatically whenever the PC powers on:

```powershell
pm2-startup install
pm2 save

```