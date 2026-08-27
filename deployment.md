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