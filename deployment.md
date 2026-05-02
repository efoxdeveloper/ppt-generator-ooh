# Windows Server Deployment Guide (`ppt-optimizer`)

## 1. Prerequisites
- Windows Server 2019/2022
- Administrator access
- Node.js LTS (v20+ recommended)
- Git (optional, if pulling repo)
- IIS + URL Rewrite + ARR (for reverse proxy, optional but recommended)

## 2. Copy Project
Place project in a stable path, for example:
- `C:\apps\ppt-optimizer`

Inside project, ensure these exist:
- `server.js`
- `package.json`
- `templates\` (or API-provided template URL flow)

## 3. Install Dependencies
Open PowerShell in project folder:

```powershell
cd C:\apps\ppt-optimizer
npm ci
```

If `npm ci` fails due to lock mismatch, use:

```powershell
npm install
```

## 4. Run Locally (Smoke Test)
```powershell
npm start
```

Check:
- `http://localhost:5000/api/ppt/job-status/test` (should return not found JSON)

Stop server after check (`Ctrl+C`).

## 5. Run with PM2 (Recommended)
PM2 is recommended for process management, restart, logs, and easy updates.

### Install PM2
```powershell
npm install -g pm2
```

### Start App with PM2
```powershell
cd C:\apps\ppt-optimizer
pm2 start server.js --name ppt-optimizer
```

### Useful PM2 Commands
```powershell
pm2 status
pm2 logs ppt-optimizer
pm2 restart ppt-optimizer
pm2 stop ppt-optimizer
pm2 delete ppt-optimizer
```

### Save Process List (Auto-resurrect)
```powershell
pm2 save
```

### Enable Startup on Boot (Windows)
Run:
```powershell
pm2 startup
```
PM2 will print a command. Run that command in elevated PowerShell, then:
```powershell
pm2 save
```

### Optional Ecosystem File
Create `ecosystem.config.js` in project root:

```js
module.exports = {
  apps: [
    {
      name: "ppt-optimizer",
      script: "server.js",
      instances: 1,
      exec_mode: "fork",
      autorestart: true,
      watch: false,
      max_memory_restart: "1G",
      env: {
        NODE_ENV: "production"
      }
    }
  ]
};
```

Start using:
```powershell
pm2 start ecosystem.config.js
pm2 save
```

## 6. Alternative: Run as Windows Service (NSSM)
Use this only if you do not want PM2.

### Install NSSM
- Download NSSM: https://nssm.cc/download
- Extract, for example: `C:\tools\nssm\`

### Create Service
```powershell
C:\tools\nssm\win64\nssm.exe install PPTOptimizer
```

Set:
- **Path**: `C:\Program Files\nodejs\node.exe`
- **Startup directory**: `C:\apps\ppt-optimizer`
- **Arguments**: `server.js`

Then:
```powershell
C:\tools\nssm\win64\nssm.exe start PPTOptimizer
```

Enable auto start:
```powershell
Set-Service -Name PPTOptimizer -StartupType Automatic
```

## 7. Open Firewall (if direct port access needed)
```powershell
New-NetFirewallRule -DisplayName "PPT Optimizer 5000" -Direction Inbound -Protocol TCP -LocalPort 5000 -Action Allow
```

## 8. IIS Reverse Proxy (Recommended)
Use IIS as public entrypoint and keep Node on localhost.

### IIS Setup
1. Install IIS role
2. Install **URL Rewrite** module
3. Install **Application Request Routing (ARR)**
4. In ARR settings, enable proxy

### Site/Web.config Rule
In IIS site root, add rewrite rule to proxy to Node app:

```xml
<configuration>
  <system.webServer>
    <rewrite>
      <rules>
        <rule name="ReverseProxyToNode" stopProcessing="true">
          <match url="(.*)" />
          <action type="Rewrite" url="http://localhost:5000/{R:1}" />
        </rule>
      </rules>
    </rewrite>
  </system.webServer>
</configuration>
```

## 9. Recommended Folder Permissions
Service account needs read/write on:
- `C:\apps\ppt-optimizer\output`
- `C:\apps\ppt-optimizer\media`
- `C:\apps\ppt-optimizer\templates`

## 10. Production API Flow
1. Start job:
   - `POST /api/ppt/generate-proposal`
2. Poll status every 5 sec:
   - `GET /api/ppt/job-status/:jobId`
3. Optional abort:
   - `POST /api/ppt/abort-job/:jobId`

## 11. Health and Logs
- Service state:
```powershell
Get-Service PPTOptimizer
```
- Live logs (if running in console): PowerShell output
- For NSSM, configure stdout/stderr log files in NSSM service settings.
- For PM2 logs:
```powershell
pm2 logs ppt-optimizer
```

## 12. Update / Redeploy
```powershell
cd C:\apps\ppt-optimizer
git pull
npm ci
pm2 restart ppt-optimizer
```

## 13. Common Troubleshooting
- `Template not found`: verify `baseUrl + TemplatePath` or local `templates\fileName`
- `Job not found`: status TTL expiry (currently 1 hour)
- PowerPoint repair prompt: ensure latest `server.js` and restart service
- Large file timings: expected for 100+ slides, use async job polling
