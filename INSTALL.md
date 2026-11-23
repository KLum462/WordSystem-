# Installation & Setup Guide

This guide provides step-by-step instructions for installing, configuring, and running the Word Automation & ASP.NET Core 8 system from scratch.

---

## 1. Prerequisites Installation

### A. .NET SDK 8.0+ & .NET Framework 4.8
1. Install **.NET SDK 8.0+** from [https://dotnet.microsoft.com/download](https://dotnet.microsoft.com/download).
2. Install **.NET Framework 4.8 Developer Pack** (included with Visual Studio desktop workload).
3. Verify SDK installation in PowerShell:
   ```powershell
   dotnet --version
   ```

### B. Database (SQLite Default, PostgreSQL or SQL Server Optional)
- **SQLite (Default):** Requires no external database server; `WordSystem.db` is created and seeded automatically on first server startup.
- **PostgreSQL (Optional):** Install PostgreSQL 14+ and create database `WordSystemDb` (or let EF Core create tables in an existing database).
- **Microsoft SQL Server 2022 (Optional):** Install SQL Server Express/Developer and ensure `MSSQLSERVER` service is running.

### C. Microsoft Word 2016+
1. Ensure **Microsoft Word 2016** (32-bit `x86` per spec, or 64-bit `x64`) is installed on Windows.
2. Microsoft Office Primary Interop Assemblies / COM automation are registered automatically with Office.

---

## 2. Database Configuration

Open `WordServer.Api/appsettings.json` and choose the desired `DatabaseProvider`:

### Option A: SQLite (Default — zero configuration)
```json
{
  "DatabaseProvider": "Sqlite",
  "ConnectionStrings": {
    "SqliteConnection": "Data Source=WordSystem.db"
  }
}
```

### Option B: PostgreSQL
```json
{
  "DatabaseProvider": "Postgres",
  "ConnectionStrings": {
    "PostgresConnection": "Host=localhost;Port=5432;Database=WordSystemDb;Username=postgres;Password=postgres"
  }
}
```

### Option C: Microsoft SQL Server 2022
```json
{
  "DatabaseProvider": "SqlServer",
  "ConnectionStrings": {
    "DefaultConnection": "Server=localhost;Database=WordSystemDb;Integrated Security=True;TrustServerCertificate=True;MultipleActiveResultSets=true"
  }
}
```
If using SQL Server with EF Core migrations, you can apply migrations via:
```powershell
dotnet ef database update --project WordServer.Api/WordServer.Api.csproj --startup-project WordServer.Api/WordServer.Api.csproj
```

On first startup (`DbInitializer`), the server creates all required tables (`Documents`, `Templates`, `Metadata`, `Settings`, `Logs`, plus `Users`, `DocumentVersions`, `TemplateCategories`, `ProcessingHistory`, `DocumentComments`), creates `WordServer.Api/App_Data/uploads`, and seeds sample document `doc-001`, template `tpl-001`, default settings, and demo API users:
- `admin` (`Admin`): `secret-internship-api-key-2026`
- `ivan_intern` (`Intern`): `internship-key-demo`

---

## 3. Configuring Client API Key & Server URL

- **Word VSTO/COM Add-In:** Edit `WordVstoAddIn/App.config` (copied automatically to `WordVstoAddIn.dll.config` in the build output):
  ```xml
  <appSettings>
    <add key="ServerUrl" value="http://localhost:5000" />
    <add key="ApiKey" value="secret-internship-api-key-2026" />
    <add key="TimeoutSeconds" value="30" />
    <add key="MaxRetries" value="3" />
    <add key="MaxFileSizeBytes" value="104857600" />
  </appSettings>
  ```
- **Desktop Test App:** Edit `WordClient.TestApp/App.config` (or change the values directly in the UI fields).

---

## 4. Building the Solution

Build all projects from the repository root:
```powershell
dotnet build WordSystem.sln
```
The solution consists of:
- `WordServer.Api`: ASP.NET Core 8 Web API (with `App_Data/uploads` and Serilog file/request logging)
- `WordClient.Core`: Shared library (`net8.0` & `net48`)
- `WordVstoAddIn`: Word 2016 VSTO/COM Add-In (`.NET Framework 4.8`, `x86` default; `x64` supported via `-p:Platform=x64`)
- `WordClient.TestApp`: Desktop interactive WinForms test client (`.NET 8`)
- `WordServer.Tests`: Integration test suite (`xUnit`, `WebApplicationFactory`, SQLite in-memory)

---

## 5. Running Integration Tests & Coverage

Run the full integration test suite with code coverage:
```powershell
dotnet test WordServer.Tests/WordServer.Tests.csproj --collect:"XPlat Code Coverage"
```
All **43 tests** execute against `TestServer` with an isolated **SQLite in-memory** database, achieving **87.5% line coverage** of `WordServer.Api` server logic (requirement: ≥ 50%).

---

## 6. Running the Application

### Step 1: Start the Backend Server
```powershell
dotnet run --project WordServer.Api/WordServer.Api.csproj --urls "http://localhost:5000"
```
- Interactive Swagger UI: [http://localhost:5000/swagger](http://localhost:5000/swagger)
- Health Check: [http://localhost:5000/api/health](http://localhost:5000/api/health)
- Uploaded `.docx` files directory: `WordServer.Api/App_Data/uploads/`
- Rolling server logs: `WordServer.Api/bin/Debug/net8.0/logs/server-YYYYMMDD.log`

### Step 2: Launch the Word Add-In or Desktop Client

1. **Microsoft Word Add-In (Ribbon tab "Internship Word Tools"):**
   - First register the COM Add-In once (run PowerShell **as Administrator**):
     ```powershell
     .\WordVstoAddIn\Register-WordAddIn.ps1
     ```
     (The script auto-detects installed Word architecture; you can also explicitly pass `-OfficeArch x86` or `-OfficeArch x64`.)
   - **Visual Studio F5 launch:** Open `WordSystem.sln`, set `WordVstoAddIn` as Startup Project, pick `Microsoft Word (x86)` (or `Microsoft Word (x64)`) in the debug profile dropdown, and press **F5**.
   - **Or start Word directly** — the custom tab **"Internship Word Tools"** appears automatically.
   - To unregister cleanly later: `.\WordVstoAddIn\Unregister-WordAddIn.ps1`.

2. **Interactive Desktop Test Runner (`WordClient.TestApp`):**
   ```powershell
   dotnet run --project WordClient.TestApp/WordClient.TestApp.csproj
   ```

---

## 7. Troubleshooting

| Issue | Cause | Solution |
|---|---|---|
| `HTTP 401 Unauthorized` | Invalid, missing, or deactivated API key | Set `ApiKey` in `App.config` to `secret-internship-api-key-2026` (or another active key from `Users` table) |
| `HTTP 413 Payload Too Large` | Document text or `.docx` upload exceeds 100 MB | Ensure document/file is within the 100 MB size limit |
| `Word COM Exception / No Active Document` | No document open in Word | Click **Load Document** or create/open a document in Word first |
| `RegAsm.exe` access denied | PowerShell not elevated | Run `Register-WordAddIn.ps1` in an Administrator PowerShell window |
| Add-In tab not visible in Word | Architecture mismatch (x86 vs x64 Office) | Re-run `.\WordVstoAddIn\Register-WordAddIn.ps1 -OfficeArch Auto` |
| Client logs location in Word | Office installation folder is read-only | Logs automatically fall back to `%LocalAppData%\WordSystem\logs\client.log` |
