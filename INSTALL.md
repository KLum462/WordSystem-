# Installation & Setup Guide

This guide provides step-by-step instructions for installing and running the Word Automation & ASP.NET Core 8 system from scratch.

---

## 1. Prerequisites Installation

### A. .NET SDK 8.0+
1. Download .NET SDK from [https://dotnet.microsoft.com/download](https://dotnet.microsoft.com/download).
2. Verify installation in PowerShell:
   ```powershell
   dotnet --version
   ```

### B. Microsoft SQL Server 2022
1. Download **SQL Server 2022 Express** from Microsoft's download page.
2. Run the installer and choose **Basic Installation**.
3. Confirm the default instance name `MSSQLSERVER` is running:
   ```powershell
   Get-Service -Name "MSSQLSERVER"
   ```

### C. SQL Server Management Studio (SSMS)
1. Download **SSMS** from [https://aka.ms/ssmsfullsetup](https://aka.ms/ssmsfullsetup).
2. Install and launch SSMS.
3. Connect using:
   - Server Name: `localhost`
   - Authentication: `Windows Authentication`

### D. Microsoft Word 2016
1. Ensure Microsoft Word 2016 (or Office 365 / Word 2019 / 2021) is installed on Windows.
2. Word COM Interop is registered automatically in the Global Assembly Cache (GAC).

---

## 2. Setting Up the Database

### A. Update Connection String (if necessary)
Open `WordServer.Api/appsettings.json`:
```json
{
  "ConnectionStrings": {
    "DefaultConnection": "Server=localhost;Database=WordSystemDb;Integrated Security=True;TrustServerCertificate=True;MultipleActiveResultSets=true",
    "SqliteConnection": "Data Source=WordSystem.db"
  },
  "DatabaseProvider": "SqlServer"
}
```

### B. Apply EF Core Migrations
Execute from the repository root:
```powershell
dotnet ef database update --project WordServer.Api/WordServer.Api.csproj --startup-project WordServer.Api/WordServer.Api.csproj
```
This creates the `WordSystemDb` database and the tables:
- `Documents`
- `Templates`
- `Metadata`
- `Settings`
- `Logs`

The server will automatically seed initial document `doc-001`, template `tpl-001`, and default configuration settings upon startup.

---

## 3. Building the Solution

Build all projects in the solution:
```powershell
dotnet build WordSystem.sln
```
The solution consists of:
- `WordServer.Api`: ASP.NET Core 8 Web API
- `WordClient.Core`: Shared library (.NET 8 & .NET Framework 4.8)
- `WordVstoAddIn`: Word 2016 VSTO Add-In (.NET Framework 4.8, x86)
- `WordClient.TestApp`: Desktop interactive test runner (.NET 8 WinForms)
- `WordServer.Tests`: Integration test suite (xUnit)

---

## 4. Running Integration Tests

Run the full integration test suite:
```powershell
dotnet test WordServer.Tests/WordServer.Tests.csproj --collect:"XPlat Code Coverage"
```
All 28 tests will execute against an isolated database using `WebApplicationFactory<Program>`, verifying >50% server-side code coverage.

---

## 5. Running the Application

### Step 1: Start the Backend Server
```powershell
dotnet run --project WordServer.Api/WordServer.Api.csproj --urls "http://localhost:5000"
```
- Interactive Swagger UI: [http://localhost:5000/swagger](http://localhost:5000/swagger)
- Health Check: [http://localhost:5000/api/health](http://localhost:5000/api/health)

### Step 2: Launch the Client Application
You can use either:
1. **Interactive Desktop Test Runner:**
   ```powershell
   dotnet run --project WordClient.TestApp/WordClient.TestApp.csproj
   ```
2. **Visual Studio VSTO Add-in:**
   Open `WordSystem.sln` in Visual Studio, set `WordVstoAddIn` as the startup project, and press **F5**. Microsoft Word 2016 will open with the custom ribbon tab **"Internship Word Tools"**.

---

## 6. Troubleshooting

| Issue | Cause | Solution |
|---|---|---|
| `Cannot connect to SQL Server` | MSSQLSERVER service stopped | Run `Start-Service MSSQLSERVER` in admin PowerShell |
| `HTTP 401 Unauthorized` | Invalid or missing API key | Ensure header `X-API-Key: secret-internship-api-key-2026` is sent |
| `HTTP 413 Payload Too Large` | Payload exceeds 100 MB | Ensure documents are within the 100 MB size limit |
| `Word COM Exception / No Active Document` | No document open in Word | Open or create a blank Word document before running text actions |
