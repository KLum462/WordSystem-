# Word System: ASP.NET Core 8 & Microsoft Word 2016 VSTO Interop Solution

[![.NET 8](https://img.shields.io/badge/.NET-8.0-blue.svg)](https://dotnet.microsoft.com/)
[![.NET Framework](https://img.shields.io/badge/.NET%20Framework-4.8%20(x86)-purple.svg)](https://dotnet.microsoft.com/)
[![SQL Server 2022](https://img.shields.io/badge/SQL%20Server-2022-red.svg)](https://www.microsoft.com/sql-server/)
[![EF Core 8](https://img.shields.io/badge/EF%20Core-8.0.13-green.svg)](https://learn.microsoft.com/ef/core/)
[![Tests](https://img.shields.io/badge/Integration%20Tests-28%2F28%20Passed-brightgreen.svg)]()

Complete enterprise implementation of the **C# Developer Training Internship Assignment**. This repository contains a production-grade backend server built on ASP.NET Core 8 with Entity Framework Core 8 and Microsoft SQL Server 2022, accompanied by a Microsoft Word 2016 VSTO Add-In and interactive test client using Microsoft Word COM Interop.

---

## Table of Contents

1. [Project Overview](#project-overview)
2. [Solution Architecture](#solution-architecture)
3. [System Requirements & Prerequisites](#system-requirements--prerequisites)
4. [Database Setup & Management (Microsoft SQL Server 2022 & SSMS)](#database-setup--management)
5. [Database Provider Switching (SQL Server, SQLite, PostgreSQL)](#database-provider-switching)
6. [Server Configuration & Execution](#server-configuration--execution)
7. [API Endpoints Reference](#api-endpoints-reference)
8. [Microsoft Word 2016 VSTO Add-In](#microsoft-word-2016-vsto-add-in)
9. [Word Desktop Test App (Interactive Runner)](#word-desktop-test-app)
10. [Word COM Interop Integration & Memory Management](#word-com-interop-integration--memory-management)
11. [Error Handling, Timeouts, Retries & Logging](#error-handling-timeouts-retries--logging)
12. [Integration Tests & Code Coverage](#integration-tests--code-coverage)
13. [Final Demonstration Walkthrough (10–15 Minutes)](#final-demonstration-walkthrough)
14. [Twelve-Week Internship Plan Alignment](#twelve-week-internship-plan-alignment)

---

## 1. Project Overview

The project bridges desktop document generation in **Microsoft Word** and modern web services in **ASP.NET Core 8**. It enables:
- Secure communication via API Key authentication (`X-API-Key` header).
- Dynamic document generation directly from structured server JSON data (`Heading1`, `Normal`, tables, fonts, sizes).
- Bidirectional text and document processing (e.g., uppercase, lowercase, reverse, word count) with automatic document insertion.
- Full CRUD template management on Microsoft SQL Server 2022 using EF Core 8 migrations.
- Strict enforcement of file and payload limits (up to **100 MB**).
- Reliable COM cleanup preventing orphaned `WINWORD.EXE` background processes.

---

## 2. Solution Architecture

The solution (`WordSystem.sln`) is organized into clean, decoupled layers:

```
d:\СТАЖИРОВКА C#\Prod\
│
├── WordServer.Api/             # ASP.NET Core 8 Web API
│   ├── Controllers/            # Health, Auth, Documents, Text, Templates
│   ├── Data/                   # AppDbContext, Entities, DbInitializer
│   ├── Migrations/             # EF Core 8 SQL Server schema migrations
│   ├── Middleware/             # ApiKeyMiddleware, ExceptionMiddleware
│   ├── appsettings.json        # Connection strings and API key config
│   └── Program.cs              # Serilog, Swagger, Kestrel 100MB body configuration
│
├── WordClient.Core/            # Shared Library (Multi-targeting net8.0 & net48)
│   ├── Configuration/          # ClientConfig (reads app.config / appsettings)
│   ├── Models/                 # DocumentDto, TemplateDto, Request/Response contracts
│   └── Services/               # ApiClient (3 retries, 30s timeout), WordDocumentService, ClientLogger
│
├── WordVstoAddIn/              # Word 2016 VSTO Add-in (.NET Framework 4.8, x86)
│   ├── Ribbon.cs / Ribbon.xml  # Custom Office Ribbon tab with 6 required buttons
│   ├── ThisAddIn.cs            # Add-In lifecycle and COM release handlers
│   └── App.config              # Client connection settings and API key
│
├── WordClient.TestApp/         # Interactive WinForms Desktop Client (.NET 8 Windows Forms)
│   ├── Form1.cs                # GUI testing all 6 operations with live logs
│   └── App.config              # Connection settings
│
└── WordServer.Tests/           # Integration Test Project (.NET 8 xUnit)
    └── ServerIntegrationTests.cs # 28 integration tests using WebApplicationFactory & TestServer
```

---

## 3. System Requirements & Prerequisites

### Server Environment:
- **Operating System:** Windows 10 / Windows 11 / Windows Server
- **.NET SDK:** .NET SDK 8.0 or later (SDK 10.0 supported)
- **Database:** Microsoft SQL Server 2022 (Express, Developer, or Standard edition)
- **Database Administration:** SQL Server Management Studio (SSMS 20 or 22)

### Client Environment:
- **Microsoft Word:** Microsoft Word 2016 (32-bit x86 or 64-bit with COM Interop)
- **.NET Framework:** .NET Framework 4.8 Developer Pack
- **Visual Studio:** Visual Studio 2022 or Visual Studio Community 2026 with:
  - ".NET desktop development" workload
  - "Office/SharePoint development" workload (for VSTO Add-In development)

---

## 4. Database Setup & Management

### A. Installing Microsoft SQL Server 2022
1. Download **SQL Server 2022 Express** or **Developer Edition** from Microsoft's official portal.
2. Select **Basic** or **Custom** installation.
3. Keep the default instance name: `MSSQLSERVER` (connects as `localhost` or `.`).
4. Ensure **Windows Authentication Mode** (or Mixed Mode) is enabled.

### B. Connecting via SQL Server Management Studio (SSMS)
1. Launch **SQL Server Management Studio 22**.
2. Set **Server type**: `Database Engine`.
3. Set **Server name**: `localhost` or `.` (or `localhost\SQLEXPRESS` if using named instance).
4. Set **Authentication**: `Windows Authentication`.
5. Click **Connect**.

### C. Database Connection String Configuration
In [appsettings.json](file:///d:/СТАЖИРОВКА%20C%23/Prod/WordServer.Api/appsettings.json):
```json
{
  "ConnectionStrings": {
    "DefaultConnection": "Server=localhost;Database=WordSystemDb;Integrated Security=True;TrustServerCertificate=True;MultipleActiveResultSets=true",
    "SqliteConnection": "Data Source=WordSystem.db"
  },
  "DatabaseProvider": "SqlServer"
}
```

### D. Applying EF Core Migrations
The initial migration `InitialCreate` is included in the repository. To apply migrations and update the database:
```powershell
dotnet ef database update --project WordServer.Api/WordServer.Api.csproj --startup-project WordServer.Api/WordServer.Api.csproj
```

The database schema includes the 5 required tables:
- `Documents`: Stores document ID, title, metadata JSON, content JSON, settings JSON, timestamps.
- `Templates`: Stores template ID, name, description, JSON document structures.
- `Metadata`: Stores author, version, timestamps, descriptions.
- `Settings`: Stores application key-value configuration.
- `Logs`: Records server activity and error telemetry.

---

## 5. Database Provider Switching

The server supports switching between database engines via the `DatabaseProvider` configuration key in `appsettings.json`:

### Microsoft SQL Server 2022 (Default):
```json
"DatabaseProvider": "SqlServer",
"ConnectionStrings": {
  "DefaultConnection": "Server=localhost;Database=WordSystemDb;Integrated Security=True;TrustServerCertificate=True;"
}
```

### SQLite (Zero-configuration file database):
```json
"DatabaseProvider": "Sqlite",
"ConnectionStrings": {
  "SqliteConnection": "Data Source=WordSystem.db"
}
```

### PostgreSQL (Optional):
Install `Npgsql.EntityFrameworkCore.PostgreSQL`, add `"NpgsqlConnection": "Host=localhost;Database=WordSystemDb;Username=postgres;Password=yourpass"`, and select `"DatabaseProvider": "Postgres"`.

---

## 6. Server Configuration & Execution

### Starting the Server:
```powershell
dotnet run --project WordServer.Api/WordServer.Api.csproj --urls "http://localhost:5000"
```

Once started:
- API Base URL: `http://localhost:5000`
- Swagger UI Documentation: `http://localhost:5000/swagger`
- Health check: `http://localhost:5000/api/health`
- Server log files are written daily to: `WordServer.Api/bin/Debug/net8.0/logs/server-YYYYMMDD.log`

### Configuring API Keys:
In `appsettings.json`:
```json
"Authentication": {
  "ApiKeys": [
    "secret-internship-api-key-2026",
    "internship-key-demo"
  ]
}
```
All protected endpoints require the HTTP header:
```http
X-API-Key: secret-internship-api-key-2026
```

---

## 7. API Endpoints Reference

| Method | Endpoint | Auth | Description |
|---|---|---|---|
| `GET` | `/api/health` | No | Server and SQL Server health check |
| `POST` | `/api/auth/validate` | Yes | Validates client API key |
| `GET` | `/api/documents/{id}` | Yes | Retrieves document JSON (`doc-001`) |
| `POST` | `/api/documents` | Yes | Creates and persists new document |
| `POST` | `/api/text/process` | Yes | Processes selected text (`uppercase`, `reverse`, etc.) |
| `POST` | `/api/documents/process` | Yes | Processes full document text (enforcing 100 MB limit) |
| `POST` | `/api/documents/check` | Yes | Validates words, paragraphs, tables, empty paragraphs |
| `GET` | `/api/templates` | Yes | Lists all document templates |
| `GET` | `/api/templates/{id}` | Yes | Retrieves template by ID (also `/api/template/{id}`) |
| `POST` | `/api/templates` | Yes | Creates new template (also `/api/template/`) |
| `PUT` | `/api/templates/{id}` | Yes | Updates template by ID |
| `DELETE` | `/api/templates/{id}` | Yes | Deletes template by ID |

### Document JSON Response Example (`GET /api/documents/doc-001`):
```json
{
  "id": "doc-001",
  "title": "Report",
  "metadata": {
    "author": "Ivan",
    "createdAt": "2026-10-07T10:00:00Z"
  },
  "content": {
    "paragraphs": [
      { "style": "Heading1", "text": "Heading" },
      { "style": "Normal", "text": "First paragraph" },
      { "style": "Normal", "text": "Second paragraph" }
    ],
    "tables": [
      {
        "name": "Table1",
        "rows": [
          ["Column1", "Column2"],
          ["Value1", "Value2"]
        ]
      }
    ]
  },
  "settings": {
    "font": "Times New Roman",
    "fontSize": 12
  }
}
```

---

## 8. Microsoft Word 2016 VSTO Add-In

The VSTO Add-in project is located in `WordVstoAddIn/`.

### Configuration:
- **Project Type:** Word 2016 VSTO Add-in
- **Framework:** .NET Framework 4.8
- **Platform:** x86
- **Configuration File:** `App.config`

### Required Ribbon Buttons:
The custom Ribbon tab **"Internship Word Tools"** exposes 6 buttons:
1. **"Authenticate"**: Sends API key to `/api/auth/validate`, verifying credentials and server status.
2. **"Load Document"**: Fetches `doc-001` JSON from `/api/documents/doc-001` and generates a styled `.docx` document in Word.
3. **"Send Selected Text"**: Reads `Selection.Range.Text`, posts to `/api/text/process`, and inserts the processed string back into the selection.
4. **"Send Document"**: Reads active document text, enforces 100 MB limit, sends to `/api/documents/process`, and inserts the result at document end.
5. **"Check Document"**: Inspects words, paragraphs, empty paragraphs, tables, and spell-check via Word Interop and `/api/documents/check`.
6. **"Save Template"**: Extracts active document structure and saves it to the SQL Server template store via `/api/templates`.

### Building the Add-In:
```powershell
dotnet build WordVstoAddIn/WordVstoAddIn.csproj
```

---

## 9. Word Desktop Test App

For instant testing and interactive demonstration without needing to attach to `WINWORD.EXE`, the solution includes `WordClient.TestApp`:

```powershell
dotnet run --project WordClient.TestApp/WordClient.TestApp.csproj
```

Features:
- GUI controls for Server URL, API Key, Document ID, Operation, and Insertion Location.
- Buttons for all 6 required actions.
- Real-time event log window displaying API requests, retry attempts, and Word COM events.
- Direct Word Interop automation with active instance detection and document generation.

---

## 10. Word COM Interop Integration & Memory Management

The client strictly adheres to Microsoft Word COM Interop best practices:
- **Objects & APIs Used:**
  - `Microsoft.Office.Interop.Word.Application`
  - `Microsoft.Office.Interop.Word.Document` (and `ActiveDocument`)
  - `Selection.Range`
  - `Paragraphs`
  - `Range.Text`
  - `Bookmarks`
  - `Tables`
  - `Styles`
  - `SaveAs2` with `WdSaveFormat.wdFormatXMLDocument`
- **COM Release Discipline:**
  - All COM object handles are wrapped in `try/finally` blocks and released using `Marshal.ReleaseComObject(...)`.
  - Garbage collection is forced (`GC.Collect()`, `GC.WaitForPendingFinalizers()`, `GC.Collect()`) upon completion.
  - Ensures no orphaned `WINWORD.EXE` background processes remain in Task Manager.
- **Handled Error Scenarios:**
  - Microsoft Word is not running (automatically starts a new Word process).
  - No active document exists (prompts user to open/create a document).
  - Target file locked or in use (detects file locks before opening).
  - File size exceeds 100 MB (validated on both client and server).
  - Server unavailable (3 retry attempts with exponential backoff before friendly error).
  - Invalid API key (HTTP 401 handled gracefully without application crash).

---

## 11. Error Handling, Timeouts, Retries & Logging

- **Retry Policy:** `ApiClient` implements an automatic retry mechanism (up to 3 attempts with progressive delay: 1s, 2s) on connection failures or server 5xx errors.
- **Request Timeout:** Configured to **30 seconds** (`Timeout = TimeSpan.FromSeconds(30)`).
- **Client Logging:** Logs every operation, request, retry attempt, and COM event to `logs/client.log`.
- **Server Logging:** Serilog writes daily rolling log files to `logs/server-YYYYMMDD.log` and logs errors to the SQL Server `Logs` table.
- **HTTP Status Codes:**
  - `200 OK` / `201 Created` / `204 NoContent`
  - `400 Bad Request` (Malformed JSON, invalid operations)
  - `401 Unauthorized` (Invalid or missing `X-API-Key`)
  - `404 Not Found` (Missing document or template ID)
  - `413 Payload Too Large` (Payloads > 100 MB)
  - `500 Internal Server Error` (Unhandled exceptions captured gracefully)

---

## 12. Integration Tests & Code Coverage

Integration tests use `WebApplicationFactory<Program>` and an isolated test database to achieve **>50% server logic coverage**.

### Running Tests:
```powershell
dotnet test WordServer.Tests/WordServer.Tests.csproj --collect:"XPlat Code Coverage"
```

### Test Suite Summary:
- **Total Tests:** 28 passed (0 failed, 0 skipped)
- **Covered Scenarios:**
  - Health check endpoint
  - API Key authentication (valid, invalid, missing)
  - Document retrieval (`doc-001` schema compliance)
  - Document 404 responses
  - Text processing (`uppercase`, `lowercase`, `reverse`, `word_count`, `trim`, `capitalize`)
  - Document processing and paragraph statistics
  - Document checking (empty paragraph detection, word counts, issues list)
  - Full CRUD lifecycle for Templates (`POST`, `GET by ID`, `GET All`, `PUT`, `DELETE`, `404 verification`)
  - Malformed JSON handling without server crash
  - Server-side 100 MB limit validation

---

## 13. Final Demonstration Walkthrough (10–15 Minutes)

Follow these steps during the final demonstration:

1. **Step 1: Start Microsoft SQL Server 2022 & Verify Database**
   - Open SSMS, connect to `localhost`.
   - Expand `Databases` -> `WordSystemDb` -> `Tables` (`Documents`, `Templates`, `Metadata`, `Settings`, `Logs`).
2. **Step 2: Start the ASP.NET Core 8 Server**
   - Run: `dotnet run --project WordServer.Api/WordServer.Api.csproj --urls "http://localhost:5000"`
   - Open browser at `http://localhost:5000/swagger` to demonstrate interactive Swagger UI and `/api/health`.
3. **Step 3: Launch Microsoft Word & Client**
   - Open Microsoft Word 2016 or launch `WordClient.TestApp` / Word VSTO Add-In.
4. **Step 4: Authenticate Client**
   - Click **"Authenticate"**. Observe successful verification dialog.
5. **Step 5: Load Document from Server**
   - Click **"Load Document"**. Observe Word opening a new document with Title Heading 1, paragraphs in Normal style, and formatted table populated from `doc-001`.
6. **Step 6: Send Selected Text**
   - Select text in Word and click **"Send Selected Text"**. Observe text being transformed on the server and replaced back into the document selection.
7. **Step 7: Send Entire Document**
   - Click **"Send Document"**. Observe 100 MB size validation, server processing, and the processed text appended at the end of the document.
8. **Step 8: Check Document & Save Template**
   - Click **"Check Document"** to inspect word counts, paragraph counts, and empty paragraphs.
   - Click **"Save Template"** to save the active document structure to SQL Server `Templates` table.
9. **Step 9: Verify COM Cleanup**
   - Close the client and observe in Task Manager that no orphaned `WINWORD.EXE` processes remain.

---

## 14. Twelve-Week Internship Plan Alignment

| Week | Topic / Deliverables | Implementation Status |
|---|---|---|
| 1 | Basic VSTO setup & Ribbon button creation | `WordVstoAddIn/Ribbon.cs`, `Ribbon.xml` with 6 buttons |
| 2 | Word Interop fundamentals (creating, styling, tables, SaveAs2) | `WordDocumentService.cs` (`GenerateDocumentFromDto`) |
| 3 | ASP.NET Core 8 server, health check, API key middleware | `Program.cs`, `ApiKeyMiddleware.cs`, `HealthController.cs` |
| 4 | SQL Server 2022 setup, EF Core 8, data models, migrations | `AppDbContext.cs`, `InitialCreate.cs`, `WordSystemDb` |
| 5 | Client configuration, HttpClient, authentication | `ClientConfig.cs`, `ApiClient.cs` (`ValidateAuthAsync`) |
| 6 | GET /api/documents/{id} & generating .docx from JSON | `DocumentsController.cs`, `WordDocumentService.cs` |
| 7 | Sending selected text & entire document; inserting results | `TextController.cs`, `InsertResult` at selection/end/bookmark |
| 8 | Document templates CRUD, metadata & settings | `TemplatesController.cs`, `TemplateEntity`, `MetadataEntity` |
| 9 | Error handling, retries, timeouts, file logging | 3-retry backoff, 30s timeout, Serilog & `ClientLogger` |
| 10 | Integration testing (>= 50% server coverage) | 28 xUnit tests with `WebApplicationFactory` (>51% coverage) |
| 11 | Enforcement of 100 MB limit & documentation | Client and server 100 MB limits, `README.md`, `INSTALL.md` |
| 12 | Final demonstration & deliverables on `master` branch | Complete repository on Git `master` branch |
