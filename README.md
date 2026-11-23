# Word System: ASP.NET Core 8 & Microsoft Word 2016 VSTO Interop Solution

[![.NET 8](https://img.shields.io/badge/.NET-8.0-blue.svg)](https://dotnet.microsoft.com/)
[![.NET Framework](https://img.shields.io/badge/.NET%20Framework-4.8-purple.svg)](https://dotnet.microsoft.com/)
[![SQLite](https://img.shields.io/badge/SQLite-Default-003B57.svg)](https://www.sqlite.org/)
[![PostgreSQL](https://img.shields.io/badge/PostgreSQL-Supported-336791.svg)](https://www.postgresql.org/)
[![SQL Server 2022](https://img.shields.io/badge/SQL%20Server-Supported-red.svg)](https://www.microsoft.com/sql-server/)
[![EF Core 8](https://img.shields.io/badge/EF%20Core-8.0.13-green.svg)](https://learn.microsoft.com/ef/core/)
[![Tests](https://img.shields.io/badge/Integration%20Tests-43%2F43%20Passed%20(87.5%25%20cov)-brightgreen.svg)]()

Complete implementation of the **C# Developer Training Internship Assignment**. This repository contains an ASP.NET Core 8 backend server with Entity Framework Core 8 (SQLite default, PostgreSQL and SQL Server supported), accompanied by a Microsoft Word 2016 VSTO/COM Add-In (`.NET Framework 4.8`, `x86`) and interactive WinForms desktop test client using Microsoft Word COM Interop.

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
│   ├── App_Data/uploads/       # Stored .docx uploads (enforcing 100 MB / HTTP 413)
│   ├── Controllers/            # Health, Auth, Documents, Text, Templates, Management
│   ├── Data/                   # AppDbContext, Entities, DbInitializer
│   ├── Migrations/             # EF Core 8 schema migrations
│   ├── Middleware/             # ApiKeyMiddleware, ExceptionMiddleware
│   ├── Services/               # TextProcessingService
│   ├── appsettings.json        # DatabaseProvider (Sqlite/Postgres/SqlServer), Storage, Auth
│   └── Program.cs              # Serilog request/file logging, Swagger, Kestrel 100 MB limit
│
├── WordClient.Core/            # Shared Library (Multi-targeting net8.0 & net48)
│   ├── Configuration/          # ClientConfig (reads add-in dll.config / App.config)
│   ├── Models/                 # DocumentDto, TemplateDto, Upload & Check DTOs
│   └── Services/               # ApiClient (3 retries, 30s timeout), WordDocumentService, ClientLogger
│
├── WordVstoAddIn/              # Word 2016 VSTO/COM Add-in (.NET Framework 4.8, x86)
│   ├── Ribbon.cs / Ribbon.xml  # Custom Office Ribbon tab with 6 required buttons
│   ├── RibbonDialogs.cs        # Document/Template loader, Processing, Inspection + Find & Replace
│   ├── ComAddInConnect.cs      # Native Word COM Add-In entry point (IDTExtensibility2)
│   ├── ThisAddIn.cs            # Add-In lifecycle and COM release handlers
│   ├── Register-WordAddIn.ps1  # Auto-detects Office arch and registers add-in
│   └── App.config              # Client connection settings and API key (copied to dll.config)
│
├── WordClient.TestApp/         # Interactive WinForms Desktop Client (.NET 8 Windows Forms)
│   ├── Form1.cs                # GUI testing all operations, .docx Open/SaveAs2/Upload & live logs
│   └── App.config              # Connection settings
│
└── WordServer.Tests/           # Integration Test Project (.NET 8 xUnit)
    └── ServerIntegrationTests.cs # 43 integration tests (WebApplicationFactory + SQLite in-memory)
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

The server supports switching between database engines via the `DatabaseProvider` key in `WordServer.Api/appsettings.json`:

### SQLite (Default — zero-configuration file database):
```json
"DatabaseProvider": "Sqlite",
"ConnectionStrings": {
  "SqliteConnection": "Data Source=WordSystem.db"
}
```

### PostgreSQL (Supported out of the box):
```json
"DatabaseProvider": "Postgres",
"ConnectionStrings": {
  "PostgresConnection": "Host=localhost;Port=5432;Database=WordSystemDb;Username=postgres;Password=postgres"
}
```
(`Npgsql.EntityFrameworkCore.PostgreSQL` is pre-installed and configured with automatic retry on transient failures.)

### Microsoft SQL Server 2022 (Optional):
```json
"DatabaseProvider": "SqlServer",
"ConnectionStrings": {
  "DefaultConnection": "Server=localhost;Database=WordSystemDb;Integrated Security=True;TrustServerCertificate=True;"
}
```

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

### Configuring Authentication & HTTPS:
Authentication validates `X-API-Key` directly against active accounts in the `Users` database table (`IsActive = true`). Default seeded accounts:
- `admin` (`Admin` role) — `secret-internship-api-key-2026`
- `ivan_intern` (`Intern` role) — `internship-key-demo`

Users can be created and activated/deactivated dynamically via `POST /api/management/users` and `PATCH /api/management/users/{id}/status`.
To enable HTTPS redirection in production, set `"Security": { "EnableHttpsRedirection": true }` in `appsettings.json` (and run with `--urls "https://localhost:5001;http://localhost:5000"`).

---

## 7. API Endpoints Reference

| Method | Endpoint | Auth | Description |
|---|---|---|---|
| `GET` | `/api/health` | No | Server and database health check |
| `POST` | `/api/auth/validate` | Yes | Validates user API key against `Users` table, returns user & role |
| `GET` | `/api/documents` | Yes | Lists all documents with author & version summary |
| `GET` | `/api/documents/{id}` | Yes | Retrieves document JSON (`doc-001`) |
| `POST` | `/api/documents` | Yes | Creates document, metadata, and v1 snapshot in `DocumentVersions` |
| `PUT` | `/api/documents/{id}` | Yes | Updates document and creates next version in `DocumentVersions` |
| `GET` | `/api/text/operations` | Yes | Lists all 12 supported processing operations |
| `POST` | `/api/text/process` | Yes | Processes text (`uppercase`, `normalize`, `anonymize`, `extract_entities`, `summarize`, `stats`, etc.) and logs to `ProcessingHistory` |
| `POST` | `/api/documents/process` | Yes | Processes full document text (100 MB limit -> `413`) and logs to `ProcessingHistory` |
| `POST` | `/api/documents/upload` | Yes | Uploads `.docx` file to `App_Data/uploads` (100 MB limit -> `413`) |
| `GET` | `/api/documents/uploads` | Yes | Lists uploaded `.docx` files stored in `App_Data/uploads` |
| `GET` | `/api/documents/uploads/{fileName}` | Yes | Downloads an uploaded `.docx` file from `App_Data/uploads` |
| `POST` | `/api/documents/check` | Yes | Quality & PII inspection (empty paragraphs, double spaces, long sentences, emails/phones) |
| `GET` | `/api/templates` | Yes | Lists templates (supports `?categoryId=` filter) |
| `GET` | `/api/templates/{id}` | Yes | Retrieves template by ID (also `/api/template/{id}`) |
| `POST` | `/api/templates` | Yes | Creates new template with optional `categoryId` |
| `PUT` | `/api/templates/{id}` | Yes | Updates template by ID |
| `DELETE` | `/api/templates/{id}` | Yes | Deletes template by ID |
| `GET/POST/PATCH` | `/api/management/users` | Yes | Manage API users and active status (`/{id}/status`) |
| `GET/POST` | `/api/management/categories` | Yes | Manage template categories |
| `GET` | `/api/management/versions/{docId}` | Yes | View document version history |
| `GET/POST/PATCH` | `/api/management/comments` | Yes | View, create, and resolve document comments |
| `GET` | `/api/management/history` | Yes | Audit log of all text/document processing operations |
| `GET/PUT` | `/api/management/settings` | Yes | Read and update key-value system settings |

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
1. **"Authenticate"**: Sends API key (read from `WordVstoAddIn/App.config` / `WordVstoAddIn.dll.config`) to `/api/auth/validate`, verifying credentials, user role, and client log location.
2. **"Load Document"**: Generates a styled `.docx` from server document JSON (`doc-001`) or server template (with optional `SaveAs2` to `.docx`), or opens an existing local `.docx` file with lock & 100 MB validation.
3. **"Send Selected Text"**: Reads `Selection.Range.Text`, posts to `/api/text/process`, and inserts the result at the chosen target (Selected Range, End of Document, or Bookmark / Content Control).
4. **"Send Document"**: Validates 100 MB limit (both file size and UTF-8 text), sends document text to `/api/documents/process`, inserts result, and optionally uploads the `.docx` file to `App_Data/uploads` (`POST /api/documents/upload`).
5. **"Check Document"**: Shows document metadata (`Title`, `Author`, `CreationDate`), word/paragraph/empty-paragraph/table counts, Word spell-check errors, server quality issues, and includes an inline **Find & Replace...** action.
6. **"Save Template"**: Extracts active document paragraphs, styles, tables, font/size, and metadata and saves it to the server template store via `/api/templates`.

### Building, Debugging (Visual Studio F5) & Registering the Add-In:
- **Where to configure API Key & Server URL:** Edit `WordVstoAddIn/App.config` (copied automatically to `WordVstoAddIn.dll.config`).
- **Build (`x86` default per spec):**
  ```powershell
  dotnet build WordVstoAddIn/WordVstoAddIn.csproj
  ```
- **Register in Microsoft Word** (run PowerShell as Administrator; auto-detects x86/x64 Office or pass `-OfficeArch x86` / `-OfficeArch x64`):
  ```powershell
  .\WordVstoAddIn\Register-WordAddIn.ps1
  ```
- **Run/Debug from Visual Studio:** Open `WordSystem.sln`, set `WordVstoAddIn` as Startup Project, choose the matching Word launch profile (`Microsoft Word (x86)` or `Microsoft Word (x64)`), and press **F5**.
- **Unregister cleanly:**
  ```powershell
  .\WordVstoAddIn\Unregister-WordAddIn.ps1
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

Integration tests use `WebApplicationFactory<Program>`, `TestServer`, and **SQLite in-memory** (`DataSource=:memory:`) to achieve **87.51% line coverage** (63.65% branch coverage) of `WordServer.Api` server logic — well above the ≥ 50% requirement.

### Running Tests:
```powershell
dotnet test WordServer.Tests/WordServer.Tests.csproj --collect:"XPlat Code Coverage"
```

### Test Suite Summary:
- **Total Tests:** 43 passed (0 failed, 0 skipped)
- **Covered Scenarios:**
  - Health check endpoint
  - API Key authentication (valid DB users, roles, deactivation, invalid, missing)
  - Document retrieval (`doc-001` schema compliance), creation, versioning, and 404 responses
  - Text processing (`uppercase`, `lowercase`, `reverse`, `word_count`, `trim`, `capitalize`, `normalize`, `anonymize`, `extract_entities`, `summarize`, `stats`)
  - Document processing, paragraph statistics, and `413 Payload Too Large` limit enforcement
  - `.docx` file upload to `App_Data/uploads` (`POST /api/documents/upload`), listing, download, extension validation (`400`), and size limit (`413`)
  - Document checking (empty paragraphs, double spaces, PII detection)
  - Full CRUD lifecycle for Templates on both `/api/templates` and `/api/template/{id}`
  - Malformed JSON (`400`) and global `ExceptionMiddleware` (`400`, `413`, `500`) handling
  - Management endpoints (users, categories, comments, settings, processing history)
  - Client configuration loading from add-in `.config` file and client-side 100 MB enforcement

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
| 10 | Integration testing (>= 50% server coverage) | 43 xUnit tests with `WebApplicationFactory` & SQLite in-memory (87.5% server coverage) |
| 11 | Enforcement of 100 MB limit, `App_Data/uploads` & documentation | Client and server 100 MB / `413` limits, `App_Data/uploads`, `README.md`, `INSTALL.md` |
| 12 | Final demonstration & deliverables on `master` branch | Complete repository on Git `master` branch |
