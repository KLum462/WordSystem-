# Word System: решение на ASP.NET Core 8 и Microsoft Word 2016 VSTO/COM Interop

[![.NET 8](https://img.shields.io/badge/.NET-8.0-blue.svg)](https://dotnet.microsoft.com/)
[![.NET Framework](https://img.shields.io/badge/.NET%20Framework-4.8-purple.svg)](https://dotnet.microsoft.com/)
[![SQLite](https://img.shields.io/badge/SQLite-По%20умолчанию-003B57.svg)](https://www.sqlite.org/)
[![PostgreSQL](https://img.shields.io/badge/PostgreSQL-Поддерживается-336791.svg)](https://www.postgresql.org/)
[![SQL Server 2022](https://img.shields.io/badge/SQL%20Server-Поддерживается-red.svg)](https://www.microsoft.com/sql-server/)
[![EF Core 8](https://img.shields.io/badge/EF%20Core-8.0.13-green.svg)](https://learn.microsoft.com/ef/core/)
[![Tests](https://img.shields.io/badge/Интеграционные%20тесты-43%2F43%20пройдено%20(87.5%25%20покрытие)-brightgreen.svg)]()

Полная реализация **учебного задания на стажировку C#-разработчика**. Репозиторий содержит серверную часть на **ASP.NET Core 8** с **Entity Framework Core 8** (по умолчанию используется **SQLite**, также поддерживаются **PostgreSQL** и **SQL Server**), надстройку для **Microsoft Word 2016 VSTO/COM** (`.NET Framework 4.8`, `x86`) и интерактивный десктопный тестовый клиент на WinForms с использованием Microsoft Word COM Interop.

---

## Содержание

1. [Обзор проекта](#1-обзор-проекта)
2. [Архитектура решения](#2-архитектура-решения)
3. [Системные требования и зависимости](#3-системные-требования-и-зависимости)
4. [Настройка и управление базой данных](#4-настройка-и-управление-базой-данных)
5. [Переключение провайдера БД (SQLite, PostgreSQL, SQL Server)](#5-переключение-провайдера-бд)
6. [Настройка и запуск сервера](#6-настройка-и-запуск-сервера)
7. [Справочник API-эндпоинтов](#7-справочник-api-эндпоинтов)
8. [Надстройка Microsoft Word 2016 VSTO/COM](#8-надстройка-microsoft-word-2016-vstocom)
9. [Десктопное тестовое приложение (WordClient.TestApp)](#9-десктопное-тестовое-приложение)
10. [Интеграция с Word COM Interop и управление памятью](#10-интеграция-с-word-com-interop-и-управление-памятью)
11. [Обработка ошибок, таймауты, повторные попытки и логирование](#11-обработка-ошибок-таймауты-повторные-попытки-и-логирование)
12. [Интеграционные тесты и покрытие кода](#12-интеграционные-тесты-и-покрытие-кода)
13. [Сценарий итоговой демонстрации (10–15 минут)](#13-сценарий-итоговой-демонстрации-1015-минут)
14. [Соответствие 12-недельному плану стажировки](#14-соответствие-12-недельному-плану-стажировки)

---

## 1. Обзор проекта

Проект объединяет генерацию документов в **Microsoft Word** с веб-сервисами на **ASP.NET Core 8**. Реализованы следующие возможности:
- Защищённое взаимодействие с авторизацией по API-ключу (заголовок `X-API-Key`).
- Динамическое формирование документов `.docx` напрямую по структурированному JSON с сервера (`Heading1`, `Normal`, таблицы, шрифты и размеры).
- Двусторонняя обработка выделенного текста и всего документа (верхний/нижний регистр, реверс, подсчёт слов, нормализация, анонимизация и др.) с автоматической вставкой результата обратно в Word.
- Полный набор CRUD-операций для шаблонов документов с сохранением в БД через EF Core 8.
- Строгий контроль ограничения на размер файлов и запросов (до **100 МБ** с возвратом `413 Payload Too Large`).
- Корректное освобождение COM-объектов без зависших фоновых процессов `WINWORD.EXE`.

---

## 2. Архитектура решения

Решение (`WordSystem.sln`) разделено на независимые слои:

```text
WordSystem/
│
├── WordServer.Api/             # ASP.NET Core 8 Web API
│   ├── App_Data/uploads/       # Хранилище загруженных .docx-файлов (лимит 100 МБ / HTTP 413)
│   ├── Controllers/            # Health, Auth, Documents, Text, Templates, Management
│   ├── Data/                   # AppDbContext, Entities, DbInitializer
│   ├── Migrations/             # Миграции схемы EF Core 8
│   ├── Middleware/             # ApiKeyMiddleware, ExceptionMiddleware
│   ├── Services/               # TextProcessingService
│   ├── appsettings.json        # DatabaseProvider (Sqlite/Postgres/SqlServer), Storage, Auth
│   └── Program.cs              # Логирование запросов и файлов (Serilog), Swagger, Kestrel (100 МБ)
│
├── WordClient.Core/            # Общая библиотека (net8.0 и net48)
│   ├── Configuration/          # ClientConfig (чтение dll.config / App.config)
│   ├── Models/                 # DocumentDto, TemplateDto, модели загрузки и проверки
│   └── Services/               # ApiClient (3 попытки, таймаут 30 с), WordDocumentService, ClientLogger
│
├── WordVstoAddIn/              # Надстройка Word 2016 VSTO/COM (.NET Framework 4.8, x86)
│   ├── Ribbon.cs / Ribbon.xml  # Пользовательская вкладка ленты «Инструменты Word»
│   ├── RibbonDialogs.cs        # Диалоги загрузки документов/шаблонов, обработки, проверки и поиска/замены
│   ├── ComAddInConnect.cs      # Точка входа COM-надстройки Word (IDTExtensibility2)
│   ├── ThisAddIn.cs            # Жизненный цикл надстройки и освобождение COM
│   ├── Register-WordAddIn.ps1  # Автоопределение разрядности Office и регистрация надстройки
│   └── App.config              # Настройки подключения и API-ключ (копируется в dll.config)
│
├── WordClient.TestApp/         # Интерактивный WinForms-клиент (.NET 8 Windows Forms)
│   ├── Form1.cs                # Графический интерфейс для всех операций, Open/SaveAs2/Upload и живые логи
│   └── App.config              # Настройки подключения
│
└── WordServer.Tests/           # Проект интеграционных тестов (.NET 8 xUnit)
    └── ServerIntegrationTests.cs # 43 интеграционных теста (WebApplicationFactory + SQLite in-memory)
```

---

## 3. Системные требования и зависимости

### Окружение сервера:
- **Операционная система:** Windows 10 / Windows 11 / Windows Server
- **.NET SDK:** .NET SDK 8.0 или новее (поддерживается SDK 10.0)
- **СУБД:** SQLite (по умолчанию, не требует установки), PostgreSQL 14+ или Microsoft SQL Server 2022
- **Администрирование БД (при использовании SQL Server):** SQL Server Management Studio (SSMS 20 или 22)

### Окружение клиента:
- **Microsoft Word:** Microsoft Word 2016 (32-bit `x86` или 64-bit `x64` с поддержкой COM Interop)
- **.NET Framework:** .NET Framework 4.8 Developer Pack
- **Visual Studio:** Visual Studio 2022 или Visual Studio Community 2026 с компонентами:
  - «Разработка классических приложений .NET» (`.NET desktop development`)
  - «Разработка для Office и SharePoint» (`Office/SharePoint development`, для разработки VSTO-надстроек)

---

## 4. Настройка и управление базой данных

### А. Быстрый старт (SQLite по умолчанию)
При первом запуске `WordServer.Api` файл базы данных `WordSystem.db` создаётся и заполняется автоматически (`DbInitializer`) — установка отдельного сервера БД не требуется.

### Б. Использование Microsoft SQL Server 2022 и SSMS (опционально)
1. Установите **SQL Server 2022 Express** или **Developer Edition** (имя экземпляра по умолчанию `MSSQLSERVER`, режим проверки подлинности Windows).
2. Подключитесь в **SSMS** к серверу `localhost` (Windows Authentication).
3. Укажите `"DatabaseProvider": "SqlServer"` в `WordServer.Api/appsettings.json`.
4. Примените миграции EF Core:
```powershell
dotnet ef database update --project WordServer.Api/WordServer.Api.csproj --startup-project WordServer.Api/WordServer.Api.csproj
```

Схема базы данных включает все 5 обязательных таблиц (и 5 дополнительных):
- `Documents`: идентификатор, заголовок, JSON метаданных, JSON содержимого, JSON настроек, даты.
- `Templates`: идентификатор шаблона, название, описание, категория, JSON структуры документа, настроек и метаданных.
- `Metadata`: автор, версия, дата создания, описание.
- `Settings`: конфигурационные пары ключ–значение.
- `Logs`: журнал серверных событий и ошибок.
- Дополнительно: `Users`, `DocumentVersions`, `TemplateCategories`, `ProcessingHistory`, `DocumentComments`.

---

## 5. Переключение провайдера БД

Сервер поддерживает переключение между СУБД через параметр `DatabaseProvider` в файле `WordServer.Api/appsettings.json`:

### SQLite (по умолчанию — файловая БД без настройки):
```json
"DatabaseProvider": "Sqlite",
"ConnectionStrings": {
  "SqliteConnection": "Data Source=WordSystem.db"
}
```

### PostgreSQL (поддерживается из коробки):
```json
"DatabaseProvider": "Postgres",
"ConnectionStrings": {
  "PostgresConnection": "Host=localhost;Port=5432;Database=WordSystemDb;Username=postgres;Password=postgres"
}
```
*(Пакет `Npgsql.EntityFrameworkCore.PostgreSQL` предустановлен и настроен с автоматическим повтором при временных сбоях.)*

### Microsoft SQL Server 2022 (опционально):
```json
"DatabaseProvider": "SqlServer",
"ConnectionStrings": {
  "DefaultConnection": "Server=localhost;Database=WordSystemDb;Integrated Security=True;TrustServerCertificate=True;"
}
```

---

## 6. Настройка и запуск сервера

### Запуск сервера:
```powershell
dotnet run --project WordServer.Api/WordServer.Api.csproj --urls "http://localhost:5000"
```

После запуска доступны:
- Базовый адрес API: `http://localhost:5000`
- Документация Swagger UI: `http://localhost:5000/swagger`
- Проверка работоспособности: `http://localhost:5000/api/health`
- Ежедневные файлы логов сервера: `WordServer.Api/bin/Debug/net8.0/logs/server-YYYYMMDD.log`

### Настройка авторизации и HTTPS:
Авторизация проверяет заголовок `X-API-Key` по активным учётным записям в таблице `Users` (`IsActive = true`). Предустановленные ключи:
- `admin` (роль `Admin`) — `secret-internship-api-key-2026`
- `ivan_intern` (роль `Intern`) — `internship-key-demo`

Пользователей можно создавать и блокировать через `POST /api/management/users` и `PATCH /api/management/users/{id}/status`.  
Для включения перенаправления на HTTPS установите `"Security": { "EnableHttpsRedirection": true }` в `appsettings.json` (и запускайте с `--urls "https://localhost:5001;http://localhost:5000"`).

---

## 7. Справочник API-эндпоинтов

| Метод | Эндпоинт | Авторизация | Описание |
|---|---|---|---|
| `GET` | `/api/health` | Нет | Проверка состояния сервера и подключения к БД |
| `POST` | `/api/auth/validate` | Да | Проверка API-ключа по таблице `Users`, возврат пользователя и роли |
| `GET` | `/api/documents` | Да | Список всех документов с автором и версией |
| `GET` | `/api/documents/{id}` | Да | Получение JSON документа (`doc-001`) |
| `POST` | `/api/documents` | Да | Создание документа, метаданных и версии v1 в `DocumentVersions` |
| `PUT` | `/api/documents/{id}` | Да | Обновление документа и создание следующей версии в `DocumentVersions` |
| `GET` | `/api/text/operations` | Да | Список всех 12 поддерживаемых операций обработки текста |
| `POST` | `/api/text/process` | Да | Обработка текста (`uppercase`, `normalize`, `anonymize`, `extract_entities`, `summarize`, `stats` и др.) и запись в `ProcessingHistory` |
| `POST` | `/api/documents/process` | Да | Обработка текста всего документа (лимит 100 МБ → `413`) и запись в `ProcessingHistory` |
| `POST` | `/api/documents/upload` | Да | Загрузка файла `.docx` в `App_Data/uploads` (лимит 100 МБ → `413`) |
| `GET` | `/api/documents/uploads` | Да | Список загруженных `.docx`-файлов в `App_Data/uploads` |
| `GET` | `/api/documents/uploads/{fileName}` | Да | Скачивание загруженного `.docx`-файла из `App_Data/uploads` |
| `POST` | `/api/documents/check` | Да | Проверка качества и персональных данных (пустые абзацы, двойные пробелы, длинные предложения, email/телефоны) |
| `GET` | `/api/templates` | Да | Список шаблонов (поддерживается фильтр `?categoryId=`) |
| `GET` | `/api/templates/{id}` | Да | Получение шаблона по ID (также `/api/template/{id}`) |
| `POST` | `/api/templates` | Да | Создание нового шаблона с опциональной категорией `categoryId` |
| `PUT` | `/api/templates/{id}` | Да | Обновление шаблона по ID |
| `DELETE` | `/api/templates/{id}` | Да | Удаление шаблона по ID |
| `GET/POST/PATCH` | `/api/management/users` | Да | Управление пользователями API и статусом активности (`/{id}/status`) |
| `GET/POST` | `/api/management/categories` | Да | Управление категориями шаблонов |
| `GET` | `/api/management/versions/{docId}` | Да | Просмотр истории версий документа |
| `GET/POST/PATCH` | `/api/management/comments` | Да | Просмотр, создание и закрытие комментариев к документу |
| `GET` | `/api/management/history` | Да | Журнал аудита всех операций обработки текста и документов |
| `GET/PUT` | `/api/management/settings` | Да | Чтение и изменение системных настроек ключ–значение |

### Пример JSON-ответа документа (`GET /api/documents/doc-001`):
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

## 8. Надстройка Microsoft Word 2016 VSTO/COM

Проект надстройки находится в папке `WordVstoAddIn/`.

### Параметры проекта:
- **Тип проекта:** Надстройка Word 2016 VSTO / COM
- **Платформа:** `.NET Framework 4.8`
- **Целевая архитектура:** `x86` (поддерживается также сборка под `x64`)
- **Файл конфигурации:** `App.config` (копируется в `WordVstoAddIn.dll.config`)

### Кнопки на пользовательской ленте («Инструменты Word»):
1. **«Авторизация»**: отправляет API-ключ из `WordVstoAddIn/App.config` (`WordVstoAddIn.dll.config`) на `/api/auth/validate`, проверяя учётные данные, роль пользователя и путь к логу клиента.
2. **«Загрузить документ»**: формирует оформленный `.docx` из JSON документа с сервера (`doc-001`) или сохранённого шаблона (с опциональным сохранением в `.docx` через `SaveAs2`), либо открывает существующий локальный `.docx`-файл с проверкой блокировки и лимита 100 МБ.
3. **«Сохранить шаблон»**: извлекает из активного документа абзацы, стили, таблицы, шрифт/размер и метаданные и сохраняет шаблон на сервер через `/api/templates`.
4. **«Отправить выделенное»**: считывает `Selection.Range.Text`, отправляет на `/api/text/process` и вставляет ответ в выбранное место (выделенный диапазон, конец документа или закладка / Content Control).
5. **«Отправить документ»**: проверяет лимит 100 МБ (по размеру файла и по UTF-8 тексту), отправляет текст на `/api/documents/process`, вставляет результат и при необходимости загружает `.docx`-файл в `App_Data/uploads` (`POST /api/documents/upload`).
6. **«Проверить документ»**: отображает метаданные (`Заголовок`, `Автор`, `Дата создания`), количество слов/абзацев/пустых абзацев/таблиц, ошибки орфографии через Word Interop и замечания сервера.
7. **«Найти и заменить»**: открывает диалог поиска и замены текста по всему активному документу Word.

### Сборка, отладка (Visual Studio F5) и регистрация надстройки:
- **Где указать API-ключ и адрес сервера:** отредактируйте `WordVstoAddIn/App.config` (автоматически копируется в `WordVstoAddIn.dll.config`).
- **Сборка (`x86` по умолчанию согласно ТЗ):**
  ```powershell
  dotnet build WordVstoAddIn/WordVstoAddIn.csproj
  ```
- **Регистрация в Microsoft Word** (запустите PowerShell **от имени администратора**; скрипт сам определяет разрядность установленного Office или принимает `-OfficeArch x86` / `-OfficeArch x64`):
  ```powershell
  .\WordVstoAddIn\Register-WordAddIn.ps1
  ```
- **Запуск и отладка из Visual Studio:** откройте `WordSystem.sln`, назначьте `WordVstoAddIn` запускаемым проектом, выберите профиль (`Microsoft Word (x86)` или `Microsoft Word (x64)`) и нажмите **F5**.
- **Удаление регистрации:**
  ```powershell
  .\WordVstoAddIn\Unregister-WordAddIn.ps1
  ```

---

## 9. Десктопное тестовое приложение

Для быстрой проверки и демонстрации работы API и Word Interop без необходимости запускать надстройку внутри `WINWORD.EXE` в решение включён вспомогательный клиент `WordClient.TestApp`:

```powershell
dotnet run --project WordClient.TestApp/WordClient.TestApp.csproj
```

Возможности:
- Поля для настройки адреса сервера, API-ключа, ID документа, операции и места вставки результата.
- Кнопки для всех основных и вспомогательных операций (включая открытие/сохранение/загрузку `.docx`, работу с шаблонами и поиск/замену).
- Окно живого журнала событий (HTTP-запросы, повторные попытки, события Word COM).
- Прямая автоматизация Word Interop с подключением к запущенному Word или созданием нового процесса.

---

## 10. Интеграция с Word COM Interop и управление памятью

Клиентская часть строго соблюдает правила работы с Microsoft Word COM Interop:
- **Используемые объекты и API:**
  - `Microsoft.Office.Interop.Word.Application`
  - `Microsoft.Office.Interop.Word.Document` (и `ActiveDocument`)
  - `Selection.Range`
  - `Paragraphs`
  - `Range.Text`
  - `Bookmarks` и `ContentControls`
  - `Tables`
  - `Styles`
  - `SaveAs2` с `WdSaveFormat.wdFormatXMLDocument`
- **Дисциплина освобождения COM:**
  - Все ссылки на COM-объекты оборачиваются в блоки `try/finally` и освобождаются через `Marshal.ReleaseComObject(...)`.
  - При завершении выполняется принудительная сборка мусора (`GC.Collect()`, `GC.WaitForPendingFinalizers()`, `GC.Collect()`).
  - Исключается появление зависших фоновых процессов `WINWORD.EXE` в Диспетчере задач.
- **Обрабатываемые внештатные ситуации:**
  - Microsoft Word не запущен (автоматически запускается новый экземпляр Word).
  - Нет активного документа (показывается понятное сообщение пользователю).
  - Целевой файл занят другим процессом (проверка блокировки перед открытием/сохранением).
  - Размер файла или текста превышает 100 МБ (проверка и на клиенте, и на сервере).
  - Сервер недоступен (3 повторные попытки с задержкой и понятным сообщением об ошибке).
  - Неверный API-ключ (корректная обработка HTTP `401` без падения приложения).

---

## 11. Обработка ошибок, таймауты, повторные попытки и логирование

- **Политика повторов (Retry):** `ApiClient` автоматически выполняет до **3 попыток** с прогрессивной задержкой (1 с, 2 с) при сетевых сбоях или ответах сервера `5xx`.
- **Таймаут запросов:** настроен на **30 секунд** (`Timeout = TimeSpan.FromSeconds(30)`).
- **Логирование клиента:** все действия, запросы, повторы и события COM пишутся в `logs/client.log` (а при запуске внутри защищённой папки Office — в `%LocalAppData%\WordSystem\logs\client.log`).
- **Логирование сервера:** Serilog записывает все HTTP-запросы, авторизацию и ошибки в ежедневные файлы `logs/server-YYYYMMDD.log`, а критические ошибки дополнительно сохраняет в таблицу `Logs` в БД.
- **Коды ответов HTTP:**
  - `200 OK` / `201 Created` / `204 NoContent`
  - `400 Bad Request` (битый JSON, неизвестная операция, неверное расширение файла)
  - `401 Unauthorized` (отсутствующий, пустой или неверный `X-API-Key`)
  - `404 Not Found` (документ, шаблон или файл не найден)
  - `413 Payload Too Large` (размер текста или файла > 100 МБ)
  - `500 Internal Server Error` (перехват необработанных исключений без падения сервера)

---

## 12. Интеграционные тесты и покрытие кода

Интеграционные тесты используют `WebApplicationFactory<Program>`, `TestServer` и **SQLite in-memory** (`DataSource=:memory:`), обеспечивая **87.51% покрытия строк** (63.65% покрытия ветвей) серверной логики `WordServer.Api` — существенно выше требования ТЗ (≥ 50%).

### Запуск тестов:
```powershell
dotnet test WordServer.Tests/WordServer.Tests.csproj --collect:"XPlat Code Coverage"
```

### Состав тестового набора:
- **Всего тестов:** 43 пройдено (0 упавших, 0 пропущенных)
- **Проверяемые сценарии:**
  - Проверка работоспособности (`/api/health`)
  - Авторизация по API-ключу (активные пользователи БД, роли, деактивация, неверный и отсутствующий ключ)
  - Получение документов (`doc-001` по схеме ТЗ), создание, версионирование и ответ `404`
  - Обработка текста (`uppercase`, `lowercase`, `reverse`, `word_count`, `trim`, `capitalize`, `normalize`, `anonymize`, `extract_entities`, `summarize`, `stats`)
  - Обработка всего документа, статистика абзацев и проверка лимита `413 Payload Too Large`
  - Загрузка `.docx`-файлов в `App_Data/uploads` (`POST /api/documents/upload`), список, скачивание, проверка расширения (`400`) и превышения лимита (`413`)
  - Проверка документа (пустые абзацы, двойные пробелы, обнаружение персональных данных)
  - Полный CRUD-цикл для шаблонов по маршрутам `/api/templates` и `/api/template/{id}`
  - Обработка битого JSON (`400`) и глобальный `ExceptionMiddleware` (`400`, `413`, `500`)
  - Административные эндпоинты (пользователи, категории, комментарии, настройки, история обработки)
  - Чтение конфигурации клиента из `.config`-файла надстройки и клиентская проверка лимита 100 МБ

---

## 13. Сценарий итоговой демонстрации (10–15 минут)

1. **Шаг 1. Запуск сервера ASP.NET Core 8**
   - Выполните: `dotnet run --project WordServer.Api/WordServer.Api.csproj --urls "http://localhost:5000"`
   - Откройте в браузере `http://localhost:5000/swagger`, чтобы показать Swagger UI и `/api/health`.
2. **Шаг 2. Запуск Microsoft Word и надстройки**
   - Откройте Microsoft Word (или запустите `WordVstoAddIn` из Visual Studio по **F5**).
   - Перейдите на вкладку **«Инструменты Word»**.
3. **Шаг 3. Авторизация клиента**
   - Нажмите **«Авторизация»** — отобразится сообщение об успешной проверке ключа и имени пользователя.
4. **Шаг 4. Загрузка документа с сервера**
   - Нажмите **«Загрузить документ»** и выберите `doc-001` — в Word сформируется документ с заголовком `Heading1`, абзацами `Normal`, оформленной таблицей и шрифтом `Times New Roman`.
5. **Шаг 5. Отправка выделенного текста**
   - Выделите текст в Word и нажмите **«Отправить выделенное»** — сервер обработает фрагмент и вставит результат обратно в выделение.
6. **Шаг 6. Отправка всего документа и загрузка `.docx`**
   - Нажмите **«Отправить документ»** — выполняется проверка лимита 100 МБ, обработка текста и вставка ответа (при желании — сохранение `.docx` в `App_Data/uploads`).
7. **Шаг 7. Проверка документа, поиск/замена и сохранение шаблона**
   - Нажмите **«Проверить документ»** для просмотра метаданных, статистики, пустых абзацев и орфографии; используйте **«Найти и заменить»** для замены текста.
   - Нажмите **«Сохранить шаблон»**, чтобы сохранить структуру документа в БД на сервере.
8. **Шаг 8. Проверка освобождения COM**
   - Закройте Word / клиент и убедитесь в Диспетчере задач, что фоновых процессов `WINWORD.EXE` не осталось.

---

## 14. Соответствие 12-недельному плану стажировки

| Неделя | Тема | Статус реализации |
|---|---|---|
| 1 | Настройка окружения, VSTO hello world, кнопки на ленте | `WordVstoAddIn/Ribbon.cs`, `Ribbon.xml` (вкладка «Инструменты Word») |
| 2 | Основы Word Interop: создание, открытие, стили, таблицы, `SaveAs2` | `WordDocumentService.cs` (`GenerateDocumentFromDto`, `SaveActiveDocumentAs`, `OpenDocument`) |
| 3 | Сервер ASP.NET Core 8: health, API key middleware | `Program.cs`, `ApiKeyMiddleware.cs`, `HealthController.cs` |
| 4 | БД, EF Core, модели, базовые endpoints | `AppDbContext.cs`, `DbInitializer.cs`, 10 таблиц БД |
| 5 | Клиент: конфиг, `HttpClient`, авторизация | `ClientConfig.cs`, `ApiClient.cs` (`ValidateAuthDetailedAsync`) |
| 6 | `GET /api/documents/{id}` → генерация `.docx` по JSON | `DocumentsController.cs`, `WordDocumentService.cs` |
| 7 | Отправка выделенного текста и всего документа, вставка результата | `TextController.cs`, `InsertResult` (выделение / конец / закладка) |
| 8 | Шаблоны, метаданные, настройки | `TemplatesController.cs`, `TemplateEntity`, `MetadataEntity`, `SettingEntity` |
| 9 | Обработка ошибок, retry, timeout, логи | 3 попытки повтора, таймаут 30 с, Serilog и `ClientLogger` |
| 10 | Интеграционные тесты, покрытие ≥ 50% | 43 xUnit-теста (`WebApplicationFactory` + SQLite in-memory, покрытие 87.51%) |
| 11 | Ограничение 100 МБ, `App_Data/uploads`, документация | Лимиты 100 МБ / `413` на клиенте и сервере, `App_Data/uploads`, `README.md`, `INSTALL.md` |
| 12 | Демо и сдача репозитория в ветке `master` | Полный код загружен в ветку `master` на GitHub |
