# DataDownloader ContPaqi

Blazor Server web app that exports the "Movimientos, Auxiliares del Catálogo" account-ledger report from a CONTPAQi Contabilidad SQL Server database into an Excel file (.xlsx) matching CONTPAQi's native report layout.

## What it does

1. The user selects a company, a date range, and an optional account (GL) range.
2. The app queries the corresponding CONTPAQi SQL Server database (`Cuentas`, `MovimientosPoliza`, `TiposPolizas`, `SaldosCuentas`, `Ejercicios`) and builds the Excel workbook directly via EPPlus: one section per account, with opening balance, movement rows with a running balance, a per-account subtotal, and a grand total.
3. The finished file is held in memory (`JobStore`) and served through a one-time download link named `{COMPANY}_{yyyyMMdd_HHmmss}.xlsx`.

## Supported companies

| Company | Database |
|---|---|
| Servicios Tractomotrices | ctASERVICIOS_TRACTOMOTRICES_SA_DE_C |
| Centro Regional de Capacitacion | ctCENTRO_REGIONAL_DE_CAPACITACION_2 |
| SecureFleet Innovations | ctSECUREFLEET_INNOVATIONS |
| Trancasa Logistica de Mexico | ctTRANCASA_LOGISTICA_DE_MEXICO |

All databases share the same SQL Server instance and credentials.

## Stack

- .NET 10 / Blazor Server
- Microsoft.Data.SqlClient
- EPPlus 7 (NonCommercial license)
- DotNetEnv

## Setup

### 1. Configure environment variables

Copy `.env.example` to `.env` and fill in the values:

```
TCAMN_SERVER=<sql-server-ip>
TCAMN_DB_USER=<username>
TCAMN_DB_PASS=<password>
TCAMN_TRUSTED_SERVER_CERTIFICATE=Yes
```

> `.env` is git-ignored. Never commit it.

### 2. Run

```bash
dotnet run
```

The app starts at `http://localhost:5229`.

## Project structure

```
Components/Pages/Download.razor   # UI and download flow
Services/DataDownloaderService.cs # SQL query + Excel generation
Services/JobStore.cs              # In-memory file store (one-time links)
Models/DownloadParameters.cs      # Form model + company → database mapping
Program.cs                        # App bootstrap + /api/download endpoint
```
