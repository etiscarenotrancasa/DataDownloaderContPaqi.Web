using System.Data;
using System.Text;
using DataDownloaderContPaqi.Web.Models;
using Microsoft.Data.SqlClient;
using OfficeOpenXml;

namespace DataDownloaderContPaqi.Web.Services;

public class DataDownloaderService(IConfiguration config, ILogger<DataDownloaderService> logger)
{
    private static readonly string[] MesesEs =
        { "Ene", "Feb", "Mar", "Abr", "May", "Jun", "Jul", "Ago", "Sep", "Oct", "Nov", "Dic" };

    private static string SpanishDate(DateTime d) => $"{d:dd}/{MesesEs[d.Month - 1]}/{d:yyyy}";

    private static string CleanText(string? s) =>
        (s ?? "").Replace("\r", "").Replace("\n", " ").Trim();

    // CONTPAQi stores account codes unmasked (digits only) and displays them grouped
    // per the company's account mask (Parametros.Mascarilla). Without that mask available,
    // apply the common default 4-2-3 grouping (e.g. "617099000" -> "6170-99-000") only to
    // purely-numeric 9-digit codes; anything else is left unmodified.
    private static string FormatAccountCode(string codigo) =>
        codigo.Length == 9 && codigo.All(char.IsDigit)
            ? $"{codigo[..4]}-{codigo.Substring(4, 2)}-{codigo.Substring(6, 3)}"
            : codigo;

    private string BuildConnectionString(string dbName)
    {
        var baseCs = config.GetConnectionString("DefaultConnection")
            ?? throw new InvalidOperationException("ConnectionStrings:DefaultConnection not configured.");
        var sb = new SqlConnectionStringBuilder(baseCs) { InitialCatalog = dbName };
        return sb.ConnectionString;
    }

    private static string BuildDebugSql(SqlCommand cmd)
    {
        var sb = new StringBuilder();
        foreach (SqlParameter p in cmd.Parameters)
        {
            sb.Append($"DECLARE {p.ParameterName} {p.SqlDbType}");
            if (p.Size > 0) sb.Append($"({p.Size})");
            sb.Append(" = ");
            if (p.Value is null or DBNull)
                sb.AppendLine("NULL;");
            else if (p.Value is string s)
                sb.AppendLine($"N'{s.Replace("'", "''")}';");
            else if (p.Value is DateTime dt)
                sb.AppendLine($"'{dt:yyyy-MM-dd HH:mm:ss}';");
            else if (p.Value is bool b)
                sb.AppendLine(b ? "1;" : "0;");
            else
                sb.AppendLine($"{p.Value};");
        }
        sb.Append(cmd.CommandText);
        return sb.ToString();
    }

    public async Task<byte[]> DownloadAsync(
        DownloadParameters p,
        IProgress<DownloadProgress> progress,
        CancellationToken ct = default)
    {
        var dbName = p.DatabaseName
            ?? throw new ArgumentException("Invalid company selection.");

        var connStr = BuildConnectionString(dbName);

        ExcelPackage.LicenseContext = LicenseContext.NonCommercial;
        using var excel = new ExcelPackage();

        using var connection = new SqlConnection(connStr);
        await connection.OpenAsync(ct);

        return p.ReportType == ReportType.GLTransactions
            ? await DownloadGLTransactionsAsync(excel, connection, p, progress, logger, ct)
            : await DownloadAuxiliaresCatalogoAsync(excel, connection, dbName, p, progress, logger, ct);
    }

    // ======================================================================
    // Reporte: Movimientos, Auxiliares del Catálogo (agrupado por cuenta)
    // ======================================================================

    private static async Task<byte[]> DownloadAuxiliaresCatalogoAsync(
        ExcelPackage excel,
        SqlConnection connection,
        string dbName,
        DownloadParameters p,
        IProgress<DownloadProgress> progress,
        ILogger logger,
        CancellationToken ct)
    {
        string? glBegin = null, glEndBound = null;
        if (p.UseGLRange)
        {
            glBegin = p.GLBegin!.Trim();
            glEndBound = p.GLEnd!.Trim() + "￿"; // sentinel so prefix ranges include all sub-accounts
        }

        var beginDate = p.BeginDate.Date;
        var endDateExclusive = p.EndDate.Date.AddDays(1);

        var ws = excel.Workbook.Worksheets.Add("Movimientos");

        var razonSocial = await ExecuteScalarStringAsync(
            connection, "SELECT TOP 1 RazonSocial FROM dbo.Parametros;", ct) ?? dbName;
        var monedaNombre = await ExecuteScalarStringAsync(
            connection, "SELECT TOP 1 Nombre FROM dbo.Monedas ORDER BY Id;", ct) ?? "Peso Mexicano";

        var fecIniEje = await ResolveEjercicioStartAsync(connection, beginDate, logger, ct);

        var openingBalances = await LoadOpeningBalancesAsync(
            connection, fecIniEje, beginDate, glBegin, glEndBound, logger, ct);

        int totalRecords = await CountMovementsAsync(
            connection, beginDate, endDateExclusive, glBegin, glEndBound, logger, ct);

        progress.Report(new DownloadProgress(0, totalRecords));

        if (totalRecords == 0)
            return Array.Empty<byte>();

        WriteReportHeader(ws, razonSocial, monedaNombre, p.BeginDate, p.EndDate);

        int row = 9; // row 8 is the fixed spacer before the first account
        string? currentCodigo = null;
        decimal runningBalance = 0, acctCargo = 0, acctAbono = 0;
        decimal grandCargo = 0, grandAbono = 0, grandBalance = 0;
        int fetched = 0;

        const string movementSql = @"
            SELECT c.Codigo, c.Nombre, mp.Fecha, ISNULL(tp.Nombre, CAST(mp.TipoPol AS nvarchar(20))), mp.Folio,
                   mp.Concepto, mp.Referencia, mp.TipoMovto, mp.Importe
            FROM dbo.MovimientosPoliza mp
            INNER JOIN dbo.Cuentas c      ON c.Id = mp.IdCuenta
            LEFT JOIN dbo.TiposPolizas tp ON tp.Id = mp.TipoPol
            WHERE mp.Fecha >= @BeginDate AND mp.Fecha < @EndDateExclusive
              AND (@GLBegin IS NULL OR c.Codigo >= @GLBegin)
              AND (@GLEnd   IS NULL OR c.Codigo <  @GLEnd)
            ORDER BY c.Codigo, mp.Fecha, mp.Folio, mp.NumMovto;";

        using (var cmd = new SqlCommand(movementSql, connection) { CommandTimeout = 300 })
        {
            cmd.Parameters.Add("@BeginDate", SqlDbType.DateTime).Value = beginDate;
            cmd.Parameters.Add("@EndDateExclusive", SqlDbType.DateTime).Value = endDateExclusive;
            cmd.Parameters.Add("@GLBegin", SqlDbType.NVarChar, 60).Value = (object?)glBegin ?? DBNull.Value;
            cmd.Parameters.Add("@GLEnd", SqlDbType.NVarChar, 60).Value = (object?)glEndBound ?? DBNull.Value;

            logger.LogDebug("SQL [MovimientosPoliza]:\n{Sql}", BuildDebugSql(cmd));
            using var rdr = await cmd.ExecuteReaderAsync(ct);

            while (await rdr.ReadAsync(ct))
            {
                ct.ThrowIfCancellationRequested();

                var codigo = rdr.GetString(0);

                if (codigo != currentCodigo)
                {
                    if (currentCodigo is not null)
                    {
                        row = WriteAccountTotalRow(ws, row, acctCargo, acctAbono, runningBalance);
                        grandBalance += runningBalance;
                        row += 2; // spacer between account sections
                    }

                    currentCodigo = codigo;
                    runningBalance = openingBalances.GetValueOrDefault(codigo, 0m);
                    acctCargo = 0;
                    acctAbono = 0;
                    row = WriteAccountHeaderRow(ws, row, codigo, rdr.GetString(1), runningBalance);
                }

                var fecha = rdr.GetDateTime(2);
                var tipoPoliza = rdr.GetString(3);
                var folio = rdr.GetInt32(4);
                var concepto = CleanText(rdr.IsDBNull(5) ? null : rdr.GetString(5));
                var referencia = CleanText(rdr.IsDBNull(6) ? null : rdr.GetString(6));
                bool esCargo = !rdr.IsDBNull(7) && !rdr.GetBoolean(7); // TipoMovto: 0=Cargo, 1=Abono
                decimal importe = rdr.IsDBNull(8) ? 0 : (decimal)rdr.GetDouble(8);

                runningBalance += esCargo ? importe : -importe;
                if (esCargo) { acctCargo += importe; grandCargo += importe; }
                else { acctAbono += importe; grandAbono += importe; }

                row = WriteMovementRow(
                    ws, row, fecha, tipoPoliza, folio, concepto, referencia,
                    esCargo ? importe : (decimal?)null,
                    esCargo ? (decimal?)null : importe,
                    runningBalance);

                fetched++;
                if (fetched % 200 == 0)
                    progress.Report(new DownloadProgress(fetched, totalRecords));
            }
        }

        if (currentCodigo is not null)
        {
            row = WriteAccountTotalRow(ws, row, acctCargo, acctAbono, runningBalance);
            grandBalance += runningBalance;
            row += 3; // spacer before the grand total
        }

        WriteGrandTotalRow(ws, row, grandCargo, grandAbono, grandBalance);
        progress.Report(new DownloadProgress(totalRecords, totalRecords));

        ws.Column(1).Width = 37; ws.Column(2).Width = 31; ws.Column(3).Width = 8; ws.Column(4).Width = 47;
        ws.Column(5).Width = 20; ws.Column(6).Width = 14; ws.Column(7).Width = 14; ws.Column(8).Width = 16;

        return excel.GetAsByteArray();
    }

    private static async Task<string?> ExecuteScalarStringAsync(SqlConnection conn, string sql, CancellationToken ct)
    {
        using var cmd = new SqlCommand(sql, conn);
        var result = await cmd.ExecuteScalarAsync(ct);
        return result is null or DBNull ? null : result.ToString();
    }

    private static async Task<DateTime> ResolveEjercicioStartAsync(SqlConnection conn, DateTime beginDate, ILogger logger, CancellationToken ct)
    {
        const string sql = @"
            SELECT TOP 1 FecIniEje
            FROM dbo.Ejercicios
            WHERE @BeginDate BETWEEN FecIniEje AND FecFinEje
            ORDER BY Ejercicio DESC;";

        using var cmd = new SqlCommand(sql, conn);
        cmd.Parameters.Add("@BeginDate", SqlDbType.DateTime).Value = beginDate;

        logger.LogDebug("SQL [ResolveEjercicioStart]:\n{Sql}", BuildDebugSql(cmd));
        var result = await cmd.ExecuteScalarAsync(ct);
        if (result is null or DBNull)
            throw new InvalidOperationException(
                $"No fiscal year (Ejercicios) is configured that covers {beginDate:yyyy-MM-dd}.");

        return (DateTime)result;
    }

    private static async Task<Dictionary<string, decimal>> LoadOpeningBalancesAsync(
        SqlConnection conn, DateTime fecIniEje, DateTime beginDate,
        string? glBegin, string? glEndBound, ILogger logger, CancellationToken ct)
    {
        const string sql = @"
            SELECT
                c.Codigo,
                ISNULL(sc.SaldoIni, 0)
                  + ISNULL(SUM(CASE WHEN mp.TipoMovto = 0 THEN mp.Importe
                                    WHEN mp.TipoMovto = 1 THEN -mp.Importe
                                    ELSE 0 END), 0) AS SaldoInicial
            FROM dbo.Cuentas c
            LEFT JOIN dbo.SaldosCuentas sc
                   ON sc.IdCuenta = c.Id AND sc.Ejercicio = YEAR(@FecIniEje) AND sc.Tipo = 1
            LEFT JOIN dbo.MovimientosPoliza mp
                   ON mp.IdCuenta = c.Id
                  AND mp.Fecha >= @FecIniEje
                  AND mp.Fecha <  @BeginDate
            WHERE (@GLBegin IS NULL OR c.Codigo >= @GLBegin)
              AND (@GLEnd   IS NULL OR c.Codigo <  @GLEnd)
            GROUP BY c.Codigo, sc.SaldoIni;";

        using var cmd = new SqlCommand(sql, conn) { CommandTimeout = 300 };
        cmd.Parameters.Add("@FecIniEje", SqlDbType.DateTime).Value = fecIniEje;
        cmd.Parameters.Add("@BeginDate", SqlDbType.DateTime).Value = beginDate;
        cmd.Parameters.Add("@GLBegin", SqlDbType.NVarChar, 60).Value = (object?)glBegin ?? DBNull.Value;
        cmd.Parameters.Add("@GLEnd", SqlDbType.NVarChar, 60).Value = (object?)glEndBound ?? DBNull.Value;

        logger.LogDebug("SQL [LoadOpeningBalances]:\n{Sql}", BuildDebugSql(cmd));
        var result = new Dictionary<string, decimal>();
        using var rdr = await cmd.ExecuteReaderAsync(ct);
        while (await rdr.ReadAsync(ct))
            result[rdr.GetString(0)] = (decimal)rdr.GetDouble(1);

        return result;
    }

    private static async Task<int> CountMovementsAsync(
        SqlConnection conn, DateTime beginDate, DateTime endDateExclusive,
        string? glBegin, string? glEndBound, ILogger logger, CancellationToken ct)
    {
        const string sql = @"
            SELECT COUNT(*)
            FROM dbo.MovimientosPoliza mp
            INNER JOIN dbo.Cuentas c ON c.Id = mp.IdCuenta
            WHERE mp.Fecha >= @BeginDate AND mp.Fecha < @EndDateExclusive
              AND (@GLBegin IS NULL OR c.Codigo >= @GLBegin)
              AND (@GLEnd   IS NULL OR c.Codigo <  @GLEnd);";

        using var cmd = new SqlCommand(sql, conn) { CommandTimeout = 300 };
        cmd.Parameters.Add("@BeginDate", SqlDbType.DateTime).Value = beginDate;
        cmd.Parameters.Add("@EndDateExclusive", SqlDbType.DateTime).Value = endDateExclusive;
        cmd.Parameters.Add("@GLBegin", SqlDbType.NVarChar, 60).Value = (object?)glBegin ?? DBNull.Value;
        cmd.Parameters.Add("@GLEnd", SqlDbType.NVarChar, 60).Value = (object?)glEndBound ?? DBNull.Value;

        logger.LogDebug("SQL [CountMovements]:\n{Sql}", BuildDebugSql(cmd));
        return Convert.ToInt32(await cmd.ExecuteScalarAsync(ct));
    }

    private static void WriteReportHeader(ExcelWorksheet ws, string razonSocial, string moneda, DateTime begin, DateTime end)
    {
        ws.Cells[1, 1].Value = "CONTPAQ i";
        ws.Cells[1, 4].Value = razonSocial;
        ws.Cells[1, 8].Value = "Hoja:      1";
        ws.Cells[2, 1].Value = "Movimientos, Auxiliares del Catálogo";
        ws.Cells[2, 8].Value = $"Fecha: {SpanishDate(DateTime.Today)}";
        ws.Cells[3, 1].Value = $"del {SpanishDate(begin)} al {SpanishDate(end)}";
        ws.Cells[4, 1].Value = $"Moneda: {moneda}";
        ws.Cells[1, 1, 4, 8].Style.Font.Size = 12;

        ws.Cells[6, 1].Value = "C u e n ta";
        ws.Cells[6, 2].Value = "N o m b r e";
        ws.Cells[6, 8].Value = "Saldo Inicial";
        ws.Cells[7, 1].Value = "Fecha";
        ws.Cells[7, 2].Value = "Tipo";
        ws.Cells[7, 3].Value = "Número ";
        ws.Cells[7, 4].Value = "Concepto";
        ws.Cells[7, 5].Value = "Referencia";
        ws.Cells[7, 6].Value = "Cargos";
        ws.Cells[7, 7].Value = "Abonos";
        ws.Cells[7, 8].Value = "Saldo";
        ws.Cells[6, 1, 7, 8].Style.Font.Size = 10;
    }

    private static int WriteAccountHeaderRow(ExcelWorksheet ws, int row, string codigo, string nombre, decimal saldoInicial)
    {
        ws.Cells[row, 1].Value = FormatAccountCode(codigo);
        ws.Cells[row, 1].Style.Font.Bold = true;
        ws.Cells[row, 2].Value = nombre;
        ws.Cells[row, 2].Style.Font.Bold = true;
        ws.Cells[row, 7].Value = "Saldo inicial :";
        ws.Cells[row, 8].Value = saldoInicial;
        ws.Cells[row, 8].Style.Numberformat.Format = "#,##0.00";
        ws.Cells[row, 1, row, 8].Style.Font.Size = 9;
        return row + 1;
    }

    private static int WriteMovementRow(
        ExcelWorksheet ws, int row, DateTime fecha, string tipoPoliza, int folio,
        string concepto, string referencia, decimal? cargo, decimal? abono, decimal saldo)
    {
        ws.Cells[row, 1].Value = SpanishDate(fecha);
        ws.Cells[row, 2].Value = tipoPoliza;
        ws.Cells[row, 3].Value = folio;
        ws.Cells[row, 3].Style.Numberformat.Format = "#,##0";
        ws.Cells[row, 4].Value = concepto;
        ws.Cells[row, 5].Value = referencia;
        if (cargo is not null)
        {
            ws.Cells[row, 6].Value = cargo;
            ws.Cells[row, 6].Style.Numberformat.Format = "#,##0.00";
        }
        if (abono is not null)
        {
            ws.Cells[row, 7].Value = abono;
            ws.Cells[row, 7].Style.Numberformat.Format = "#,##0.00";
        }
        ws.Cells[row, 8].Value = saldo;
        ws.Cells[row, 8].Style.Numberformat.Format = "#,##0.00";
        ws.Cells[row, 1, row, 8].Style.Font.Size = 9;
        return row + 1;
    }

    private static int WriteAccountTotalRow(ExcelWorksheet ws, int row, decimal cargo, decimal abono, decimal saldo)
    {
        ws.Cells[row, 5].Value = "Total:";
        ws.Cells[row, 6].Value = cargo;
        ws.Cells[row, 7].Value = abono;
        ws.Cells[row, 8].Value = saldo;
        ws.Cells[row, 6, row, 8].Style.Numberformat.Format = "#,##0.00";
        ws.Cells[row, 6, row, 8].Style.Font.Bold = true;
        ws.Cells[row, 1, row, 8].Style.Font.Size = 9;
        return row + 1;
    }

    private static void WriteGrandTotalRow(ExcelWorksheet ws, int row, decimal cargo, decimal abono, decimal saldo)
    {
        ws.Cells[row, 5].Value = "T o t a l: ";
        ws.Cells[row, 6].Value = cargo;
        ws.Cells[row, 7].Value = abono;
        ws.Cells[row, 8].Value = saldo;
        ws.Cells[row, 6, row, 8].Style.Numberformat.Format = "#,##0.00";
        ws.Cells[row, 5, row, 8].Style.Font.Bold = true;
        ws.Cells[row, 1, row, 8].Style.Font.Size = 9;
    }

    // ======================================================================
    // Reporte: GL Transactions (homologado a GP View_GL_Trx)
    //
    // Ver "GP View_GL_Trx.txt" para las reglas de negocio validadas
    // (persona/asiento solo se resuelven cuando son únicos por movimiento,
    // Originating_Doc_Number siempre NULL, Exchange_Rate derivado, etc.).
    // ======================================================================

    private static readonly string[] GLTransactionsColumns =
    {
        "Trx_Status", "Trx_Date", "Journal_Entry", "Account_Number", "Account_Description",
        "Debit_Amount", "Credit_Amount", "Description", "Reference", "Source_Document",
        "Originating_Trx_Source", "Originating_Master_Id", "Originating_Master_Name", "Originating_Doc_Number",
        "Currency_Id", "Last_User", "User_Who_Posted", "Batch_Number", "Series", "Unique_Id",
        "Post_Date", "Originating_Debit_Amount", "Originating_Credit_Amount", "Exchange_Rate", "Currency_Description",
    };

    private const string GLTransactionsCte = @"
        ;WITH BaseMovimientos AS
        (
            SELECT
                m.Id            AS IdMovimiento,
                m.Guid          AS GuidMovimiento,
                m.NumMovto,
                m.TipoMovto,
                m.Importe,
                m.ImporteME,
                m.Referencia,
                m.Concepto,

                p.Id            AS IdPoliza,
                p.Guid          AS GuidPoliza,
                p.Ejercicio,
                p.Periodo,
                p.TipoPol,
                p.Folio,
                p.Fecha,
                p.SistOrig,
                p.IdUsuario,

                c.Id            AS IdCuenta,
                c.Codigo        AS CodigoCuenta,
                c.Nombre        AS NombreCuenta,
                c.IdMoneda

            FROM dbo.MovimientosPoliza m
            INNER JOIN dbo.Polizas p ON p.Id = m.IdPoliza
            INNER JOIN dbo.Cuentas c ON c.Id = m.IdCuenta

            WHERE
                p.Fecha >= @BeginDate
                AND p.Fecha < @EndDateExclusive
        ),

        MovimientoDocumento AS
        (
            SELECT
                b.IdMovimiento,

                COUNT(DISTINCT NULLIF(LTRIM(RTRIM(d.CodigoPersona)), '')) AS PersonasDistintas,

                CASE
                    WHEN COUNT(DISTINCT NULLIF(LTRIM(RTRIM(d.CodigoPersona)), '')) = 1
                        THEN MIN(LTRIM(RTRIM(per.Codigo)))
                    ELSE NULL
                END AS PersonaCodigo,

                CASE
                    WHEN COUNT(DISTINCT NULLIF(LTRIM(RTRIM(d.CodigoPersona)), '')) = 1
                        THEN MIN(per.Nombre)
                    ELSE NULL
                END AS PersonaNombre,

                COUNT(DISTINCT NULLIF(LTRIM(RTRIM(d.CodigoAsiento)), '')) AS AsientosDistintos,

                CASE
                    WHEN COUNT(DISTINCT NULLIF(LTRIM(RTRIM(d.CodigoAsiento)), '')) = 1
                        THEN MIN(LTRIM(RTRIM(a.Codigo)))
                    ELSE NULL
                END AS AsientoCodigo,

                CASE
                    WHEN COUNT(DISTINCT NULLIF(LTRIM(RTRIM(d.CodigoAsiento)), '')) = 1
                        THEN MIN(a.Nombre)
                    ELSE NULL
                END AS AsientoNombre

            FROM BaseMovimientos b
            LEFT JOIN dbo.AsocCFDIs cf ON cf.GuidRef = b.GuidMovimiento
            LEFT JOIN dbo.DocumentosAdministrativos d ON d.UUID = cf.UUID
            LEFT JOIN dbo.Personas per
                ON LTRIM(RTRIM(per.Codigo)) COLLATE DATABASE_DEFAULT = LTRIM(RTRIM(d.CodigoPersona)) COLLATE DATABASE_DEFAULT
            LEFT JOIN dbo.Asientos a
                ON LTRIM(RTRIM(a.Codigo)) COLLATE DATABASE_DEFAULT = LTRIM(RTRIM(d.CodigoAsiento)) COLLATE DATABASE_DEFAULT

            GROUP BY b.IdMovimiento
        )
        ";

    private const string GLTransactionsFilterSql = @"
        WHERE (@GLBegin IS NULL OR b.CodigoCuenta >= @GLBegin)
          AND (@GLEnd   IS NULL OR b.CodigoCuenta <  @GLEnd)
          AND (@JournalEntry IS NULL OR CONCAT(b.Ejercicio, '-', RIGHT('00' + CAST(b.Periodo AS varchar(2)), 2), '-', b.TipoPol, '-', b.Folio) LIKE '%' + @JournalEntry + '%')
          AND (@OriginalMasterName IS NULL OR md.PersonaNombre LIKE '%' + @OriginalMasterName + '%')
        ";

    private const string GLTransactionsCountSql = @"
        SELECT COUNT(*)
        FROM BaseMovimientos b
        LEFT JOIN MovimientoDocumento md ON md.IdMovimiento = b.IdMovimiento
        " + GLTransactionsFilterSql + ";";

    private const string GLTransactionsSelectSql = @"
        SELECT
            'Poliza' AS Trx_Status,
            CONVERT(char(10), b.Fecha, 105) AS Trx_Date,
            b.Folio AS Journal_Entry,
            LTRIM(RTRIM(b.CodigoCuenta)) AS Account_Number,
            b.NombreCuenta AS Account_Description,
            CASE WHEN b.TipoMovto = 0 THEN b.Importe ELSE 0 END AS Debit_Amount,
            CASE WHEN b.TipoMovto = 1 THEN b.Importe ELSE 0 END AS Credit_Amount,
            b.Concepto AS [Description],
            b.Referencia AS Reference,
            md.AsientoNombre AS Source_Document,
            CAST(b.SistOrig AS varchar(20)) AS Originating_Trx_Source,
            md.PersonaCodigo AS Originating_Master_Id,
            md.PersonaNombre AS Originating_Master_Name,
            CAST(NULL AS varchar(100)) AS Originating_Doc_Number,
            mon.CodigoSAT AS Currency_Id,
            u.Codigo AS Last_User,
            CAST(NULL AS varchar(100)) AS User_Who_Posted,
            CAST(NULL AS varchar(100)) AS Batch_Number,
            CAST(NULL AS varchar(100)) AS Series,
            CONCAT(DB_NAME(), '|', b.GuidPoliza, '|', b.IdMovimiento) AS Unique_Id,
            CAST(NULL AS char(10)) AS Post_Date,
            CASE
                WHEN b.TipoMovto <> 0 THEN 0
                WHEN b.IdMoneda = 1 THEN b.Importe
                WHEN NULLIF(b.ImporteME, 0) IS NOT NULL THEN b.ImporteME
                ELSE NULL
            END AS Originating_Debit_Amount,
            CASE
                WHEN b.TipoMovto <> 1 THEN 0
                WHEN b.IdMoneda = 1 THEN b.Importe
                WHEN NULLIF(b.ImporteME, 0) IS NOT NULL THEN b.ImporteME
                ELSE NULL
            END AS Originating_Credit_Amount,
            CASE
                WHEN b.IdMoneda = 1 THEN CAST(1.0 AS float)
                WHEN NULLIF(b.ImporteME, 0) IS NOT NULL THEN b.Importe / b.ImporteME
                ELSE NULL
            END AS Exchange_Rate,
            ISNULL(mon.Nombre, 'No Description') AS Currency_Description
        FROM BaseMovimientos b
        LEFT JOIN MovimientoDocumento md ON md.IdMovimiento = b.IdMovimiento
        LEFT JOIN dbo.Monedas mon ON mon.Id = b.IdMoneda
        LEFT JOIN GeneralesSQL.dbo.Usuarios u ON u.Id = b.IdUsuario
        " + GLTransactionsFilterSql + @"
        ORDER BY
            b.Fecha, b.Ejercicio, b.Periodo, b.TipoPol, b.Folio, b.NumMovto, b.IdMovimiento;
        ";

    private static async Task<byte[]> DownloadGLTransactionsAsync(
        ExcelPackage excel,
        SqlConnection connection,
        DownloadParameters p,
        IProgress<DownloadProgress> progress,
        ILogger logger,
        CancellationToken ct)
    {
        string? glBegin = null, glEndBound = null;
        if (p.UseGLRange)
        {
            glBegin = p.GLBegin!.Trim();
            glEndBound = p.GLEnd!.Trim() + "￿"; // sentinel so prefix ranges include all sub-accounts
        }

        var journalEntry = string.IsNullOrWhiteSpace(p.JournalEntry) ? null : p.JournalEntry.Trim();
        var originalMasterName = string.IsNullOrWhiteSpace(p.OriginalMasterName) ? null : p.OriginalMasterName.Trim();

        var beginDate = p.BeginDate.Date;
        var endDateExclusive = p.EndDate.Date.AddDays(1);

        int totalRecords = await CountGLTransactionsAsync(
            connection, beginDate, endDateExclusive, glBegin, glEndBound, journalEntry, originalMasterName, logger, ct);

        progress.Report(new DownloadProgress(0, totalRecords));

        if (totalRecords == 0)
            return Array.Empty<byte>();

        var ws = excel.Workbook.Worksheets.Add("GL_Trx");
        WriteGLTransactionsHeader(ws);

        int row = 2;
        int fetched = 0;

        using (var cmd = new SqlCommand(GLTransactionsCte + GLTransactionsSelectSql, connection) { CommandTimeout = 300 })
        {
            AddGLTransactionsParameters(cmd, beginDate, endDateExclusive, glBegin, glEndBound, journalEntry, originalMasterName);

            logger.LogDebug("SQL [GLTransactions]:\n{Sql}", BuildDebugSql(cmd));
            using var rdr = await cmd.ExecuteReaderAsync(ct);

            while (await rdr.ReadAsync(ct))
            {
                ct.ThrowIfCancellationRequested();

                WriteGLTransactionRow(ws, row, rdr);
                row++;

                fetched++;
                if (fetched % 200 == 0)
                    progress.Report(new DownloadProgress(fetched, totalRecords));
            }
        }

        progress.Report(new DownloadProgress(totalRecords, totalRecords));
        SetGLTransactionsColumnWidths(ws);

        return excel.GetAsByteArray();
    }

    private static async Task<int> CountGLTransactionsAsync(
        SqlConnection conn, DateTime beginDate, DateTime endDateExclusive,
        string? glBegin, string? glEndBound, string? journalEntry, string? originalMasterName, ILogger logger, CancellationToken ct)
    {
        using var cmd = new SqlCommand(GLTransactionsCte + GLTransactionsCountSql, conn) { CommandTimeout = 300 };
        AddGLTransactionsParameters(cmd, beginDate, endDateExclusive, glBegin, glEndBound, journalEntry, originalMasterName);
        logger.LogDebug("SQL [CountGLTransactions]:\n{Sql}", BuildDebugSql(cmd));
        return Convert.ToInt32(await cmd.ExecuteScalarAsync(ct));
    }

    private static void AddGLTransactionsParameters(
        SqlCommand cmd, DateTime beginDate, DateTime endDateExclusive,
        string? glBegin, string? glEndBound, string? journalEntry, string? originalMasterName)
    {
        cmd.Parameters.Add("@BeginDate", SqlDbType.DateTime).Value = beginDate;
        cmd.Parameters.Add("@EndDateExclusive", SqlDbType.DateTime).Value = endDateExclusive;
        cmd.Parameters.Add("@GLBegin", SqlDbType.NVarChar, 60).Value = (object?)glBegin ?? DBNull.Value;
        cmd.Parameters.Add("@GLEnd", SqlDbType.NVarChar, 60).Value = (object?)glEndBound ?? DBNull.Value;
        cmd.Parameters.Add("@JournalEntry", SqlDbType.NVarChar, 60).Value = (object?)journalEntry ?? DBNull.Value;
        cmd.Parameters.Add("@OriginalMasterName", SqlDbType.NVarChar, 254).Value = (object?)originalMasterName ?? DBNull.Value;
    }

    private static void WriteGLTransactionsHeader(ExcelWorksheet ws)
    {
        for (int i = 0; i < GLTransactionsColumns.Length; i++)
        {
            ws.Cells[1, i + 1].Value = GLTransactionsColumns[i];
            ws.Cells[1, i + 1].Style.Font.Bold = true;
        }
    }

    private static void WriteGLTransactionRow(ExcelWorksheet ws, int row, SqlDataReader rdr)
    {
        ws.Cells[row, 1].Value = rdr[0]?.ToString();                                   // Trx_Status
        ws.Cells[row, 2].Value = rdr[1]?.ToString();                                   // Trx_Date (dd-mm-yyyy)
        ws.Cells[row, 3].Value = rdr[2]?.ToString();                                   // Journal_Entry
        ws.Cells[row, 4].Value = FormatAccountCode(rdr[3]?.ToString() ?? "");         // Account_Number
        ws.Cells[row, 5].Value = rdr[4]?.ToString();                                   // Account_Description

        ws.Cells[row, 6].Value = rdr.IsDBNull(5) ? 0m : (decimal)rdr.GetDouble(5);   // Debit_Amount
        ws.Cells[row, 6].Style.Numberformat.Format = "#,##0.00";
        ws.Cells[row, 7].Value = rdr.IsDBNull(6) ? 0m : (decimal)rdr.GetDouble(6);   // Credit_Amount
        ws.Cells[row, 7].Style.Numberformat.Format = "#,##0.00";

        ws.Cells[row, 8].Value = rdr.IsDBNull(7) ? null : CleanText(rdr[7]?.ToString());    // Description
        ws.Cells[row, 9].Value = rdr.IsDBNull(8) ? null : CleanText(rdr[8]?.ToString());    // Reference
        ws.Cells[row, 10].Value = rdr.IsDBNull(9)  ? null : rdr[9]?.ToString();             // Source_Document
        ws.Cells[row, 11].Value = rdr.IsDBNull(10) ? null : rdr[10]?.ToString();            // Originating_Trx_Source
        ws.Cells[row, 12].Value = rdr.IsDBNull(11) ? null : rdr[11]?.ToString();            // Originating_Master_Id
        ws.Cells[row, 13].Value = rdr.IsDBNull(12) ? null : rdr[12]?.ToString();            // Originating_Master_Name
        ws.Cells[row, 14].Value = rdr.IsDBNull(13) ? null : rdr[13]?.ToString();            // Originating_Doc_Number (siempre vacío)
        ws.Cells[row, 15].Value = rdr.IsDBNull(14) ? null : rdr[14]?.ToString();            // Currency_Id
        ws.Cells[row, 16].Value = rdr.IsDBNull(15) ? null : rdr[15]?.ToString();            // Last_User
        ws.Cells[row, 17].Value = rdr.IsDBNull(16) ? null : rdr[16]?.ToString();            // User_Who_Posted (siempre vacío)
        ws.Cells[row, 18].Value = rdr.IsDBNull(17) ? null : rdr[17]?.ToString();            // Batch_Number (siempre vacío)
        ws.Cells[row, 19].Value = rdr.IsDBNull(18) ? null : rdr[18]?.ToString();            // Series (siempre vacío)
        ws.Cells[row, 20].Value = rdr[19]?.ToString();                                      // Unique_Id
        ws.Cells[row, 21].Value = rdr.IsDBNull(20) ? null : rdr[20]?.ToString();            // Post_Date (siempre vacío)

        if (!rdr.IsDBNull(21))
        {
            ws.Cells[row, 22].Value = (decimal)rdr.GetDouble(21);                    // Originating_Debit_Amount
            ws.Cells[row, 22].Style.Numberformat.Format = "#,##0.00000";
        }
        if (!rdr.IsDBNull(22))
        {
            ws.Cells[row, 23].Value = (decimal)rdr.GetDouble(22);                    // Originating_Credit_Amount
            ws.Cells[row, 23].Style.Numberformat.Format = "#,##0.00000";
        }
        if (!rdr.IsDBNull(23))
        {
            ws.Cells[row, 24].Value = rdr.GetDouble(23);                             // Exchange_Rate
            ws.Cells[row, 24].Style.Numberformat.Format = "#,##0.0000000";
        }

        ws.Cells[row, 25].Value = rdr.GetString(24);                                 // Currency_Description

        ws.Cells[row, 1, row, 25].Style.Font.Size = 9;
    }

    private static void SetGLTransactionsColumnWidths(ExcelWorksheet ws)
    {
        ws.Column(1).Width = 12;  // Trx_Status
        ws.Column(2).Width = 11;  // Trx_Date
        ws.Column(3).Width = 20;  // Journal_Entry
        ws.Column(4).Width = 16;  // Account_Number
        ws.Column(5).Width = 32;  // Account_Description
        ws.Column(6).Width = 13;  // Debit_Amount
        ws.Column(7).Width = 13;  // Credit_Amount
        ws.Column(8).Width = 30;  // Description
        ws.Column(9).Width = 20;  // Reference
        ws.Column(10).Width = 20; // Source_Document
        ws.Column(11).Width = 12; // Originating_Trx_Source
        ws.Column(12).Width = 16; // Originating_Master_Id
        ws.Column(13).Width = 32; // Originating_Master_Name
        ws.Column(14).Width = 18; // Originating_Doc_Number
        ws.Column(15).Width = 11; // Currency_Id
        ws.Column(16).Width = 14; // Last_User
        ws.Column(17).Width = 14; // User_Who_Posted
        ws.Column(18).Width = 14; // Batch_Number
        ws.Column(19).Width = 10; // Series
        ws.Column(20).Width = 20; // Unique_Id
        ws.Column(21).Width = 11; // Post_Date
        ws.Column(22).Width = 15; // Originating_Debit_Amount
        ws.Column(23).Width = 15; // Originating_Credit_Amount
        ws.Column(24).Width = 13; // Exchange_Rate
        ws.Column(25).Width = 20; // Currency_Description
    }
}
