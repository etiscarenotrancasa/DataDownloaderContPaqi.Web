namespace DataDownloaderContPaqi.Web.Models;

public enum ReportType
{
    AuxiliaresCatalogo,
    GLTransactions,
}

public class DownloadParameters
{
    public string Company { get; set; } = "Centro Regional de Capacitacion";
    public ReportType ReportType { get; set; } = ReportType.GLTransactions;
    public DateTime BeginDate { get; set; } = DateTime.Today;
    public DateTime EndDate { get; set; } = DateTime.Today;
    public bool UseGLRange { get; set; }
    public string? GLBegin { get; set; }
    public string? GLEnd { get; set; }
    public string? JournalEntry { get; set; }
    public string? OriginalMasterName { get; set; }

    public static readonly Dictionary<string, string> CompanyDatabases = new()
    {
        ["Servicios Tractomotrices"]        = "ctSERVICIOS_TRACTOMOTRICES_SA_DE_C2",
        ["Centro Regional de Capacitacion"] = "ctCENTRO_REGIONAL_DE_CAPACITACION_2",
        ["SecureFleet Innovations"]         = "ctSECUREFLEET_INNOVATIONS",
        ["Trancasa Logistica de Mexico"]    = "ctTRANCASA_LOGISTICA_DE_MEXICO",
    };

    public string? DatabaseName =>
        CompanyDatabases.TryGetValue(Company, out var db) ? db : null;
}
