using DataDownloaderContPaqi.Web.Components;
using DataDownloaderContPaqi.Web.Services;

DotNetEnv.Env.Load();

var builder = WebApplication.CreateBuilder(args);

// Build the base SQL Server connection string from environment variables.
// InitialCatalog is set per-request inside DataDownloaderService.
var dbServer   = builder.Configuration["TCAMN_SERVER"]   ?? throw new InvalidOperationException("TCAMN_SERVER not configured.");
var dbUser     = builder.Configuration["TCAMN_DB_USER"]  ?? throw new InvalidOperationException("TCAMN_DB_USER not configured.");
var dbPassword = builder.Configuration["TCAMN_DB_PASS"]  ?? throw new InvalidOperationException("TCAMN_DB_PASS not configured. Run: dotnet user-secrets set \"TCAMN_DB_PASS\" \"<password>\"");
var trustCert  = builder.Configuration["TCAMN_TRUSTED_SERVER_CERTIFICATE"] ?? "Yes";

builder.Configuration["ConnectionStrings:DefaultConnection"] =
    $"Server=tcp:{dbServer},49801;User Id={dbUser};Password={dbPassword};TrustServerCertificate={trustCert};Encrypt=False;";

builder.Services.AddRazorComponents()
    .AddInteractiveServerComponents();

builder.Services.AddScoped<DataDownloaderService>();
builder.Services.AddSingleton<JobStore>();

var app = builder.Build();

if (!app.Environment.IsDevelopment())
{
    app.UseExceptionHandler("/Error", createScopeForErrors: true);
}
app.UseStatusCodePagesWithReExecute("/not-found", createScopeForStatusCodePages: true);
app.UseAntiforgery();

app.MapStaticAssets();
app.MapRazorComponents<App>()
    .AddInteractiveServerRenderMode();

// Minimal API: serve completed job file and remove it from the store
app.MapGet("/api/download/{jobId:guid}", (Guid jobId, JobStore store) =>
{
    if (!store.TryGet(jobId, out var bytes, out var fileName))
        return Results.NotFound("Download link expired or not found.");

    store.Remove(jobId);
    return Results.File(
        bytes,
        "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        fileName);
});

app.Run();
