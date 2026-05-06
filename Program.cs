using Azure.Identity;
using Azure.Storage.Blobs;
using WordTemplateBlobWebApi.Options;
using WordTemplateBlobWebApi.Services;

var builder = WebApplication.CreateBuilder(args);

builder.Services.AddControllers();
builder.Services.AddEndpointsApiExplorer();
builder.Services.AddSwaggerGen();

builder.Services.Configure<AzureStorageOptions>(
    builder.Configuration.GetSection(AzureStorageOptions.SectionName));

builder.Services.AddSingleton(sp =>
{
    var configuration = sp.GetRequiredService<IConfiguration>();
    var accountUrl = configuration["AzureStorage:AccountUrl"];

    if (string.IsNullOrWhiteSpace(accountUrl) || accountUrl.Contains("YOUR_STORAGE_ACCOUNT", StringComparison.OrdinalIgnoreCase))
    {
        throw new InvalidOperationException(
            "AzureStorage:AccountUrl must be configured in appsettings.json, user secrets, environment variables, or Azure App Configuration.");
    }

    return new BlobServiceClient(new Uri(accountUrl), new DefaultAzureCredential());
});

builder.Services.AddScoped<IDocumentGeneratorService, DocumentGeneratorService>();

var app = builder.Build();

if (app.Environment.IsDevelopment())
{
    app.UseSwagger();
    app.UseSwaggerUI();
}
else
{
    app.UseHttpsRedirection();
}

app.MapControllers();

app.Run();
