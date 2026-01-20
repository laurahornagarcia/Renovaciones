using Microsoft.AspNetCore.Components;
using Microsoft.AspNetCore.Components.Web;
using Ofertum.Renovaciones.Data;
using Ofertum.Renovaciones.Models;
using Ofertum.Renovaciones.Services;
using Microsoft.AspNetCore.Http;
using Microsoft.AspNetCore.Mvc;

var builder = WebApplication.CreateBuilder(args);

// Add services to the container.
builder.Services.AddRazorPages();
builder.Services.AddServerSideBlazor();
builder.Services.AddSingleton<WeatherForecastService>();
builder.Services.AddSingleton<PriceProfileStore>();
builder.Services.AddScoped<ExcelOfferService>();

var app = builder.Build();

// Configure the HTTP request pipeline.
if (!app.Environment.IsDevelopment())
{
    app.UseExceptionHandler("/Error");
    // The default HSTS value is 30 days. You may want to change this for production scenarios, see https://aka.ms/aspnetcore-hsts.
    app.UseHsts();
}

app.UseHttpsRedirection();

app.UseStaticFiles();

app.UseRouting();

app.MapBlazorHub();

// 1) Endpoint ANTES del fallback
app.MapPost("/api/process", async (
    [FromForm] IFormFile file,
    PriceProfileStore profileStore,
    ExcelOfferService excelService,
    string? profileId,
    int? sequence) =>
{
    Console.WriteLine($"DEBUG (Program.cs): Received request to /api/process.");
    Console.WriteLine($"DEBUG (Program.cs): profileId received: '{profileId ?? "(null)"}', sequence received: '{sequence ?? 0}'");

    int effectiveSequence = sequence ?? 1;
    if (file is null || file.Length == 0)
        return Results.BadRequest("No file uploaded.");

    PriceProfile? priceProfile = null;
    if (!string.IsNullOrWhiteSpace(profileId))
    {
        Console.WriteLine($"DEBUG (Program.cs): Attempting to load profile with ID: '{profileId}'");
        priceProfile = await profileStore.GetAsync(profileId);
        if (priceProfile is null)
        {
            Console.WriteLine($"DEBUG (Program.cs): PriceProfile with ID '{profileId}' NOT found.");
            return Results.NotFound($"PriceProfile with ID {profileId} not found.");
        }
        else
        {
            Console.WriteLine($"DEBUG (Program.cs): PriceProfile with ID '{profileId}' loaded successfully. Name: '{priceProfile.Name}'");
        }
    }

    byte[] fileBytes;
    using (var memoryStream = new MemoryStream())
    {
        await file.CopyToAsync(memoryStream);
        fileBytes = memoryStream.ToArray();
    }

    excelService.ProcessExcelOffer(
        fileBytes,
        priceProfile,
        DateTime.Now,
        file.FileName,
        effectiveSequence,
        out var resultBytes,
        out var newOfferNumber,
        out var outputFileName);

    return Results.File(
        resultBytes,
        "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        outputFileName);
});

// 2) Fallback SIEMPRE lo último
app.MapFallbackToPage("/_Host");

app.Run();
