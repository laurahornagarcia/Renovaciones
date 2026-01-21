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

app.MapPost("/api/process", async (
    HttpRequest request,
    PriceProfileStore profileStore,
    ExcelOfferService excelService) =>
{
    Console.WriteLine("DEBUG (Program.cs): Received request to /api/process.");

    if (!request.HasFormContentType)
        return Results.BadRequest("Request must be multipart/form-data.");

    var form = await request.ReadFormAsync();

    var file = form.Files.GetFile("file");
    var profileId = form["profileId"].ToString();
    var sequenceStr = form["sequence"].ToString();

    Console.WriteLine($"DEBUG (Program.cs): profileId received: '{(string.IsNullOrWhiteSpace(profileId) ? "(null/empty)" : profileId)}', sequence received: '{sequenceStr}'");

    var effectiveSequence = 1;
    if (!string.IsNullOrWhiteSpace(sequenceStr) && int.TryParse(sequenceStr, out var seq) && seq > 0)
        effectiveSequence = seq;

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

        Console.WriteLine($"DEBUG (Program.cs): PriceProfile loaded. Name: '{priceProfile.Name}'");
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
