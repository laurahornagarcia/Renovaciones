using System;
using System.IO;
using System.Linq;
using System.Globalization;
using System.Text.RegularExpressions;
using ClosedXML.Excel;
using Ofertum.Renovaciones.Models;

namespace Ofertum.Renovaciones.Services
{
    public class ExcelOfferService
    {
        public void ProcessExcelOffer(byte[] xlsxBytes, PriceProfile? profile, DateTime todayLocal, string fileName, int dailySequence, out byte[] resultBytes, out string newOfferNumber, out string outputFileName)
        {
            using (var stream = new MemoryStream(xlsxBytes))
            {
                using (var workbook = new XLWorkbook(stream))
                {
                    // Initialize output parameters
                    newOfferNumber = "";
                    outputFileName = "";
                    resultBytes = Array.Empty<byte>();

                    // Rule 6: Generate newOfferNumber
                    string prefix = "OF"; // Default prefix

                    // Try to get prefix from "Nº de Oferta" cell
                    foreach (var worksheet in workbook.Worksheets)
                    {
                        var offerNumberCell = worksheet.CellsUsed().FirstOrDefault(c => c.Value.ToString().Contains("Nº de Oferta", StringComparison.OrdinalIgnoreCase));
                        if (offerNumberCell != null)
                        {
                            var rightCell = offerNumberCell.CellRight();
                            if (!string.IsNullOrWhiteSpace(rightCell.Value.ToString()) && rightCell.Value.ToString().Length >= 2)
                            {
                                prefix = rightCell.Value.ToString().Substring(0, 2).ToUpperInvariant();
                                break;
                            }
                        }
                    }

                    // If prefix not found in cell, try from filename
                    if (prefix == "OF" && fileName.Length >= 2)
                    {
                        prefix = fileName.Substring(0, 2).ToUpperInvariant();
                    }

                    newOfferNumber = $"{prefix}-{todayLocal:yyyyMMdd}-{dailySequence:D2}";

                    // Rule 7: outputFileName
                    string baseFileName = Path.GetFileNameWithoutExtension(fileName);
                    string extension = Path.GetExtension(fileName);
                    string offerNumberRegexPattern = "[A-Z]{2}-\\d{8}-\\d{2}"; // Pattern for XX-YYYYMMDD-NN

                    Match match2 = Regex.Match(baseFileName, offerNumberRegexPattern);
                    if (match2.Success)
                    {
                        outputFileName = Regex.Replace(baseFileName, offerNumberRegexPattern, newOfferNumber) + extension;
                    }
                    else
                    {
                        // If no existing offer number pattern found, append the new one
                        outputFileName = $"{baseFileName}_{newOfferNumber}{extension}";
                    }

                    var cultureInfo = new CultureInfo("es-ES");

                    foreach (var worksheet in workbook.Worksheets)
                    {
                        Console.WriteLine($"DEBUG: Processing worksheet: {worksheet.Name}");
                        if (profile?.Prices != null && profile.Prices.Any())
                        {
                            Console.WriteLine($"DEBUG: Profile prices being searched: {string.Join(", ", profile.Prices.Select(p => $"'{p.Key}':'{p.Value}'"))}");
                        }

                        // Rule 1, 2, 3
                        foreach (var cell in worksheet.CellsUsed())
                        {
                            string cellValue = cell.Value.ToString();

                            if (cellValue.Contains("Nº de Oferta", StringComparison.OrdinalIgnoreCase))
                            {
                                cell.CellRight().SetValue(newOfferNumber);
                            }
                            else if (cellValue.Contains("F. Ppto.", StringComparison.OrdinalIgnoreCase))
                            {
                                cell.CellRight().SetValue(todayLocal.ToString("dd/MM/yyyy", cultureInfo));
                            }
                            else if (cellValue.Contains("Validez", StringComparison.OrdinalIgnoreCase))
                            {
                                cell.CellRight().SetValue(todayLocal.AddMonths(2).ToString("dd/MM/yyyy", cultureInfo));
                            }
                            // Rule 4: Vigencia
                            else if (cellValue.Contains("Vigencia:", StringComparison.OrdinalIgnoreCase))
                            {
                                var match = Regex.Match(cellValue, @"Vigencia: del (\d{2}/\d{2}/\d{4}) al (\d{2}/\d{2}/\d{4})");
                                if (match.Success)
                                {
                                    DateTime startDate = DateTime.ParseExact(match.Groups[1].Value, "dd/MM/yyyy", cultureInfo);
                                    DateTime endDate = DateTime.ParseExact(match.Groups[2].Value, "dd/MM/yyyy", cultureInfo);

                                    string newStartDate = startDate.AddYears(1).ToString("dd/MM/yyyy", cultureInfo);
                                    string newEndDate = endDate.AddYears(1).ToString("dd/MM/yyyy", cultureInfo);

                                    cell.SetValue($"Vigencia: del {newStartDate} al {newEndDate}");
                                }
                            }
                        }

                        // Rule 5: LICENCIAS
                        var licensesCell = worksheet.CellsUsed().FirstOrDefault(c => c.Value.ToString().Contains("LICENCIAS", StringComparison.OrdinalIgnoreCase));
                        if (licensesCell != null)
                        {
                            int startRow = Math.Max(licensesCell.WorksheetRow().RowNumber() + 1, 20); // Start from row 20 or after LICENCIAS
                            int endRow = 40; // End at row 40

                            for (int currentRow = startRow; currentRow <= endRow; currentRow++)
                            { 
                                var cellF = worksheet.Cell(currentRow, "F");
                Console.WriteLine($"DEBUG: startRow '{startRow}', endRow  '{endRow}', cellF.Value.ToString()  '{cellF.Value.ToString()}' ");
                                if (profile?.Prices != null && !string.IsNullOrWhiteSpace(cellF.Value.ToString()))
                                {
                                    string cellFValue = cellF.Value.ToString();
                                    string normalizedKey = NormalizePriceKey(cellFValue);
                                    Console.WriteLine($"DEBUG: Original cell F value: '{cellFValue}', Normalized key: '{normalizedKey}'");

                                    if (profile.Prices.TryGetValue(normalizedKey, out string? newPrice))
                                    {
                                        Console.WriteLine($"DEBUG: Match found for '{normalizedKey}'. New price: '{newPrice}'");
                                        cellF.SetValue(newPrice);
                                    }
                                    else
                                    {
                                        Console.WriteLine($"DEBUG: No match found for normalized key '{normalizedKey}' in profile prices.");
                                    }
                                }
                            }
                        }
                    }

                    using (var outputStream = new MemoryStream())
                    {
                        workbook.SaveAs(outputStream);
                        resultBytes = outputStream.ToArray();
                    }
                }
            }
        }

        private string NormalizePriceKey(string input)
        {
            // Convertir a Double para manejar formatos decimales y luego a string sin decimales si es un entero
            if (double.TryParse(input, NumberStyles.Any, CultureInfo.InvariantCulture, out double value))
            {
                // Si es un número entero (ej. 348.0), lo normalizamos a "348"
                if (value == Math.Floor(value))
                {
                    return value.ToString("0", CultureInfo.InvariantCulture);
                }
                // Si tiene decimales, lo mantenemos con un punto como separador
                return value.ToString(CultureInfo.InvariantCulture);
            }
            // Si no es un número, eliminamos símbolos de moneda, espacios y normalizamos a punto
            string normalized = input.Replace("€", "").Replace(" ", "").Replace(",", ".");
            // Intentar convertir a número nuevamente para asegurar uniformidad (ej. "348.00" -> "348")
            if (double.TryParse(normalized, NumberStyles.Any, CultureInfo.InvariantCulture, out value))
            {
                if (value == Math.Floor(value))
                {
                    return value.ToString("0", CultureInfo.InvariantCulture);
                }
                return value.ToString(CultureInfo.InvariantCulture);
            }
            return normalized.Trim();
        }
    }
}
