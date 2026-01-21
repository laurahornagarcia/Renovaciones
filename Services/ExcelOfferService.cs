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
        public void ProcessExcelOffer(
            byte[] xlsxBytes,
            PriceProfile? profile,
            DateTime todayLocal,
            string fileName,
            int dailySequence,
            out byte[] resultBytes,
            out string newOfferNumber,
            out string outputFileName)
        {
            using var stream = new MemoryStream(xlsxBytes);
            using var workbook = new XLWorkbook(stream);

            // ---------- 1) Generar newOfferNumber ----------
            string prefix = "OF";

            foreach (var ws in workbook.Worksheets)
            {
                var offerNumberCell = ws.CellsUsed()
                    .FirstOrDefault(c => c.GetString().Contains("Nº de Oferta", StringComparison.OrdinalIgnoreCase));

                if (offerNumberCell != null)
                {
                    var rightCellText = offerNumberCell.CellRight().GetString();
                    var letters = new string((rightCellText ?? "").Where(char.IsLetter).ToArray()).ToUpperInvariant();
                    if (letters.Length >= 2)
                    {
                        prefix = letters.Substring(0, 2);
                        break;
                    }
                }
            }

            if (prefix == "OF")
            {
                var baseName = Path.GetFileNameWithoutExtension(fileName) ?? "";
                var letters = new string(baseName.Where(char.IsLetter).ToArray()).ToUpperInvariant();
                if (letters.Length >= 2)
                    prefix = letters.Substring(0, 2);
            }

            newOfferNumber = $"{prefix}-{todayLocal:yyyyMMdd}-{dailySequence:D2}";

            // ---------- 2) Nombre del fichero de salida ----------
          // ---------- 2) Nombre del fichero de salida ----------
var baseFileName = Path.GetFileNameWithoutExtension(fileName) ?? "";
var extension = Path.GetExtension(fileName);

// Permite:
// "LH-20250109-1 - ..."
// "LH - 20250109 - 01 - ..."
// "·LH-200250109-1 - ..."
// Fecha 8 o 9 dígitos, secuencia 1 o 2 dígitos
var leadingOfferPattern =
    @"^\s*[\p{P}\p{S}]*(?<prefix>LH|RC)\s*-\s*(?<date>\d{8,9})\s*-\s*(?<seq>\d{1,2})(?<rest>.*)$";

var m = Regex.Match(baseFileName, leadingOfferPattern, RegexOptions.IgnoreCase);

if (m.Success)
{
    // Sustituye SOLO el número inicial, conserva el resto exacto
    outputFileName = $"{newOfferNumber}{m.Groups["rest"].Value}{extension}";
}
else
{
    // Si no detecta número al inicio, lo pone delante (sin duplicar)
    outputFileName = $"{newOfferNumber} - {baseFileName}{extension}";
}

            var cultureInfo = CultureInfo.GetCultureInfo("es-ES");

            // ---------- 3) Procesar todas las hojas ----------
            foreach (var worksheet in workbook.Worksheets)
            {
                // 3.1 Reglas de celdas a la derecha (sin romper formato)
                ApplyRightCellReplacements(worksheet, todayLocal, newOfferNumber);

                // 3.2 Vigencia: D20:D40 (solo cambia fechas, no el texto)
                UpdateVigenciaInColumnD(worksheet, cultureInfo);

                // 3.3 Precios bajo LICENCIAS (columna F, filas 20..40)
                ReplacePricesUnderLicencias(worksheet, profile);
            }

            using var outputStream = new MemoryStream();
            workbook.SaveAs(outputStream);
            resultBytes = outputStream.ToArray();
        }

        // ------------------ Reglas Nº Oferta / F. Ppto. / Validez ------------------

        private static void ApplyRightCellReplacements(IXLWorksheet worksheet, DateTime todayLocal, string newOfferNumber)
        {
            foreach (var cell in worksheet.CellsUsed())
            {
                var cellText = cell.GetString();

                if (cellText.Contains("Nº de Oferta", StringComparison.OrdinalIgnoreCase))
                {
                    SetTextKeepStyle(cell.CellRight(), newOfferNumber);
                }
                else if (cellText.Contains("F. Ppto.", StringComparison.OrdinalIgnoreCase))
                {
                    // Fecha real (mantiene formato de celda y evita texto)
                    SetDateKeepStyle(cell.CellRight(), todayLocal);
                }
                else if (cellText.Contains("Validez", StringComparison.OrdinalIgnoreCase))
                {
                    SetDateKeepStyle(cell.CellRight(), todayLocal.AddMonths(2));
                }
            }
        }

        // ------------------ Vigencia D20:D40 (solo fechas) ------------------

        private static void UpdateVigenciaInColumnD(IXLWorksheet worksheet, CultureInfo cultureInfo)
        {
            // Detecta pares: dd-MM-yyyy al dd-MM-yyyy  o  dd/MM/yyyy al dd/MM/yyyy
            var datePairRegex = new Regex(
                @"(\d{2})([-/])(\d{2})\2(\d{4})\s*al\s*(\d{2})([-/])(\d{2})\6(\d{4})",
                RegexOptions.IgnoreCase | RegexOptions.Compiled);

            var formats = new[] { "dd-MM-yyyy", "dd/MM/yyyy" };

            for (int row = 20; row <= 40; row++)
            {
                var cell = worksheet.Cell(row, 4); // D
                var text = cell.GetString();

                if (string.IsNullOrWhiteSpace(text))
                    continue;

                if (!text.Contains("Vigencia", StringComparison.OrdinalIgnoreCase))
                    continue;

                // Reemplaza TODAS las coincidencias dentro de la celda (por si hay varias líneas)
                var replaced = datePairRegex.Replace(text, m =>
                {
                    var sep1 = m.Groups[2].Value;
                    var sep2 = m.Groups[6].Value;

                    var d1 = $"{m.Groups[1].Value}{sep1}{m.Groups[3].Value}{sep1}{m.Groups[4].Value}";
                    var d2 = $"{m.Groups[5].Value}{sep2}{m.Groups[7].Value}{sep2}{m.Groups[8].Value}";

                    if (!DateTime.TryParseExact(d1, formats, cultureInfo, DateTimeStyles.None, out var start))
                        return m.Value;

                    if (!DateTime.TryParseExact(d2, formats, cultureInfo, DateTimeStyles.None, out var end))
                        return m.Value;

                    var newStart = start.AddYears(1).ToString($"dd{sep1}MM{sep1}yyyy", cultureInfo);
                    var newEnd = end.AddYears(1).ToString($"dd{sep2}MM{sep2}yyyy", cultureInfo);

                    // Devuelve SOLO el tramo "fecha al fecha"
                    return $"{newStart} al {newEnd}";
                });

                if (!string.Equals(text, replaced, StringComparison.Ordinal))
                {
                    SetTextKeepStyle(cell, replaced);
                }
            }
        }

        // ------------------ Precios bajo LICENCIAS ------------------

        private static void ReplacePricesUnderLicencias(IXLWorksheet worksheet, PriceProfile? profile)
        {
            if (profile?.Prices == null || profile.Prices.Count == 0)
                return;

            var licensesCell = worksheet.CellsUsed()
                .FirstOrDefault(c => c.GetString().Contains("LICENCIAS", StringComparison.OrdinalIgnoreCase));

            if (licensesCell == null)
                return;

            // Mapa normalizado para que coincida aunque el JSON tenga "401" y la celda "401,00 €"
            var normalizedMap = profile.Prices
                .GroupBy(kv => NormalizePriceKey(kv.Key))
                .ToDictionary(g => g.Key, g => g.Last().Value);

            int startRow = Math.Max(licensesCell.WorksheetRow().RowNumber() + 1, 20);
            int endRow = 40;

            for (int row = startRow; row <= endRow; row++)
            {
                var cellF = worksheet.Cell(row, "F");

                // Lo que ve el usuario (mejor para matching)
                var raw = cellF.GetFormattedString();
                if (string.IsNullOrWhiteSpace(raw))
                    continue;

                var key = NormalizePriceKey(raw);

                if (normalizedMap.TryGetValue(key, out var newPrice) && !string.IsNullOrWhiteSpace(newPrice))
                {
                    SetPriceKeepStyle(cellF, newPrice);
                }
            }
        }

        // ------------------ Helpers: mantener formato ------------------

        private static void SetTextKeepStyle(IXLCell cell, string text)
        {
            var style = cell.Style;
            cell.Value = text;
            cell.Style = style;
        }

        private static void SetDateKeepStyle(IXLCell cell, DateTime date)
        {
            var style = cell.Style;
            var format = cell.Style.DateFormat.Format;

            cell.Value = date.Date;

            // restaura el estilo general
            cell.Style = style;

            // asegura formato fecha
            if (string.IsNullOrWhiteSpace(format))
                cell.Style.DateFormat.Format = "dd/MM/yyyy";
            else
                cell.Style.DateFormat.Format = format;
        }

        private static void SetPriceKeepStyle(IXLCell cell, string newPrice)
        {
            var style = cell.Style;
            var numberFormat = cell.Style.NumberFormat.Format;

            var cleaned = (newPrice ?? "").Replace("€", "").Trim();

            if (decimal.TryParse(cleaned, NumberStyles.Any, CultureInfo.GetCultureInfo("es-ES"), out var dec) ||
                decimal.TryParse(cleaned, NumberStyles.Any, CultureInfo.InvariantCulture, out dec))
            {
                cell.Value = dec;
                cell.Style = style;

                if (!string.IsNullOrWhiteSpace(numberFormat))
                    cell.Style.NumberFormat.Format = numberFormat;
            }
            else
            {
                cell.Value = newPrice;
                cell.Style = style;
            }
        }

        private static string NormalizePriceKey(string input)
        {
            var s = (input ?? "").Trim();
            s = s.Replace("€", "").Replace(" ", "");

            if (decimal.TryParse(s, NumberStyles.Any, CultureInfo.GetCultureInfo("es-ES"), out var d) ||
                decimal.TryParse(s, NumberStyles.Any, CultureInfo.InvariantCulture, out d))
            {
                return d.ToString("0.##", CultureInfo.InvariantCulture);
            }

            return s.Replace(",", ".").Trim();
        }
    }
}
