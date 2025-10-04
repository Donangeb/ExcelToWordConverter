using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;
using OfficeOpenXml;
using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.IO;
using System.Linq;
using System.Threading.Tasks;

namespace ExcelToWordConverter.Models
{
    public static class ExamConverter
    {
        public static async Task ConvertAsync(string excelPath, string outputPath)
        {
            await Task.Run(() =>
            {
                ExcelPackage.LicenseContext = OfficeOpenXml.LicenseContext.NonCommercial;

                var year = GetStartYear(excelPath);
                if (!year.HasValue)
                    throw new Exception("Не удалось определить год начала подготовки!");

                var semesters = GetSemestersByYear(year.Value);
                if (semesters.Count == 0)
                    throw new Exception("Для этого года не определены семестры!");

                using var package = new ExcelPackage(new FileInfo(excelPath));
                var worksheet = package.Workbook.Worksheets.FirstOrDefault(ws => ws.Name.Contains("План", StringComparison.OrdinalIgnoreCase));
                if (worksheet == null)
                    throw new Exception("Не найден лист 'План'");

                using var doc = WordprocessingDocument.Create(outputPath, WordprocessingDocumentType.Document);
                MainDocumentPart mainPart = doc.AddMainDocumentPart();
                mainPart.Document = new Document(new Body());
                Body body = mainPart.Document.Body;

                // Заголовок
                var titleRun = new Run(new Text("ГРАФИК СДАЧИ ЗАЧЕТОВ И ЭКЗАМЕНОВ"))
                {
                    RunProperties = new RunProperties(new Bold(), new DocumentFormat.OpenXml.Wordprocessing.FontSize() { Val = "32" })
                };
                body.Append(new Paragraph(titleRun));
                body.Append(new Paragraph(new Run(new Text($"Год начала подготовки: {year}"))));

                foreach (var sem in semesters)
                {
                    var zachety = new List<string>();
                    var ekzamen = new List<string>();

                    int nameCol = GetColumnIndex(worksheet, "Наименование");
                    int zachetCol = GetColumnIndex(worksheet, "Зачет");
                    int zachetOcenkaCol = GetColumnIndex(worksheet, "Зачет с оценкой");
                    int kpCol = GetColumnIndex(worksheet, "КП");
                    int ekzamenCol = GetColumnIndex(worksheet, "Экзамен");

                    for (int row = 4; row <= worksheet.Dimension.End.Row; row++)
                    {
                        var name = worksheet.Cells[row, nameCol].Value?.ToString()?.Trim();
                        if (string.IsNullOrEmpty(name) ||
                            name.ToLower().StartsWith("блок") ||
                            name.ToLower().StartsWith("обязательная часть") ||
                            name.ToLower().Contains("дисциплины по выбору"))
                            continue;

                        if (zachetCol > 0 && CellContainsSemester(worksheet.Cells[row, zachetCol].Value, sem))
                            zachety.Add(name);

                        if (zachetOcenkaCol > 0 && CellContainsSemester(worksheet.Cells[row, zachetOcenkaCol].Value, sem))
                            zachety.Add($"{name} (зачет с оценкой)");

                        if (kpCol > 0 && CellContainsSemester(worksheet.Cells[row, kpCol].Value, sem))
                            zachety.Add($"{name} (КП)");

                        if (ekzamenCol > 0 && CellContainsSemester(worksheet.Cells[row, ekzamenCol].Value, sem))
                            ekzamen.Add(name);
                    }

                    // Заголовок семестра
                    var semRun = new Run(new Text($"{sem} семестр"))
                    {
                        RunProperties = new RunProperties(new Bold(), new DocumentFormat.OpenXml.Wordprocessing.FontSize() { Val = "24" })
                    };
                    body.Append(new Paragraph(semRun));

                    // Таблица
                    Table table = new Table(new TableProperties(
                        new TableBorders(
                            new TopBorder { Val = BorderValues.Single, Size = 4 },
                            new BottomBorder { Val = BorderValues.Single, Size = 4 },
                            new LeftBorder { Val = BorderValues.Single, Size = 4 },
                            new RightBorder { Val = BorderValues.Single, Size = 4 },
                            new InsideHorizontalBorder { Val = BorderValues.Single, Size = 4 },
                            new InsideVerticalBorder { Val = BorderValues.Single, Size = 4 }
                        )
                    ));

                    TableRow headerRow = new TableRow();
                    headerRow.Append(CreateTableCell("Зачеты", true));
                    headerRow.Append(CreateTableCell("Экзамены", true));
                    table.Append(headerRow);

                    int maxRows = Math.Max(zachety.Count, ekzamen.Count);
                    for (int i = 0; i < maxRows; i++)
                    {
                        TableRow dataRow = new TableRow();
                        dataRow.Append(CreateTableCell(i < zachety.Count ? zachety[i] : ""));
                        dataRow.Append(CreateTableCell(i < ekzamen.Count ? ekzamen[i] : ""));
                        table.Append(dataRow);
                    }

                    body.Append(table);
                    body.Append(new Paragraph());
                }

                mainPart.Document.Save();
            });
        }

        private static int? GetStartYear(string excelPath)
        {
            using var package = new ExcelPackage(new FileInfo(excelPath));
            var worksheet = package.Workbook.Worksheets.FirstOrDefault(ws => ws.Name.Contains("Титул", StringComparison.OrdinalIgnoreCase));
            if (worksheet == null) return null;

            for (int row = 1; row <= worksheet.Dimension.End.Row; row++)
            {
                for (int col = 1; col <= worksheet.Dimension.End.Column; col++)
                {
                    var cellValue = worksheet.Cells[row, col].Value?.ToString();
                    if (!string.IsNullOrEmpty(cellValue) && cellValue.Contains("Год начала подготовки"))
                    {
                        for (int offset = 1; offset < 20; offset++)
                        {
                            var yearCellValue = worksheet.Cells[row, col + offset].Value?.ToString();
                            if (!string.IsNullOrEmpty(yearCellValue))
                            {
                                if (int.TryParse(yearCellValue.Trim(), out int year))
                                    return year;
                                if (double.TryParse(yearCellValue.Trim().Replace(",", "."),
                                    System.Globalization.NumberStyles.Any,
                                    System.Globalization.CultureInfo.InvariantCulture, out double y))
                                    return (int)y;
                            }
                        }
                    }
                }
            }
            return null;
        }

        private static List<int> GetSemestersByYear(int startYear)
        {
            int diff = DateTime.Now.Year - startYear;
            return diff switch
            {
                0 => new() { 1, 2 },
                1 => new() { 3, 4 },
                2 => new() { 5, 6 },
                3 => new() { 7, 8 },
                _ => new()
            };
        }

        private static bool CellContainsSemester(object cellValue, int sem)
        {
            if (cellValue == null) return false;
            string text = cellValue.ToString().Trim();
            if (text.Contains(sem.ToString()) || text.StartsWith(sem.ToString() + "."))
                return true;
            if (double.TryParse(text.Replace(",", "."), System.Globalization.NumberStyles.Any,
                System.Globalization.CultureInfo.InvariantCulture, out double d))
                return (int)d == sem;
            return false;
        }

        private static TableCell CreateTableCell(string text, bool isHeader = false)
        {
            Run run = new Run(new Text(text))
            {
                RunProperties = new RunProperties(new DocumentFormat.OpenXml.Wordprocessing.FontSize() { Val = "20" }, new RunFonts() { Ascii = "Times New Roman" })
            };
            if (isHeader) run.RunProperties.Append(new Bold());

            Paragraph p = new Paragraph(run)
            {
                ParagraphProperties = new ParagraphProperties(new SpacingBetweenLines() { After = "0" })
            };

            TableCell cell = new TableCell(p)
            {
                TableCellProperties = new TableCellProperties(new TableCellVerticalAlignment { Val = TableVerticalAlignmentValues.Center })
            };
            return cell;
        }

        private static int GetColumnIndex(ExcelWorksheet worksheet, string columnName)
        {
            for (int row = 1; row <= Math.Min(10, worksheet.Dimension.End.Row); row++)
            {
                for (int col = 1; col <= worksheet.Dimension.End.Column; col++)
                {
                    var header = worksheet.Cells[row, col].Value?.ToString()?.Trim();
                    if (string.IsNullOrEmpty(header)) continue;

                    string normalizedHeader = header.ToLower().Replace(" ", "").Replace(".", "");
                    string normalizedName = columnName.ToLower().Replace(" ", "").Replace(".", "");

                    if (normalizedHeader.Contains(normalizedName))
                        return col;

                    if (columnName.ToLower().Contains("с оценкой") && normalizedHeader.Contains("зачетс"))
                        return col;
                }
            }
            return -1;
        }
    }
}
