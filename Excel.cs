using System;
using System.Collections.Generic;
using System.Globalization;
using System.Linq;
using System.Text;
using OfficeOpenXml;
using OfficeOpenXml.Style;

namespace HistData2Excel
{
    class Excel
    {
        const string dateFormat = "MM/dd/yy";
        static readonly CultureInfo fromCulture = new CultureInfo("en-US");
        public static string ExcelFilePath { get; set; } = System.IO.Path.Combine(Dde.TargetDir, $"HistData_{Dde.StartDate:yyyy-MM-dd_HH-mm}_{Dde.Duration}_{DateTime.Now.Ticks}.xlsx");


        public static void WriteNew(Dictionary<string, string> tags, List<string> Data)
        {
            //int sheetCount = 0;

            Console.WriteLine("Erstelle die Datei " + ExcelFilePath);

            using (ExcelPackage pPackage = new ExcelPackage())
            {
                ExcelWorksheet pWorkSheet = pPackage.Workbook.Worksheets.Add("Blatt_1");

                int startRow = 0;

                foreach (string data in Data)
                {
                    string[] lines = data.Split(new[] { '\r', '\n' }, StringSplitOptions.RemoveEmptyEntries);

                    for (int row = 0; row < lines.Length - 1; row++) // .Lenghth - 1 => letzte Zeile Weglassen; Letzte Zeile = erste Zeile nächste Abfrage
                    {
                        // Console.Write("\r\nRow=" + row);
                        string[] items = lines[row].Split(';');

                        if (row == 0) //Erste Zeile in data
                        {
                            if (startRow == 0) //Erste Zeile im Blatt
                                for (int col = 0; col < items.Length; col++)
                                {
                                    string tagName = items[col];

                                    using (ExcelRange pRange = pWorkSheet.Cells[startRow + row + 1, col + 1]) // Zeilen Id Start, Spalten Id Start[, Zeilen Id Ende, Spalten Id Ende]
                                    {
                                        pRange.Style.Font.Bold = true;

                                        if (tagName == "$Date")
                                            pRange.Value = "Datum";
                                        else if (tagName == "$Time")
                                            pRange.Value = "Zeit";
                                        else if (tags.ContainsKey(tagName))
                                            pRange.Value = tags[tagName];
                                        else
                                            pRange.Value = tagName;
                                    }
                                }
                        }
                        else
                        {
                            for (int col = 0; col < items.Length; col++)
                            {
                                //Console.Write("Col=" + col);
                                using (ExcelRange pRange = pWorkSheet.Cells[startRow + row + 1, col + 1]) // Zeilen Id Start, Spalten Id Start[, Zeilen Id Ende, Spalten Id Ende]
                                {
                                    if (col == 0 && DateTime.TryParseExact(items[col], dateFormat, fromCulture, DateTimeStyles.None, out DateTime date)) //1. Spalte Datum                                                                                                                                                                                     
                                        SetDateToExcelCell(pRange, date);
                                    //else if (col == 1 && DateTime.TryParseExact(items[col], timeFormat, fromCulture, DateTimeStyles.None, out DateTime time)) //2. Spalte Uhrzeit
                                    else if (col == 1 && DateTime.TryParse(items[col], out DateTime time)) //2. Spalte Uhrzeit
                                        pRange.Value = time.ToShortTimeString();                                
                                    else
                                        SetNumberToExcelCell(pRange, items[col], !items[col].Contains(','));

                                    pRange.Style.Border.BorderAround(ExcelBorderStyle.Dotted);
                                }
                            }

                        }
                    }

                    startRow += lines.Length - 2;
#if DEBUG
                    Console.WriteLine("### Letzte Zeile " + startRow);
#endif

                    pWorkSheet.Cells[pWorkSheet.Dimension.Address].Style.Border.BorderAround(ExcelBorderStyle.Thin); //Zellenrand für alle
                    //pWorkSheet.Calculate()

                    //Make all text fit the cells
                    pWorkSheet.Cells[pWorkSheet.Dimension.Address].AutoFitColumns();
                }
                //Speichern
                System.IO.FileInfo fileInfo = new System.IO.FileInfo(ExcelFilePath);
                pPackage.SaveAs(fileInfo);
            }
        }


        private static void SetDateToExcelCell(ExcelRange pRange, DateTime pDateTime)
        {
            if (pRange == null) return;
            if (pDateTime.Equals(DateTime.MinValue)) return;


            //pRange.Style.Numberformat.Format = "yyyy-mm-dd";
            pRange.Style.Numberformat.Format = "dd.mm.yyyy";
            pRange.Value = pDateTime;
        }

        //private static void SetTimeToExcelCell(ExcelRange pRange, DateTime pDateTime)
        //{
        //    if (pRange == null) return;
        //    if (pDateTime.Equals(DateTime.MinValue)) return;

        //    pRange.Style.Numberformat.Format = "HH:mm";
        //    pRange.Value = pDateTime.ToShortTimeString();
        //}

        private static void SetNumberToExcelCell(ExcelRange pRange, string value, bool isInteger = false)
        {
            if (!isInteger && double.TryParse(value, out double number))
            {
                pRange.Style.Numberformat.Format = "0.0";
                pRange.Value = number;
            }
            else if (isInteger && int.TryParse(value, out int integer))
            {
                pRange.Style.Numberformat.Format = "0";
                pRange.Value = integer;
            }
            else
                pRange.Value = value;
        }


    }
}
