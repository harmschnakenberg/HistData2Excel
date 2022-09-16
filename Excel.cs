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

        public static void WriteNew(Dictionary<string, string> tags, string data)
        {
            string targetFile = System.IO.Path.Combine(Dde.TargetDir, $"HistData_{Dde.StartDate:yyyy-MM-dd_HH-mm}_{Dde.Duration}_{DateTime.Now.Ticks}.xlsx");
            Console.WriteLine("Erstelle die Datei " + targetFile);

            string[] lines = data.Split(new[] { '\r', '\n' }, StringSplitOptions.RemoveEmptyEntries);

            using (ExcelPackage pPackage = new ExcelPackage())
            {
                ExcelWorksheet pWorkSheet = pPackage.Workbook.Worksheets.Add("HistData_1");

                for (int row = 0; row < lines.Length; row++)
                {
                    // Console.Write("\r\nRow=" + row);
                    string[] items = lines[row].Split(';');

                    if (row == 0) //Erste Zeile
                    {
                        pWorkSheet.Cells[1, 1, lines.Length, items.Length].Style.Border.BorderAround(ExcelBorderStyle.Dotted); //Zellenrand für alle

                        for (int col = 0; col < items.Length; col++)
                        {
                            string tagName = items[col];

                            using (ExcelRange pRange = pWorkSheet.Cells[row + 1, col + 1]) // Zeilen Id Start, Spalten Id Start[, Zeilen Id Ende, Spalten Id Ende]
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
                            using (ExcelRange pRange = pWorkSheet.Cells[row + 1, col + 1]) // Zeilen Id Start, Spalten Id Start[, Zeilen Id Ende, Spalten Id Ende]
                            {
                                if (col == 0 && DateTime.TryParseExact(items[col], dateFormat, fromCulture, DateTimeStyles.None, out DateTime date)) //1. Spalte Datum                            
                                //if (col == 0 && DateTime.TryParse(items[col], out DateTime date)) //1. Spalte Datum    
                                    SetDateToExcelCell(pRange, date);
                                else if (col == 1 && DateTime.TryParse(items[col], out DateTime time)) //2. Spalte Uhrzeit
                                    SetTimeToExcelCell(pRange, time);
                                else
                                    SetNumberToExcelCell(pRange, items[col], !items[col].Contains(','));
                            }
                        }

                    }
                }

                //pWorkSheet.Calculate()

                //Make all text fit the cells
                pWorkSheet.Cells[pWorkSheet.Dimension.Address].AutoFitColumns();

                //Speichern
                System.IO.FileInfo fileInfo = new System.IO.FileInfo(targetFile);
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

        private static void SetTimeToExcelCell(ExcelRange pRange, DateTime pDateTime)
        {
            if (pRange == null) return;
            if (pDateTime.Equals(DateTime.MinValue)) return;

            //pRange.Style.Numberformat.Format = "yyyy-mm-dd";
            pRange.Style.Numberformat.Format = "hh:mm";
            pRange.Value = pDateTime;
        }

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
