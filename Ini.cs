using System;
using System.Collections.Generic;
using System.IO;

namespace HistData2Excel
{
    internal partial class Program
    {
        /// <summary>
        /// Erzeugt eine INI mit den Abfrage-Parametern
        /// </summary>
        /// <param name="IniPath"></param>
        private static void CreateIni(string IniPath)
        {

            if (File.Exists(IniPath)) return;

            Console.WriteLine("Erstelle neue INI-Datei {0}.", IniPath);

            using (StreamWriter w = File.AppendText(IniPath))
            {
                try
                {

                    w.WriteLine("[öäü " + w.Encoding.EncodingName + ", Build " + System.Reflection.Assembly.GetExecutingAssembly().GetName().Version + "]\r\n" +
                                "\r\n[Pfade]\r\n" +
                                ";#Pfad zum Logordner mit *.lgh-Dateien (z.B. D:\\Into_105\\Log)\r\n" +
                                $";{nameof(Dde.DataDir)}={Dde.DataDir}\r\n" +
                                ";#Pfad zum IntOuch-Projekt (z.B. D:\\Into_105\\MeinProjekt)\r\n" +
                                $";{nameof(Dde.DBDir)}={Dde.DBDir}\r\n" +
                                ";#Zielordner für CSV-Datei(en) (z.B. D:\\Archiv)\r\n" +
                                $";{nameof(Dde.TargetDir)}={Dde.TargetDir}\r\n" +

                                "\r\n[Zeiten]\r\n" +
                                ";#Beginn der Daten:\r\n" +
                                $";{nameof(Dde.StartDate)}={Dde.StartDate}\r\n" +
                                ";#Zeitraum bis zum Ende der Daten (w, d, h, m, s):\r\n" +
                                $";{nameof(Dde.Duration)}={Dde.Duration}\r\n" +
                                ";#Intervall der Datensätze (w, d, h, m, s):\r\n" +
                                $";{nameof(Dde.Interval)}={Dde.Interval}\r\n" +

                                "\r\n[Schalter]\r\n" +
                                ";#Rohdaten in CSV speichern:\r\n" +
                                $";{nameof(Dde.WriteToCsv)}={(Dde.WriteToCsv ? "1" : "0")}\r\n" +

                                 ""
                                );
                }
                catch
                {
                    Console.WriteLine("FEHLER beim Erstellen von {0}. Siehe Log.", IniPath);
                }
            }


        }

        /// <summary>
        /// Erzeugt eine CSV mmit den abzufragenden TagNames und Comments
        /// </summary>
        /// <param name="CsvPath"></param>
        private static void CreateCsv(string CsvPath)
        {
            if (File.Exists(CsvPath)) return;

            using (StreamWriter w = File.AppendText(CsvPath))
            {
                try
                {
                    w.WriteLine("$Date;Datum");
                    w.WriteLine("$Time;Uhrzeit");
                    w.WriteLine("MittelAussenT;Mittlere Außentemperatur");
                }
                catch
                {
                    Console.WriteLine("FEHLER beim Erstellen von {0}.", CsvPath);
                }
            }

        }

        /// <summary>
        /// Liest die Abfrageparameter aus der INI-Datei
        /// </summary>
        /// <param name="IniPath"></param>
        private static void ReadIni(string IniPath)
        {
            if (!File.Exists(IniPath))
                CreateIni(IniPath);

            try
            {
                #region Einlesen
                string configAll = System.IO.File.ReadAllText(IniPath, System.Text.Encoding.UTF8);
                char[] delimiters = new char[] { '\r', '\n' };
                string[] configLines = configAll.Split(delimiters, StringSplitOptions.RemoveEmptyEntries);
                Dictionary<string, string> dict = new Dictionary<string, string>();

                Console.WriteLine(configLines.Length + " Zeilen gelesen aus " + IniPath);

                foreach (string line in configLines)
                {
                    if (line[0] != ';' && line[0] != '[')
                    {
                        string[] item = line.Split('=');
                        string val = item[1].Trim();
                        if (item.Length > 1)
                        {
                            for (int n = 2; n < item.Length; n++)
                            {
                                val += "=" + item[n].Trim();
                            }
                        }
                        dict.Add(item[0].Trim(), val);

                        //Console.WriteLine(item[0].Trim() + "=" + val);
                    }
                }

                if (dict.Count == 0)
                {
                    Console.WriteLine("Es wurden keine Werte ausgelesen aus " + IniPath);
                    return;
                }
                #endregion

                #region Wertzuweisung
                if (dict.TryGetValue(nameof(Dde.DataDir), out string v) && v?.Length > 0 && Directory.Exists(v))
                    Dde.DataDir = v;

                if (dict.TryGetValue(nameof(Dde.DBDir), out v) && v?.Length > 0 && Directory.Exists(v))
                    Dde.DBDir = v;

                if (dict.TryGetValue(nameof(Dde.TargetDir), out v) && v?.Length > 0 && Directory.Exists(v))
                    Dde.TargetDir = v;

                if (dict.TryGetValue(nameof(Dde.StartDate), out v) && v?.Length > 0 && DateTime.TryParse(v, out DateTime startDate))
                    Dde.StartDate = startDate;

                if (dict.TryGetValue(nameof(Dde.Duration), out v) && v?.Length > 0)
                    Dde.Duration = v;

                if (dict.TryGetValue(nameof(Dde.Interval), out v) && v?.Length > 0)
                    Dde.Interval = v;

                if (dict.TryGetValue(nameof(Dde.WriteToCsv), out v) && v?.Length > 0 && int.TryParse(v, out int i))
                    Dde.WriteToCsv = i != 0;

                #endregion
            }
            catch
            {
                Console.WriteLine("FEHLER beim Lesen von {0}.", IniPath);
            }

            //Console.WriteLine("Werte von {0} eingelesen.", IniPath);
        }

        /// <summary>
        /// Liest die abzufragenden Tags aus der CSV-Datei
        /// </summary>
        /// <param name="csvPath"></param>
        /// <returns></returns>
        private static Dictionary<string, string> ReadCsv(string csvPath)
        {
            Dictionary<string, string> dict = new Dictionary<string, string>();

            try
            {
                if (!System.IO.File.Exists(csvPath))
                    CreateCsv(csvPath);

                //Console.WriteLine("Lese aus {0}.", csvPath);

                using (StreamReader reader = new StreamReader(csvPath))
                {
                    string line;

                    while ((line = reader.ReadLine()) != null)
                    {
                        var x = line.Split(';');
                        if (x.Length == 2)
                            dict.Add(x[0], x[1]);
                    }
                }
            }
            catch
            {
                Console.WriteLine("FEHLER beim Lesen von {0}.", csvPath);
            }

            Console.WriteLine(dict.Count + " Zeilen gelesen aus " + csvPath);

            return dict;
        }


    }
}
