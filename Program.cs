using System;
using System.Collections.Generic;
using System.IO;

namespace HistData2Excel
{
    internal partial class Program
    {
        internal static Dictionary<string, string> Tags = new Dictionary<string, string>();
        private static readonly string AppFolder = AppDomain.CurrentDomain.BaseDirectory;

        const string IniFile = "HistConfig.ini";
        const string CsvFile = "HistTags.csv";


        static void Main(string[] args)
        {
            #region Initialisierungswerte einlesen
            string iniPath = System.IO.Path.Combine(AppFolder, IniFile);
            if (!File.Exists(iniPath))
                CreateIni(iniPath);
            else
                ReadIni(iniPath);

            string csvPath = System.IO.Path.Combine(AppFolder, CsvFile);
            if (!File.Exists(csvPath))
                CreateCsv(csvPath);
            else
                Tags = ReadCsv(csvPath);
            #endregion

            string errorText = Dde.HistDataPoke(new List<string>(Tags.Keys));

            Console.WriteLine($"Das Programm gab folgenden Fehler zurück:\r\n" + errorText);

#if DEBUG
            foreach (var data in Dde.Data)
            {
                Console.WriteLine();
                Console.WriteLine(data);
                Console.WriteLine(new string('_', 32));
            }
#endif

            Excel.WriteNew(Tags, Dde.Data);

            Console.WriteLine("\r\nBeliebige Taste zum Beenden.");
            Console.ReadKey();
        }

        internal static void Countdown(int sec)
        {
            for (int i = sec; i > 0; i--)
            {
                Console.Write("{00}",i);
                System.Threading.Thread.Sleep(1000);
                Console.Write("\r");
            }
        }
    }
}
