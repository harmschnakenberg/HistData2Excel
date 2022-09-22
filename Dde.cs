using NDde.Client;
using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Text;

namespace HistData2Excel
{
    internal static class Dde
    {


        #region Const

        const string _DATADIR = "DATADIR";
        const string _DBDIR = "DBDIR";
        const string _STARTDATE = "STARTDATE";
        const string _STARTTIME = "STARTTIME";
        const string _DURATION = "DURATION";
        const string _INTERVAL = "INTERVAL";
        const string _FILENAME = "FILENAME";
        const string _WRITEFILE = "WRITEFILE";
        const string _ERROR = "ERROR";
        const string _PRINTTAGNAMES = "PRINTTAGNAMES";
        const string _DATA = "DATA";
        const string _STATUS = "STATUS";
        const string _SENDDATA = "SENDDATA";

        #endregion


        #region Properties
        public static string DdeServer { get; set; } = "HISTDATA";

        public static string DdeTopic { get; set; } = "TOPIC";

        private static string _DataDir = @"D:\Into_105\Log";
        public static string DataDir
        {
            get
            {
                return _DataDir;
            }
            set
            {
                if (System.IO.Directory.Exists(value))
                {
                    _DataDir = value;
                }
            }
        }

        private static string _DBDir = @"D:\Into_105\MeinProjekt";

        public static string DBDir
        {
            get
            {
                return _DBDir;
            }
            set
            {
                if (System.IO.Directory.Exists(value))
                {
                    _DBDir = value;
                }
            }
        }

        private static string _TargetDir = @"D:\Archiv";

        public static string TargetDir
        {
            get
            {
                return _TargetDir;
            }
            set
            {
                if (System.IO.Directory.Exists(value))
                {
                    _TargetDir = value;
                }
            }
        }

        internal static string TargetFile { get; set; } = string.Empty;

        public static DateTime StartDate { get; set; } = DateTime.Now.AddDays(-7);

        public static string Duration { get; set; } = @"1w";

        public static string Interval { get; set; } = @"1h";

        public static bool WriteToCsv { get; set; } = false;

        public static List<string> Data { get; set; } = new List<string>();

        private static readonly string WonderwareFolder = System.IO.Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.ProgramFiles), @"Wonderware\InTouch");
        public static string HistDataExePath { get; set; } = System.IO.Path.Combine(WonderwareFolder, "histdata.exe");

        #endregion

        /// <summary>
        /// Abfragen, die nacheinander ausgeführt werden sollen: <STARTDATE><DURATION>
        /// </summary>
        internal static Queue<TimeChunck> Inqueries = new Queue<TimeChunck>();


        /// <summary>
        /// Bereitet alles für eine Anfrage bei HistData.exe vor und führt sie aus. Daten werden später in HistDataRead() ausgelesen.
        /// </summary>
        /// <param name="tagNames">Anzufragende InTouch-TagNames</param>
        /// <returns>Fehlertext</returns>
        public static string HistDataPoke(List<string> tagNames)
        {
            string returnText = "Es ist ein Fehler aufgetreten";

            try
            {
                #region Prüfe, ob HistData.exe läuft und ggf. starten
                if (Process.GetProcessesByName(DdeServer).Length == 0)
                {
                    Console.WriteLine(DdeServer + ".exe läuft nicht.");

                    if (System.IO.File.Exists(HistDataExePath))
                    {
                        Console.WriteLine("Starte " + HistDataExePath);
                        System.Diagnostics.Process.Start(HistDataExePath);

                        Program.
                            );
                    }
                    else
                        Console.WriteLine(HistDataExePath + " konnte nicht gefunden werden.");
                }
                #endregion

                using (DdeClient gDdeClient = new DdeClient(DdeServer, DdeTopic))
                {
                    #region Connect
                    gDdeClient.Connect();

                    if (!gDdeClient.IsConnected)//(!gDdeClient.IsConnected && gDdeClient.TryConnect() != 0)
                    {
                        Console.WriteLine($"DDE-Verbindung zu {DdeServer}|{DdeTopic} konnte nicht aufgebaut werden. TryConnect = " + gDdeClient.TryConnect());
                        return returnText;
                    }
                    #endregion

                   
                    if (gDdeClient.IsConnected || gDdeClient.TryConnect() == 0)
                    {
                        Inqueries.Enqueue(new TimeChunck(StartDate, Duration));
                        
                        gDdeClient.Setup(tagNames);

                        do
                        {
                            TimeChunck timeChunck = Inqueries.Dequeue();
                            
                            if (Dde.IsDurationOk(timeChunck))
                            {
                                Console.WriteLine($"\r\nBereite Abfrage vor ab {timeChunck.StartDate} über {timeChunck.Duration}.");

                                gDdeClient.SetupTime(timeChunck);
                                if (gDdeClient.StartDataCollection())
                                    gDdeClient.HarvestData();
                                //Program.Countdown(2);
                            }

                        } while (Inqueries.Count > 0);

                    }
                    else
                        Console.WriteLine($"DDE-Verbindung zu {DdeServer}|{DdeTopic} konnte nicht aufgebaut werden.");

                    gDdeClient.Disconnect();

                    return returnText;
                }
            }

            catch (Exception ex)
            {
                Console.WriteLine("ES IST EIN FEHLER AUFGETRETEN:\r\n" + ex);
                return ex.Message;
            }
        }


        /// <summary>
        /// Schreibt vorbereitend Parameter an HistData.exe
        /// Zeitkritische Parameter werden in SetupTime() gesetzt.
        /// </summary>
        /// <param name="gDdeClient"></param>
        /// <param name="tagNames"></param>
        private static void Setup(this DdeClient gDdeClient, List<string> tagNames)
        {

            if (gDdeClient.TryPoke(_DATADIR, Encoding.Default.GetBytes(DataDir), 1, 10000) == 0)
                Console.WriteLine(_DATADIR + "\t\t=" + gDdeClient.Request(_DATADIR, 1000));
            else
                Console.WriteLine(_DATADIR + " konnte nicht gesetzt werden.");

            if (gDdeClient.TryPoke(_DBDIR, Encoding.Default.GetBytes(DBDir), 1, 10000) == 0)
                Console.WriteLine(_DBDIR + "\t\t=" + gDdeClient.Request(_DBDIR, 1000));
            else
                Console.WriteLine(_DBDIR + " konnte nicht gesetzt werden.");

            if (gDdeClient.TryPoke(_INTERVAL, Encoding.Default.GetBytes(Interval), 1, 10000) == 0)
                Console.WriteLine(_INTERVAL + "\t=" + gDdeClient.Request(_INTERVAL, 1000));
            else
                Console.WriteLine(_INTERVAL + " konnte nicht gesetzt werden.");

            Console.WriteLine();


            #region TagNames zusammenbauen
            /*
                Beispiel: 
                TAGS="$Date,ProdStufe,ProdTemp,+"
                TAGS1="ReaktStufe,Temp,GasStufe,+"
                TAGS2="MotorStatus" 
             */

            string tagString = "$DATE,$TIME";
            string tagString1 = string.Empty;
            string tagString2 = string.Empty;

            foreach (var tagName in tagNames)
            {
                if (tagString.Length + tagName.Length + 1 < 254)
                    tagString += "," + tagName;
                else if (tagString1.Length + tagName.Length + 1 < 254)
                    tagString1 += "," + tagName;
                else if (tagString2.Length + tagName.Length + 1 < 254)
                    tagString2 += "," + tagName;
            }

            tagString = "'" + tagString + (tagString1.Length > 1 ? "+'" : "'");
            gDdeClient.Poke("TAGS", tagString, 60000);

            if (tagString1.Length > 0)
            {
                tagString1 = "'" + tagString1.Remove(0, 1) + (tagString2.Length > 1 ? "+'" : "'");
                gDdeClient.Poke("TAGS1", tagString1, 60000);
            }

            if (tagString2.Length > 0)
            {
                tagString2 = "'" + tagString2.Remove(0, 1) + "'";
                gDdeClient.Poke("TAGS2", tagString2, 60000);
            }

            Console.WriteLine("TAGS=" + tagString);
            
            if(tagString1.Length > 3)
                Console.WriteLine("TAGS1=" + tagString1);
            
            if (tagString2.Length > 3) 
                Console.WriteLine("TAGS2=" + tagString2);

            #endregion

        }


        /// <summary>
        /// Setzt die zeitkritischen Parameter für StartDate / StartTime / Duration in HistData.exe
        /// </summary>
        /// <param name="gDdeClient"></param>
        /// <param name="startDate"></param>
        /// <param name="interval"></param>
        private static void SetupTime(this DdeClient gDdeClient, TimeChunck timeChunck)
        {
            Duration = timeChunck.Duration;
            StartDate = timeChunck.StartDate;

            if (gDdeClient.TryPoke(_STARTDATE, Encoding.Default.GetBytes(timeChunck.StartDate.ToShortDateString()), 1, 10000) == 0)
                Console.WriteLine(_STARTDATE + "\t=" + gDdeClient.Request(_STARTDATE, 1000));
            else
                Console.WriteLine(_STARTDATE + " konnte nicht gesetzt werden.");

            if (gDdeClient.TryPoke(_STARTTIME, Encoding.Default.GetBytes(timeChunck.StartDate.ToShortTimeString()), 1, 10000) == 0)
                Console.WriteLine(_STARTTIME + "\t=" + gDdeClient.Request(_STARTTIME, 1000));
            else
                Console.WriteLine(_STARTTIME + " konnte nicht gesetzt werden.");

            if (gDdeClient.TryPoke(_DURATION, Encoding.Default.GetBytes(timeChunck.Duration), 1, 10000) == 0)
                Console.WriteLine(_DURATION + "\t=" + gDdeClient.Request(_DURATION, 1000));
            else
                Console.WriteLine(_DURATION + " konnte nicht gesetzt werden.");

            if (WriteToCsv)
            {
                TargetFile = System.IO.Path.Combine(TargetDir, $"HistData_{timeChunck.StartDate:yyyy-MM-dd_HH-mm}_{Duration}_{DateTime.Now.Ticks}.csv");

                if (gDdeClient.TryPoke(_FILENAME, Encoding.Default.GetBytes(TargetFile), 1, 10000) == 0)
                    Console.WriteLine(_FILENAME + "\t=" + gDdeClient.Request(_FILENAME, 1000));
                else
                    Console.WriteLine(_FILENAME + " konnte nicht gesetzt werden.");
            }
        }


        /// <summary>
        /// Startet die Zusammenstellung der Daten durch HistData.exe.
        /// Verringert bei entsprechender Fehlermeldung von HistData.exe den Datenzeitraum DURATION und startet die Zusammenstellung erneut.
        /// Verzögert den Programmablauf ggf. wenn von HistData.exe noch keine 'bereit'-Meldung da ist.
        /// </summary>
        /// <param name="gDdeClient"></param>
        /// <returns>true = Datenabfrage ist durchgelaufen. false = Datenabfrage wurde verworfen und neue Abfrage mit kürzerem Zeitraum erstellt.</returns>
        private static bool StartDataCollection(this DdeClient gDdeClient)
        {
            /*
            "SENDDATA" Löst die Bereitstellung der angeforderten Daten im Item DATA aus. 
            Wenn dieser Wert auf 1 gesetzt wird, werden die angeforderten Daten in das Item DATA geschrieben. 
            Nach Abschluss des Vorgangs wird SENDDATA automatisch auf 0 zurückgesetzt.

            Wenn Sie eine Fehlermeldung erhalten, die besagt, dass Sie mit SENDDATA zu viele Daten angefordert haben, 
            kürzen Sie DURATION oder verringern Sie die Anzahl der angeforderten Variablen. 
            Jeder Variablenname darf nur einmal angegeben werden.Jedes dieser Items darf maximal 512 Byte enthalten.
            */

            //vorsichtshalber SendData rücksetzen?
            if (gDdeClient.Request(_SENDDATA, 10000) == "1")
                gDdeClient.Poke(_SENDDATA, "0", 10000);

            //vorsichtshalber Status rücksetzen?
            gDdeClient.TryPoke(_STATUS, Encoding.Default.GetBytes("0"), 1, 10000);

            //Datensammlung in HistData.exe anstoßen
            if (gDdeClient.TryPoke(_SENDDATA, Encoding.Default.GetBytes("1"), 1, 10000) != 0)
                Console.WriteLine(_SENDDATA + ": Tags konnten nicht angefragt werden.");

            #region Warten auf Datenbereitstellung
            for (int i = 0; i < 5; i++)
            {
                //Console.WriteLine("Warte " + i);
                Program.Countdown(3); // DAS KOSTET ZEIT!! Reduzieren?

               string fehlerText = gDdeClient.Request(_ERROR, 10000);
                //Beispiel:
                //Zu viele Daten angefordert -- bitte verringern Sie die Dauer oder die Anzahl der Variablen
                if (!fehlerText.Contains("None"))
                    Console.WriteLine($"Fehlertext: '{fehlerText.Replace(Environment.NewLine," ")}'");

                #region Abbruch und neue Anfrage bei "Zu viele Daten angefordert..."
                if (fehlerText.StartsWith("Zu viele Daten angefordert"))
                {
                    //Neue Abfrage mit kürzerer DURATION hinzufügen und diese Abfrage abbrechen.
                    Dde.DivideDuration(StartDate, Duration);                    
                    return false;
                }
                #endregion

                //Nach Abschluss des Lese-Vorgangs wird SENDDATA automatisch auf 0 zurückgesetzt.
                if (gDdeClient.TryRequest(_SENDDATA, 1, 10000, out byte[] sendData) == 0)
                    if (Encoding.Default.GetString(sendData) == "0")
                        Console.WriteLine("Datenbereitstellung von Server beendet.");
                    else
                        Console.WriteLine("Daten werden bereitgestellt...");

                //"STATUS" Der Status des letzten HistData-Vorgangs.
                //1 bedeutet, dass die Daten erfolgreich aus den Archivdateien abgerufen wurden. 0 bedeutet, dass ein Fehler aufgetreten ist.
                if (gDdeClient.TryRequest(_STATUS, 1, 10000, out byte[] status) == 0)
                    if (Encoding.Default.GetString(status) == "0")
                        Console.WriteLine("Bei der Datenbereitstellung ist ein Fehler aufgetreten.");
                    else
                    {
                        Console.WriteLine("Daten bereit zur Abfrage.");
                        break;
                    }
                #endregion
            }

            return true;
        }


        /// <summary>
        /// Ruft dei zuvor von HistData.exe bereitgestellten Daten im CSV-Fpormat ab.
        /// </summary>
        /// <param name="gDdeClient"></param>
        private static void HarvestData(this DdeClient gDdeClient)
        {
            #region Datenabfrage

            string fehlerText = gDdeClient.Request(_ERROR, 10000);

            if (!fehlerText.Contains("None"))
                Console.WriteLine("Fehler vor Auslesen der Daten: " + fehlerText);
            else if (gDdeClient.TryRequest(_DATA, 1, 120000, out byte[] data) == 0)
            {
                Data.Add(Encoding.Default.GetString(data));
                Console.WriteLine($"Das Programm hat {data?.Length} Zeichen ausgelesen.");                
            }
            else
                Console.WriteLine("Beim Auslesen von DATA ist ein Fehler aufgetreten: " + gDdeClient.Request(_ERROR, 10000));

            #endregion

            #region CSV-Datei erstellen

            if (WriteToCsv)
            {
                gDdeClient.Poke(_PRINTTAGNAMES, "1", 10000);
                if (gDdeClient.TryPoke(_WRITEFILE, Encoding.Default.GetBytes("1"), 1, 10000) == 0)
                    if (gDdeClient.Request(_WRITEFILE, 10000) == "0")
                        Console.WriteLine($"Schreiben der CSV-Datei '{TargetFile}' wurde angestoßen.");
            }

            #endregion
        }


        /// <summary>
        /// Teilt die Anfrage in mehere Anfragen auf.
        /// </summary>
        /// <returns></returns>
        internal static void DivideDuration(DateTime startDate1, string duration1)
        {            
            int val1 = ParseDuration(duration1, out string unit);
            string duration2;
            DateTime startDate2;

            // Console.WriteLine("Zahlwert" +  val1);

            if (val1 > 1) //Aufteilen in meherer Abfragen
            {
                for (int i = 0; i < val1; i++)
                {
                    duration1 = $"1{unit}";

                    Inqueries.Enqueue(new TimeChunck(startDate1, duration1));

                    switch (unit.ToLower())
                    {
                        case "z": //Monat
                            startDate1 = startDate1.AddMonths(1);
                            break;
                        case "w":
                            startDate1 = startDate1.AddDays(7);
                            break;
                        case "d":
                            startDate1 = startDate1.AddDays(1);
                            break;
                        case "h":
                            startDate1 = startDate1.AddHours(1);
                            break;
                        default:
                            startDate1 = startDate1.AddDays(1);
                            break;
                    }
                }
            }
            else
            {
                #region Neue Unit setzen
                switch (unit.ToLower())
                {
                    case "z": //Monat
                        duration1 = "2w";
                        startDate2 = startDate1.AddDays(14);
                        duration2 = "2w";
                        break;
                    case "w":
                        duration1 = "4d";
                        startDate2 = startDate1.AddDays(4);
                        duration2 = "3d";
                        break;
                    case "d":
                        duration1 = "12h";
                        startDate2 = startDate1.AddHours(12);
                        duration2 = "12h";
                        break;
                    case "h":
                        duration1 = "30m";
                        startDate2 = startDate1.AddMinutes(30);
                        duration2 = "30m";
                        break;
                    default:
                        duration1 = "1d";
                        startDate2 = startDate1.AddDays(1);
                        duration2 = "1d";
                        break;
                }
                #endregion

                Inqueries.Enqueue(new TimeChunck(startDate1, duration1));
                Inqueries.Enqueue(new TimeChunck(startDate2, duration2));
            }

            Console.WriteLine(Inqueries.Count + " Anfragen in der Pipeline.");
        }

        /// <summary>
        /// Checkt, ob die DURATION offensichtlich zu lang ist und teilt die Abfrage ggf. auf.
        /// </summary>
        /// <param name="startDate1"></param>
        /// <param name="duration1"></param>
        /// <returns>false= Duration ist definitiv zu lang und wird in zwei kürzere Abfragen aufgteilt.</returns>
        internal static bool IsDurationOk(TimeChunck timeChunck)
        {            
            int val1 = ParseDuration(timeChunck.Duration, out string unit);

            #region Wenn DURATION zu lang (bei Test >6w, >42d), wird nur ein Teil der Daten abgefragt
            bool IsTooLong = false;

            switch (unit.ToLower())
            {
                case "z": //Monat
                    IsTooLong = true;
                    break;
                case "w":
                    IsTooLong = val1 > 6;
                    break;
                case "h":
                    IsTooLong = val1 > 6 * 7;
                    break;
                case "m":
                    IsTooLong = val1 > 6 * 7 * 60;
                    break;
            }
            #endregion

            if (IsTooLong)
            {
                Console.WriteLine($"Der Zeitraum {timeChunck.Duration} ist zu lang und wird aufgeteilt.");
                DivideDuration(timeChunck.StartDate, timeChunck.Duration);
            }

            return !IsTooLong;
        }

        private static int ParseDuration(string duration1, out string unit)
        {
            string strOrg = duration1;
            string strNumber = string.Empty;
            int val1 = 0;
            string unit1 = string.Empty;

            #region Aus DURATION Zahlwert und Einheit auslesen
            for (int i = 0; i < strOrg.Length; i++)
            {
                if (char.IsDigit(strOrg[i]))
                    strNumber += strOrg[i];
                else
                    unit1 = strOrg[i].ToString();

                if (strNumber.Length > 0)
                    val1 = int.Parse(strNumber);
            }
            #endregion
            unit = unit1;

            return val1;
        }

    }


    public class TimeChunck
    {
        public TimeChunck(DateTime startDate, string interval)
        {
            StartDate = startDate;
            Duration = interval ?? "1h";
        }

        public DateTime StartDate { get; set; }

        public string Duration { get; set; }
    }
}
