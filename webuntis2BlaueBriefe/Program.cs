// Published under the terms of GPLv3 Stefan Bäumer 2023
using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.IO;
using System.Linq;

namespace webuntis2BlaueBriefe
{
    class Program
    {
        public static string User = System.Security.Principal.WindowsIdentity.GetCurrent().Name.ToUpper().Split('\\')[1];
        public static string Folder = Environment.GetFolderPath(Environment.SpecialFolder.Desktop) + "\\BlaueBriefe-" + DateTime.Now.ToString("yyyyMMdd-hhmm");
        
        static void Main(string[] args)
        {
            System.IO.Directory.CreateDirectory(Folder);
            string steuerdatei = Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.Desktop), DateTime.Now.ToString("yyMMdd-HHmmss") + "_webuntisnoten2atlantis_" + System.Security.Principal.WindowsIdentity.GetCurrent().Name.ToUpper().Split('\\')[1] + ".csv");

            try
            {
                Console.WriteLine(" Webuntis2BlaueBriefe | Published under the terms of GPLv3 | Stefan Bäumer 2023 | Version 20230320");
                Console.WriteLine("====================================================================================================");
                Console.WriteLine("");

                string sourceMarksPerLesson = CheckFile(User, "MarksPerLesson");

                Periodes periodes = new Periodes();
                Leistungen alleDefizitäreWebuntisLeistungen = new Leistungen(sourceMarksPerLesson);

                var xxx = alleDefizitäreWebuntisLeistungen.GetInteressierendeKlassen();

                var defizitäreWebuntisLeistungen = new Leistungen();
                defizitäreWebuntisLeistungen.AddRange((from t in alleDefizitäreWebuntisLeistungen where xxx.Contains(t.Klasse) select t).ToList());

                Lehrers lehrers = new Lehrers(periodes);
                Klasses klasses = new Klasses(lehrers, periodes, defizitäreWebuntisLeistungen);
                
                Leistungen atlantisLeistungen = new Leistungen(Global.ConnectionStringAtlantis, defizitäreWebuntisLeistungen);
                
                Schuelers schuelerMitDefiziten = new Schuelers(defizitäreWebuntisLeistungen, atlantisLeistungen, klasses, lehrers);

                

                foreach (var sd in schuelerMitDefiziten)
                {
                    var hzAnzahl5en = (from s in sd.DefizitäreLeistungen where s.NoteHalbjahr == 5 select s).Count();
                    var hzAnzahl6en = (from s in sd.DefizitäreLeistungen where s.NoteHalbjahr == 6 select s).Count();
                    var jetztAnzahl5en = (from s in sd.DefizitäreLeistungen where s.NoteJetzt == 5 select s).Count();
                    var jetztAnzahl6en = (from s in sd.DefizitäreLeistungen where s.NoteJetzt == 6 select s).Count();
                    var nochWeitereDefiziteHinzugekommen = (from s in sd.DefizitäreLeistungen where s.NeueDefizitLeistung select s).Any();
                    var bereitsImHalbjahrGefährdet = (from s in sd.DefizitäreLeistungen where s.NoteHalbjahr >= 5 select s.NoteHalbjahr).Sum() >= 6 ? true : false;
                    var bereitsImHalbjahrEine5 = (from s in sd.DefizitäreLeistungen where s.NoteHalbjahr == 5 select s).Count() == 1 ? true : false;
                    var imHalbjahrKeinDefizit = (from s in sd.DefizitäreLeistungen where s.NoteHalbjahr >=5 select s.NoteHalbjahr).Any() ? false : true;
                    var verschlechterungvon5auf6 = (from s in sd.DefizitäreLeistungen where s.NoteHalbjahr == 5 where s.NoteJetzt == 6 select s.NoteHalbjahr).Any() ? true : false;

                    Console.WriteLine(sd.Klasse.PadRight(6) + sd.Nachname + "," + sd.Vorname + " ...");
                    Global.WriteLine(Folder,sd.Nachname + "," + sd.Vorname + ": " + (sd.Volljaehrig ? "Volljährig" : " Minderjährig") + ", " + sd.Klasse);

                    if (!nochWeitereDefiziteHinzugekommen && !verschlechterungvon5auf6)
                    {
                        Global.WriteLine(Folder, sd.Nachname + "," + sd.Vorname + ": keine weiteren Defizite seit dem Halbjahr, keine Mitteilung.");
                    }

                    // HZ: kein Defizit; 
                    
                    if (imHalbjahrKeinDefizit && nochWeitereDefiziteHinzugekommen)
                    {
                        Global.Write(Folder,"imHalbjahrKeinDefizit,");
                        
                        //jetzt eine 5: Mitteilung über Leistungsstand

                        if ((from s in sd.DefizitäreLeistungen where s.NeueDefizitLeistung select s.NoteJetzt).Sum() == 5)
                        {
                            Global.Write(Folder,"jetzt eine 5,");                                                        
                            sd.RenderMitteilung("M", Folder);
                        }

                        // HZ kein Defizit; jetzt zwei oder mehr 5: Gefährdung
                        // HZ kein Defizit; jetzt eine 6 oder mehr: Gefährdung

                        if ((from s in sd.DefizitäreLeistungen where s.NeueDefizitLeistung select s.NoteJetzt).Sum() >= 6)
                        {
                            Global.Write(Folder,"jetzt zwei oder mehr 5 oder eine 6,");
                            sd.RenderMitteilung("G", Folder);
                        }
                    }

                    // HZ eine 5; jetzt eine oder mehrere zusätzliche 5en: Gefährdung
                    // HZ eine 5; jetzt eine oder mehrere zusätzliche 6en: Gefährdung

                    if (bereitsImHalbjahrEine5 && nochWeitereDefiziteHinzugekommen)
                    {
                        Global.Write(Folder,"bereits im Halbjahr eine 5; jetzt eine o. mehrere zusätzliche 5en o. 6en;");
                        sd.RenderMitteilung("G", Folder);
                    }

                    // HZ eine 5; jetzt Verschlechterung auf 6: Gefährdung

                    if (bereitsImHalbjahrEine5 && verschlechterungvon5auf6 && !nochWeitereDefiziteHinzugekommen)
                    {
                        Global.Write(Folder,"im Hj genau eine 5, also bisher nicht gefährdet; jetzt 6;");
                        sd.RenderMitteilung("V", Folder);
                    }

                    // HZ: Zwei oder mehr 5en oder mindestens eine 6; jetzt eine oder zusätzliche 5en oder 6er: Gefährdung

                    if (bereitsImHalbjahrGefährdet && nochWeitereDefiziteHinzugekommen)
                    {
                        Global.Write(Folder, sd.Nachname + "," + sd.Vorname + ": " + "bereits im Halbjahr gefährdet; jetzt eine o. mehrere zusätzliche 5en o. 6en;");
                        sd.RenderMitteilung("G", Folder);
                    }

                    // Zeilen für alle gefährdeten Fächer andrucken

                    foreach (var item in (from d in sd.DefizitäreLeistungen select d))
                    {
                        Global.WriteLine(Folder, sd.Nachname + "," + sd.Vorname + ": " + item.Fach.PadRight(5) + item.NoteHalbjahr.ToString() + " => " + item.NoteJetzt);
                    }
                    Global.WriteLine(Folder, "");
                }

                Console.WriteLine("");
                Console.WriteLine("Verarbeitung beendet. ENTER");
                Process.Start(Folder);
                Console.ReadKey();
                
            }
            catch(IOException ex)
            {
                Console.WriteLine("");
                Console.WriteLine("");
                if (ex.ToString().Contains("bereits vorhanden"))
                {
                    Console.WriteLine("FEHLER: Die Datei existiert bereits. Bitte zuerst löschen. Dann erneut starten.");
                }
                else
                {
                    Console.WriteLine(ex);
                }            
                Console.ReadKey();
                Environment.Exit(0);
            }
            catch (Exception ex)
            {
                Console.WriteLine("Heiliger Bimbam! Es ist etwas schiefgelaufen! Die Verarbeitung wird gestoppt.");
                Console.WriteLine("");
                Console.WriteLine(ex);
                Console.ReadKey();
                Environment.Exit(0);
            }
        }

        private static void RenderNotenexportCsv(string inputNotenCsv)
        {
            Console.WriteLine("Die Datei " + inputNotenCsv + " existiert nicht.");
            Console.WriteLine("Exportieren Sie die Datei aus dem Digitalen Klassenbuch, indem Sie als Administrator");
            Console.WriteLine(" 1. Klassenbuch > Berichte klicken");
            Console.WriteLine(" 2. Alle Klassen auswählen und als Zeitraum das ganze Schuljahr wählen");
            Console.WriteLine(" 3. Unter \"Noten\" die Prüfungsart \"Alle\" auswählen.");
            Console.WriteLine(" 4. Hinter \"Noten pro Schüler\" auf CSV klicken");
            Console.WriteLine(" 5. Die Datei \"MarksPerLesson.csv\" auf dem Desktop speichern");
            Console.WriteLine("ENTER beendet das Programm");
            Console.ReadKey();
            Environment.Exit(0);
        }

        private static string CheckFile(string user, string kriterium)
        {
            var sourceFile = (from f in Directory.GetFiles(@"c:\users\" + user + @"\Downloads", "*.csv", SearchOption.AllDirectories) where f.Contains(kriterium) orderby File.GetLastWriteTime(f) select f).LastOrDefault();

            if ((sourceFile == null || System.IO.File.GetLastWriteTime(sourceFile).Date != DateTime.Now.Date))
            {
                Console.WriteLine("");
                Console.WriteLine(" Die " + kriterium + "<...>.csv" + (sourceFile == null ? " existiert nicht im Download-Ordner" : " im Download-Ordner ist nicht von heute. \n Es werden keine Daten aus der Datei importiert") + ".");
                Console.WriteLine(" Exportieren Sie die Datei frisch aus Webuntis, indem Sie als Administrator:");

                if (kriterium.Contains("MarksPerLesson"))
                {
                    Console.WriteLine("   1. Klassenbuch > Berichte klicken");
                    Console.WriteLine("   2. Alle Klassen auswählen und ggfs. den Zeitraum einschränken");
                    Console.WriteLine("   3. Unter \"Noten\" die Prüfungsart (-Alle-) auswählen");
                    Console.WriteLine("   4. Unter \"Noten\" den Haken bei Notennamen ausgeben _NICHT_ setzen");
                    Console.WriteLine("   5. Hinter \"Noten pro Schüler\" auf CSV klicken");
                    Console.WriteLine("   6. Die Datei \"MarksPerLesson<...>.CSV\" im Download-Ordner zu speichern");
                    Console.WriteLine(" ");
                    Console.WriteLine(" ENTER beendet das Programm.");
                    Console.ReadKey();
                    Environment.Exit(0);
                }

                if (kriterium.Contains("AbsenceTimesTotal"))
                {
                    Console.WriteLine("   1. Administration > Export klicken");
                    Console.WriteLine("   2. Zeitraum begrenzen, also die Woche der Zeugniskonferenz und vergange Abschnitte herauslassen");
                    Console.WriteLine("   2. Das CSV-Icon hinter Gesamtfehlzeiten klicken");
                    Console.WriteLine("   4. Die Datei \"AbsenceTimesTotal<...>.CSV\" im Download-Ordner zu speichern");
                }
                Console.WriteLine(" ");
                sourceFile = null;
            }

            if (sourceFile != null)
            {
                Console.WriteLine("Ausgewertete Datei: " + (Path.GetFileName(sourceFile) + " ").PadRight(53, '.') + ". Erstell-/Bearbeitungszeitpunkt heute um " + System.IO.File.GetLastWriteTime(sourceFile).ToShortTimeString());
            }

            return sourceFile;
        }
    }
}