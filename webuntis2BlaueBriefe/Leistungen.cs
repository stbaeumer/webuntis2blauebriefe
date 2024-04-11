// Published under the terms of GPLv3 Stefan Bäumer 2023

using System;
using System.Collections.Generic;
using System.Data;
using System.Data.Odbc;
using System.IO;
using System.Linq;
using System.Text.RegularExpressions;

namespace webuntis2BlaueBriefe
{
    public class Leistungen : List<Leistung>
    {
        public Leistungen(string sourceMarksPerLesson)
        {
            using (StreamReader reader = new StreamReader(sourceMarksPerLesson))
            {
                string überschrift = reader.ReadLine();

                int i = 1;

                Leistung leistung = new Leistung();

                while (true)
                {
                    string line = reader.ReadLine();

                    try
                    {
                        if (line != null)
                        {
                            var x = line.Split('\t');
                            i++;

                            if (x.Length == 10)
                            {
                                leistung = new Leistung();
                                leistung.Datum = DateTime.ParseExact(x[0], "dd.MM.yyyy", System.Globalization.CultureInfo.InvariantCulture);
                                leistung.Name = x[1];
                                leistung.Klasse = x[2];
                                leistung.Fach = x[3];
                                leistung.Prüfungsart = x[4];
                                leistung.NoteJetzt = GesamtPunkte2Gesamtnote(x[5]);
                                leistung.Bemerkung = x[6];
                                leistung.Benutzer = x[7];
                                leistung.SchlüsselExtern = Convert.ToInt32(x[8]);
                            }

                            // Wenn in den Bemerkungen eine zusätzlicher Umbruch eingebaut wurde:

                            if (x.Length == 7)
                            {
                                leistung = new Leistung();
                                leistung.Datum = DateTime.ParseExact(x[0], "dd.MM.yyyy", System.Globalization.CultureInfo.InvariantCulture);
                                leistung.Name = x[1];
                                leistung.Klasse = x[2];
                                leistung.Fach = x[3];
                                leistung.NoteJetzt = GesamtPunkte2Gesamtnote(x[5]);
                                leistung.Bemerkung = x[6];
                                Console.WriteLine("\n\n  [!] Achtung: In den Zeilen " + (i - 1) + "-" + i + " hat vermutlich die Lehrkraft eine Bemerkung mit einem Zeilen-");
                                Console.Write("      umbruch eingebaut. Es wird nun versucht trotzdem korrekt zu importieren ... ");
                            }

                            if (x.Length == 4)
                            {
                                leistung.Benutzer = x[1];
                                leistung.SchlüsselExtern = Convert.ToInt32(x[2]);
                                leistung.NoteJetzt = GesamtPunkte2Gesamtnote(x[3]);
                                Console.WriteLine("hat geklappt.\n");
                            }

                            if (x.Length < 4)
                            {
                                Console.WriteLine("\n\n[!] MarksPerLesson.CSV: In der Zeile " + i + " stimmt die Anzahl der Spalten nicht. Das kann passieren, wenn z. B. die Lehrkraft bei einer Bemerkung einen Umbruch eingibt. Mit Suchen & Ersetzen kann die Datei MarksPerLesson.CSV korrigiert werden.");
                                Console.ReadKey();
                                throw new Exception("\n\n[!] MarksPerLesson.CSV: In der Zeile " + i + " stimmt die Anzahl der Spalten nicht. Das kann passieren, wenn z. B. die Lehrkraft bei einer Bemerkung einen Umbruch eingibt. Mit Suchen & Ersetzen kann die Datei MarksPerLesson.CSV korrigiert werden.");
                            }

                            if (leistung.Prüfungsart == Global.BlaueBriefe)
                            {
                                if (leistung.Fach != null)
                                {
                                    this.Add(leistung);
                                }
                                else
                                {
                                    Console.WriteLine("ACHTUNG: Blauer Brief ohne Fach bei Zeile: " + i);
                                }
                            }
                        }
                    }
                    catch (Exception ex)
                    {
                        throw ex;
                    }

                    if (line == null)
                    {
                        break;
                    }
                }
                Console.WriteLine(("Defizitäre Webuntisleistungen " + ".".PadRight(this.Count / 150, '.')).PadRight(48, '.') + (" " + this.Count).ToString().PadLeft(4), '.');
                if (this.Count == 0)
                {
                    Console.WriteLine("Keine einzige Leistung wurde aus Webuntis augelesen.");
                    Console.ReadKey();
                }
            }
        }

        internal List<string> GetInteressierendeKlassen()
        {
            var vorbelegung = Global.List2String((from t in this select t.Klasse).Distinct().ToList(), ',');

            var interessierendeKlassen = new List<string>();

            try
            {
                do
                {
                    Console.Write("  Bitte eine Klasse wählen " + (vorbelegung == null ? "" : "(" + vorbelegung + ")") + " : ");

                    var x = Console.ReadLine();

                    List<string> xx = x.ToUpper().Replace(" ", "").Split(',').ToList();

                    if (vorbelegung != null)
                    {
                        xx.Add(vorbelegung);
                    }

                    foreach (var eingabe in xx)
                    {
                        foreach (var item in this)
                        {
                            if (item.Klasse != "" && (item.Klasse == eingabe || item.Klasse.StartsWith(eingabe)))
                            {
                                if (!interessierendeKlassen.Contains(item.Klasse))
                                {
                                    interessierendeKlassen.Add(item.Klasse);
                                }
                            }
                        }
                    }
                } while (interessierendeKlassen.Count == 0);
            }
            catch (Exception ex)
            {
                Console.WriteLine("Bei der Auswahl der interessierenden Klasse ist es zum Fehler gekommen. \n " + ex);
            }
            Console.WriteLine("   Ihre Auswahl: " + Global.List2String(interessierendeKlassen, ','));
            Console.WriteLine(" ");

            return interessierendeKlassen;
        }

            private int GesamtPunkte2Gesamtnote(string gesamtpunkte)
        {
            if (gesamtpunkte == "0.0")
            {
                return 6;
            }
            if (gesamtpunkte == "1.0")
            {
                return 5;
            }
            if (gesamtpunkte == "2.0")
            {
                return 5;
            }
            if (gesamtpunkte == "3.0")
            {
                return 5;
            }
            if (gesamtpunkte == "4.0")
            {
                return 4;
            }
            if (gesamtpunkte == "5.0")
            {
                return 4;
            }
            if (gesamtpunkte == "6.0")
            {
                return 4;
            }
            if (gesamtpunkte == "7.0")
            {
                return 3;
            }
            if (gesamtpunkte == "8.0")
            {
                return 3;
            }
            if (gesamtpunkte == "9.0")
            {
                return 3;
            }
            if (gesamtpunkte == "10.0")
            {
                return 2;
            }
            if (gesamtpunkte == "11.0")
            {
                return 2;
            }
            if (gesamtpunkte == "12.0")
            {
                return 2;
            }
            if (gesamtpunkte == "13.0")
            {
                return 1;
            }
            if (gesamtpunkte == "14.0")
            {
                return 1;
            }
            if (gesamtpunkte == "15.0")
            {
                return 1;
            }
            return 0;
        }

        internal Leistung GetKorrespondierendeAtlantisLeistung(Leistung webuntisLeistung)
        {
            var x = (from t in this
                     where t.SchlüsselExtern == webuntisLeistung.SchlüsselExtern
                     where t.Klasse == webuntisLeistung.Klasse
                     where t.Fach.Replace("  ", " ") == webuntisLeistung.Fach
                     select t).ToList();

            // Es muss exakt eine korrespondierende Leistung geben

            if (x.Count != 1)
            {
                Console.WriteLine("");
                Console.WriteLine("Die Leistung aus Webuntis (" + webuntisLeistung.Name + ", " + webuntisLeistung.Klasse + ", " + webuntisLeistung.Fach + ") kann keiner Atlantisleistung zugeordnet werden.");

                bool wiederholen = true;

                int index = 0;

                string eingabe;

                do
                {
                    int i = 0;

                    foreach (var item in (from t in this
                                          where t.SchlüsselExtern == webuntisLeistung.SchlüsselExtern
                                          where t.Klasse == webuntisLeistung.Klasse
                                          select t).ToList())
                    {
                        Console.WriteLine(i + 1 + ". " + item.Klasse + "|" + item.Fach);
                        i++;
                    }

                    Console.Write(" Bitte Zahl (0 = keine Zuordnung) 1 bis " + (from t in this
                                                                                where t.SchlüsselExtern == webuntisLeistung.SchlüsselExtern
                                                                                where t.Klasse == webuntisLeistung.Klasse
                                                                                select t).Count() + " eingeben: ");

                    try
                    {
                        eingabe = Console.ReadLine();
                        index = int.Parse(eingabe);

                        if (index >= 0 && index <= (from t in this
                                                    where t.SchlüsselExtern == webuntisLeistung.SchlüsselExtern
                                                    where t.Klasse == webuntisLeistung.Klasse
                                                    select t).Count())
                        {
                            wiederholen = false;
                        }
                    }
                    catch (Exception)
                    {
                    }
                } while (wiederholen);

                if (index == 0)
                {
                    Console.WriteLine("Das Fach " + webuntisLeistung.Fach + " wird nicht zugeordnet und verworfen.");
                }
                else
                {
                    Console.WriteLine("Ihre Auswahl: " + webuntisLeistung.Fach + " wird " + (from t in this
                                                                                             where t.SchlüsselExtern == webuntisLeistung.SchlüsselExtern
                                                                                             where t.Klasse == webuntisLeistung.Klasse
                                                                                             select t).ToList()[index - 1].Fach + " zugeordnet.");
                    (from t in this
                     where t.SchlüsselExtern == webuntisLeistung.SchlüsselExtern
                     where t.Klasse == webuntisLeistung.Klasse
                     select t).ToList()[index - 1].Fach = webuntisLeistung.Fach;
                    
                    return (from t in this
                            where t.SchlüsselExtern == webuntisLeistung.SchlüsselExtern
                            where t.Klasse == webuntisLeistung.Klasse
                            select t).ToList()[index - 1];
                }
                return null;
            }

            return x[0];
        }

        public Leistungen(string connetionstringAtlantis, Leistungen alleWebuntisLeistungen)
        {
            var interessierendeKlassen = (from w in alleWebuntisLeistungen select w.Klasse).Distinct().ToList();

            var abfrage = "";

            foreach (var iK in interessierendeKlassen)
            {
                var schuelersId = (from a in alleWebuntisLeistungen
                                   where a.Klasse == iK
                                   where a.Prüfungsart == Global.BlaueBriefe
                                   select a.SchlüsselExtern).Distinct().ToList();

                foreach (var schuelerId in schuelersId)
                {
                    abfrage += "(DBA.schueler.pu_id = " + schuelerId + @" AND  DBA.klasse.klasse = '" + iK + @"') OR ";
                }
            }
            try
            {
                abfrage = abfrage.Substring(0, abfrage.Length - 4);
            }
            catch (Exception)
            {
                Console.WriteLine("Kann es sein, dass für die Auswahl keine Datensätze vorliegen?");
            }
            try
            {
                using (OdbcConnection connection = new OdbcConnection(connetionstringAtlantis))
                {
                    DataSet dataSet = new DataSet();
                    OdbcDataAdapter schuelerAdapter = new OdbcDataAdapter(@"
SELECT DBA.noten_einzel.noe_id AS LeistungId,
DBA.noten_einzel.fa_id,
DBA.noten_einzel.kurztext AS Fach,
DBA.noten_einzel.zeugnistext AS Zeugnistext,
DBA.noten_einzel.s_note AS Note,
DBA.noten_einzel.punkte AS Punkte,
DBA.noten_einzel.punkte_12_1 AS Punkte_12_1,
DBA.noten_einzel.punkte_12_2 AS Punkte_12_2,
DBA.noten_einzel.punkte_13_1 AS Punkte_13_1,
DBA.noten_einzel.punkte_13_2 AS Punkte_13_2,
DBA.noten_einzel.s_tendenz AS Tendenz,
DBA.noten_einzel.s_einheit AS Einheit,
DBA.noten_einzel.ls_id_1 AS LehrkraftAtlantisId,
DBA.schueler.name_1 AS Nachname,
DBA.schueler.name_2 AS Vorname,
DBA.schueler.dat_geburt,
DBA.schueler.pu_id AS SchlüsselExtern,
DBA.schue_sj.s_religions_unterricht AS Religion,
DBA.schue_sj.dat_austritt AS ausgetreten,
DBA.schue_sj.vorgang_akt_satz_jn AS SchuelerAktivInDieserKlasse,
DBA.schue_sj.vorgang_schuljahr AS Schuljahr,
(substr(schue_sj.s_berufs_nr,4,5)) AS Fachklasse,
DBA.klasse.s_klasse_art AS Anlage,
DBA.klasse.jahrgang AS Jahrgang,
DBA.schue_sj.s_gliederungsplan_kl AS Gliederung,
DBA.noten_kopf.s_typ_nok AS HzJz,
DBA.noten_kopf.nok_id AS NOK_ID,
s_art_fach,
DBA.noten_kopf.s_art_nok AS Zeugnisart,
DBA.noten_kopf.bemerkung_block_1 AS Bemerkung1,
DBA.noten_kopf.bemerkung_block_2 AS Bemerkung2,
DBA.noten_kopf.bemerkung_block_3 AS Bemerkung3,
DBA.noten_kopf.dat_notenkonferenz AS Konferenzdatum,
DBA.klasse.klasse AS Klasse
FROM(((DBA.noten_kopf JOIN DBA.schue_sj ON DBA.noten_kopf.pj_id = DBA.schue_sj.pj_id) JOIN DBA.klasse ON DBA.schue_sj.kl_id = DBA.klasse.kl_id) JOIN DBA.noten_einzel ON DBA.noten_kopf.nok_id = DBA.noten_einzel.nok_id ) JOIN DBA.schueler ON DBA.noten_einzel.pu_id = DBA.schueler.pu_id
WHERE schue_sj.s_typ_vorgang = 'A' AND s_typ_nok = 'HZ' AND schue_sj.vorgang_schuljahr = '" + Global.AktSjAtlantis + @"'  AND
(  
  " + abfrage + @"
)
ORDER BY DBA.klasse.s_klasse_art DESC, DBA.noten_kopf.dat_notenkonferenz DESC, DBA.klasse.klasse ASC, DBA.noten_kopf.nok_id, DBA.noten_einzel.position_1; ", connection);


                    connection.Open();
                    schuelerAdapter.Fill(dataSet, "DBA.leistungsdaten");

                    string bereich = "";

                    foreach (DataRow theRow in dataSet.Tables["DBA.leistungsdaten"].Rows)
                    {
                        if (theRow["s_art_fach"].ToString() == "U")
                        {
                            bereich = theRow["Zeugnistext"].ToString();
                        }
                        else
                        {
                            DateTime austrittsdatum = theRow["ausgetreten"].ToString().Length < 3 ? new DateTime() : DateTime.ParseExact(theRow["ausgetreten"].ToString(), "dd.MM.yyyy HH:mm:ss", System.Globalization.CultureInfo.InvariantCulture);

                            Leistung leistung = new Leistung();

                            try
                            {
                                // Wenn der Schüler nicht in diesem Schuljahr ausgetreten ist ...

                                if (!(austrittsdatum > new DateTime(DateTime.Now.Month >= 8 ? DateTime.Now.Year : DateTime.Now.Year - 1, 8, 1) && austrittsdatum < DateTime.Now))
                                {
                                    leistung.LeistungId = Convert.ToInt32(theRow["LeistungId"]);
                                    leistung.SchlüsselExtern = Convert.ToInt32(theRow["SchlüsselExtern"]);

                                    leistung.ReligionAbgewählt = theRow["Religion"].ToString() == "N";
                                    leistung.Schuljahr = theRow["Schuljahr"].ToString();
                                    leistung.Gliederung = theRow["Gliederung"].ToString();
                                    leistung.HatBemerkung = (theRow["Bemerkung1"].ToString() + theRow["Bemerkung2"].ToString() + theRow["Bemerkung3"].ToString()).Contains("Fehlzeiten") ? true : false;
                                    leistung.Jahrgang = Convert.ToInt32(theRow["Jahrgang"].ToString().Substring(3, 1));
                                    leistung.Name = theRow["Nachname"] + " " + theRow["Vorname"];
                                    leistung.Nachname = theRow["Nachname"].ToString();
                                    leistung.Vorname = theRow["Vorname"].ToString();

                                    if ((theRow["LehrkraftAtlantisId"]).ToString() != "")
                                    {
                                        leistung.LehrkraftAtlantisId = Convert.ToInt32(theRow["LehrkraftAtlantisId"]);
                                    }
                                    leistung.Bereich = bereich;
                                    leistung.Geburtsdatum = theRow["dat_geburt"].ToString().Length < 3 ? new DateTime() : DateTime.ParseExact(theRow["dat_geburt"].ToString(), "dd.MM.yyyy HH:mm:ss", System.Globalization.CultureInfo.InvariantCulture);
                                    leistung.Volljährig = leistung.Geburtsdatum.AddYears(18) > DateTime.Now ? false : true;
                                    leistung.Klasse = theRow["Klasse"].ToString();
                                    leistung.Fach = theRow["Fach"] == null ? "" : theRow["Fach"].ToString();
                                    var gesamtnote = theRow["Note"].ToString() == "" ? null : theRow["Note"].ToString() == "Attest" ? "A" : theRow["Note"].ToString();
                                    try
                                    {
                                        leistung.NoteHalbjahr = Convert.ToInt32(gesamtnote);
                                    }
                                    catch (Exception)
                                    {
                                    }

                                    leistung.Gesamtpunkte = theRow["Punkte"].ToString() == "" ? null : (theRow["Punkte"].ToString()).Split(',')[0];
                                    leistung.Tendenz = theRow["Tendenz"].ToString() == "" ? null : theRow["Tendenz"].ToString();
                                    leistung.EinheitNP = theRow["Einheit"].ToString() == "" ? "N" : theRow["Einheit"].ToString();
                                    leistung.SchlüsselExtern = Convert.ToInt32(theRow["SchlüsselExtern"].ToString());
                                    leistung.HzJz = theRow["HzJz"].ToString();
                                    leistung.Anlage = theRow["Anlage"].ToString();
                                    leistung.Zeugnisart = theRow["Zeugnisart"].ToString();
                                    leistung.BezeichnungImZeugnis = theRow["Zeugnistext"].ToString();
                                    leistung.Konferenzdatum = theRow["Konferenzdatum"].ToString().Length < 3 ? new DateTime() : (DateTime.ParseExact(theRow["Konferenzdatum"].ToString(), "dd.MM.yyyy HH:mm:ss", System.Globalization.CultureInfo.InvariantCulture)).AddHours(15);


                                    if (
                                        leistung.Schuljahr == Global.AktSjAtlantis &&
                                        leistung.HzJz == "HZ" &&
                                        leistung.NoteHalbjahr != 0
                                        )
                                    {
                                        this.Add(leistung);
                                    }
                                }
                            }
                            catch (Exception ex)
                            {
                                Console.WriteLine("Fehler beim Einlesen der Atlantis-Leistungsdatensätze: ENTER" + ex);
                                Console.ReadKey();
                            }
                        }
                    }
                    connection.Close();
                }
            }
            catch (Exception ex)
            {
                throw ex;
            }

            var alleFächerDesAktuellenAbschnitts = (from l in this where (l.Konferenzdatum > DateTime.Now.AddDays(-20) || l.Konferenzdatum.Year == 1) select l.Fach).ToList();

        }

        public Leistungen()
        {
        }
    }
}