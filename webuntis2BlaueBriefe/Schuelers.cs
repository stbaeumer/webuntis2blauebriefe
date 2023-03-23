// Published under the terms of GPLv3 Stefan Bäumer 2023
using System;
using System.Collections.Generic;
using System.Configuration;
using System.Data;
using System.Data.Odbc;
using System.Diagnostics;
using System.IO;
using System.Linq;


namespace webuntis2BlaueBriefe
{
    internal class Schuelers : List<Schueler>
    {
        public Schuelers()
        {
        }

        public Schuelers(Leistungen defizitäreWebuntisLeistungen, Leistungen defizitäreAtlantisLeistungen, Klasses klasses, Lehrers lehrers)
        {
            foreach (var idAtlantis in (from t in defizitäreWebuntisLeistungen
                                        where t.Prüfungsart.StartsWith("Mahnung")
                                        select t.SchlüsselExtern).Distinct().ToList())
            {
                using (OdbcConnection connection = new OdbcConnection(Global.ConnectionStringAtlantis))
                {
                    connection.Open();

                    if (connection != null)
                    {
                        OdbcCommand command = connection.CreateCommand();
                        command.CommandText = @"SELECT 
DBA.schue_sj.pu_id as ID,
DBA.schue_sj.s_jahrgang AS Jahrgang,
DBA.adresse.s_typ_adr as Typ,
DBA.klasse.klasse as Klasse,
DBA.schueler.name_1 as Vorname,
DBA.schueler.name_2 as Nachname,
DBA.adresse.name_2 AS EVorname,
DBA.adresse.name_1 AS ENachname,
DBA.schueler.dat_geburt as Geburtsdatum,
DBA.schueler.s_geschl as Geschlecht,
DBA.adresse.strasse AS Strasse,
DBA.adresse.plz AS Plz,
DBA.adresse.ort AS Ort,
DBA.adresse.sorge_berechtigt_jn,
DBA.adresse.s_anrede,
DBA.schueler.s_erzb_1_art,
DBA.schueler.s_erzb_2_art,
DBA.schueler.id_hauptadresse,
DBA.adresse.hauptadresse_jn,
DBA.adresse.anrede_text,
DBA.schueler.anrede_text,
DBA.adresse.name_3,
DBA.adresse.plz_postfach as PlzPostfach,
DBA.adresse.postfach as Postfach,
DBA.adresse.s_titel_ad,
DBA.adresse.s_sorgerecht,
DBA.adresse.brief_adresse,
DBA.schue_sj.kl_id, 
DBA.adresse.s_famstand_adr
FROM((DBA.schue_sj JOIN DBA.klasse ON DBA.schue_sj.kl_id = DBA.klasse.kl_id) JOIN DBA.schueler ON DBA.schue_sj.pu_id = DBA.schueler.pu_id) JOIN DBA.adresse ON DBA.schueler.pu_id = DBA.adresse.pu_id
WHERE vorgang_schuljahr = '" + Global.AktSjAtlantis + @"' AND schue_sj.pu_id = " + idAtlantis + ";";

                        OdbcDataReader reader = command.ExecuteReader();

                        int fCount = reader.FieldCount;

                        var x = (from s in this where s.IdAtlantis == idAtlantis select s).FirstOrDefault();

                        Schueler schueler = new Schueler();
                        schueler.Sorgeberechtigte = new Sorgeberechtigte();

                        while (reader.Read())
                        {
                            schueler.IdAtlantis = Convert.ToInt32(reader.GetValue(0));
                            var sorgeberechtigtJn = reader.GetValue(13).ToString();
                            var anrede = reader.GetValue(14).ToString();
                            var typ = reader.GetValue(2).ToString();

                            if (sorgeberechtigtJn == "N" && typ == "0")
                            {
                                schueler.Vorname = reader.GetValue(5).ToString();                                
                                schueler.Nachname = reader.GetValue(4).ToString();
                                schueler.Strasse = reader.GetValue(10).ToString();
                                schueler.Plz = reader.GetValue(11).ToString();
                                schueler.Ort = reader.GetValue(12).ToString();
                                schueler.Geburtsdatum = DateTime.ParseExact(reader.GetValue(8).ToString(), "dd.MM.yyyy HH:mm:ss", System.Globalization.CultureInfo.InvariantCulture);
                                schueler.Volljaehrig = schueler.Geburtsdatum.AddYears(18) > DateTime.Now ? false : true;
                                schueler.Geschlecht = reader.GetValue(9).ToString();
                                schueler.Jahrgang = reader.GetValue(1).ToString();
                                schueler.Anrede = anrede;
                                schueler.Typ = reader.GetValue(2).ToString(); // 0 = Schüler V = Vater  M = Mutter
                                schueler.Klasse = reader.GetValue(3).ToString();
                                schueler.Klassenleitung = (from k in klasses where k.NameUntis == schueler.Klasse select k.Klassenleitungen[0].Vorname + " " + k.Klassenleitungen[0].Nachname).FirstOrDefault();
                                schueler.KlassenleitungMail = (from k in klasses where k.NameUntis == schueler.Klasse select k.Klassenleitungen[0].Mail).FirstOrDefault();
                                schueler.KlassenleitungMw = (from l in lehrers where l.Vorname + " " + l.Nachname == schueler.Klassenleitung select l.Anrede).FirstOrDefault();

                            }

                            if (sorgeberechtigtJn == "J" && anrede == "H")
                            {
                                var vorname = reader.GetValue(6).ToString();
                                var nachname = reader.GetValue(7).ToString();
                                var strasse = reader.GetValue(10).ToString();
                                var plz = reader.GetValue(11).ToString();
                                var ort = reader.GetValue(12).ToString();
                                schueler.Sorgeberechtigte.Add(new Sorgeberechtigt(vorname, nachname, strasse, plz, ort));
                            }

                            if (sorgeberechtigtJn == "J" && anrede == "F")
                            {
                                var vorname = reader.GetValue(6).ToString();
                                var nachname = reader.GetValue(7).ToString();
                                var strasse = reader.GetValue(10).ToString();
                                var plz = reader.GetValue(11).ToString();
                                var ort = reader.GetValue(12).ToString();
                                schueler.Sorgeberechtigte.Add(new Sorgeberechtigt(vorname, nachname, strasse, plz, ort));
                            }
                        }

                        schueler.DefizitäreLeistungen = new Leistungen();

                        // Defizitäre Leistungen kommen aus Webuntis ...

                        foreach (var wl in (from t in defizitäreWebuntisLeistungen
                                            where t.Prüfungsart.StartsWith("Mahnung")
                                            where t.SchlüsselExtern == schueler.IdAtlantis
                                            select t).ToList())
                        {
                            var al = defizitäreAtlantisLeistungen.GetKorrespondierendeAtlantisLeistung(wl);

                            if (al != null)
                            {
                                wl.NoteHalbjahr = al.NoteHalbjahr;
                                wl.BezeichnungImZeugnis = al.BezeichnungImZeugnis;
                                wl.NeueDefizitLeistung = wl.NoteHalbjahr <= 4 && wl.NoteJetzt >= 5 ? true : false;
                                wl.NochmaligeVerschlechterungAuf6 = wl.NoteHalbjahr == 5 && wl.NoteJetzt == 6 ? true : false;
                                if (!(from d in schueler.DefizitäreLeistungen where d.Fach == al.Fach where d.NoteHalbjahr == wl.NoteHalbjahr where d.NoteJetzt == wl.NoteJetzt select d).Any())
                                {
                                    schueler.DefizitäreLeistungen.Add(wl);
                                }
                            }                            
                        }

                        // ... oder aus Atlantis 

                        foreach (var al in (from t in defizitäreAtlantisLeistungen
                                            where t.NoteHalbjahr >= 5
                                            where t.SchlüsselExtern == schueler.IdAtlantis
                                            select t).ToList())
                        {
                            if (!(from d in schueler.DefizitäreLeistungen where d.Fach == al.Fach where d.NoteHalbjahr == al.NoteHalbjahr where d.NoteJetzt == al.NoteJetzt select d).Any())
                            {
                                al.NoteJetzt = 4;
                                al.NeueDefizitLeistung = al.NoteHalbjahr <= 4 && al.NoteJetzt >= 5 ? true : false;
                                al.NochmaligeVerschlechterungAuf6 = al.NoteHalbjahr == 5 && al.NoteJetzt == 6 ? true : false;
                                if (!(from d in schueler.DefizitäreLeistungen where d.Fach == al.Fach select d).Any())
                                {
                                    schueler.DefizitäreLeistungen.Add(al);
                                }
                            }
                        }

                        schueler.Dateien = new List<string>();

                        if (schueler.DefizitäreLeistungen.Count > 0)
                        {
                            this.Add(schueler);
                        }

                        reader.Close();
                        command.Dispose();
                    }
                }
            }
            Console.WriteLine(("Schüler mit Defiziten " + ".".PadRight(this.Count / 150, '.')).PadRight(48, '.') + (" " + this.Count).ToString().PadLeft(4), '.');
        }
    }
}