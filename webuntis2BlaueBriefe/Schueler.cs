// Published under the terms of GPLv3 Stefan Bäumer 2023
using Microsoft.Office.Interop.Word;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;

namespace webuntis2BlaueBriefe
{
    internal class Schueler
    {
        public int IdAtlantis { get; internal set; }
        public string Art { get; internal set; }
        public string Nachname { get; internal set; }
        public string Anrede { get; internal set; }
        public string Vorname { get; internal set; }
        public string Plz { get; internal set; }
        public string Ort { get; internal set; }
        public string Strasse { get; internal set; }
        public string Klasse { get; set; }
        public string Jahrgang { get; set; }
        public DateTime Geburtsdatum { get; set; }
        public bool Volljaehrig { get; set; }
        public string GeschlechtMw { get; set; }
        public Leistungen DefizitäreLeistungen { get; set; }
        public string Typ { get; set; }
        public string MAnrede { get; internal set; }
        public string MSorgeberechtigtJn { get; internal set; }
        public string MOrt { get; internal set; }
        public string MPlz { get; internal set; }
        public string MStrasse { get; internal set; }
        public string MNachname { get; internal set; }
        public string MVorname { get; internal set; }
        public string VAnrede { get; internal set; }
        public string VSorgeberechtigtJn { get; internal set; }
        public string VOrt { get; internal set; }
        public string VPlz { get; internal set; }
        public string VStrasse { get; internal set; }
        public string VNachname { get; internal set; }
        public string VVorname { get; internal set; }
        public string Geschlecht { get; internal set; }
        public Sorgeberechtigte Sorgeberechtigte { get; internal set; }
        public string Klassenleitung { get; internal set; }
        public string KlassenleitungMw { get; internal set; }
        public string KlassenleitungMail { get; internal set; }
        public List<string> Dateien { get; set; }

        internal void RenderMitteilung(string art, string folder)
        {      
            // Für jede unterschiedliche Adresse

            var x = (from s in this.Sorgeberechtigte select s.Strasse).Distinct().Count();

            var sss = (from s in this.Sorgeberechtigte select s.Strasse).Distinct().ToList();

            var z = (from s in this.Sorgeberechtigte select s).FirstOrDefault();
            
            if (Volljaehrig)
            {
                sss = new List<string>() { Strasse };
            }

            if (x==0)
            {
                sss.Add(Strasse);
            }

            foreach (var strasse in sss)
            {
                var sorgeberechtigter = (from s in this.Sorgeberechtigte where s.Strasse == strasse select s).FirstOrDefault();

                var origFileName = "Blaue Briefe.docx";

                var fileName = folder + "\\" + Klasse + "-" + Nachname.Substring(0,2) + "-" + Vorname.Substring(0,2) + (x > 1 ? strasse : "") + (art == "G" ? "-Gefährdung.docx" : "-Mitteilung.docx");

                Dateien.Add(fileName);

                if (File.Exists(fileName))
                {
                    File.Delete(fileName);
                }

                System.IO.File.Copy(origFileName.ToString(), fileName.ToString());

                object oMissing = System.Reflection.Missing.Value;

                Application wordApp = new Application { Visible = true };
                Document doc = wordApp.Documents.Open(fileName, ReadOnly: false, Visible: true);
                doc.Activate();

                if (Volljaehrig)
                {
                    FindAndReplace(wordApp, doc, "<AnDieErziehungsberechtigtenVon>", "");
                }
                else
                {
                    FindAndReplace(wordApp, doc, "<AnDieErziehungsberechtigtenVon>", "An die Erziehungsberechtigten von");
                }

                FindAndReplace(wordApp, doc, "<anrede>", GetAnrede());
                FindAndReplace(wordApp, doc, "<anredeLerncoaching>", GetAnredeLerncoaching());
                FindAndReplace(wordApp, doc, "<vorname>", Vorname);
                FindAndReplace(wordApp, doc, "<nachname>", Nachname);
                FindAndReplace(wordApp, doc, "<dichSie>", Volljaehrig ? "Sie" : "Dich");

                if (!Volljaehrig)
                {
                    FindAndReplace(wordApp, doc, "<plz>", sorgeberechtigter == null ? Plz : sorgeberechtigter.Plz);
                    FindAndReplace(wordApp, doc, "<straße>", sorgeberechtigter == null ? Strasse : sorgeberechtigter.Strasse);
                    FindAndReplace(wordApp, doc, "<ort>", sorgeberechtigter == null ? Ort : sorgeberechtigter.Ort);
                }
                else
                {
                    FindAndReplace(wordApp, doc, "<plz>", "");
                    FindAndReplace(wordApp, doc, "<straße>", "!!! Kein Briefversand bei Volljährigen !!!");
                    FindAndReplace(wordApp, doc, "<ort>", "");
                }
                FindAndReplace(wordApp, doc, "<klasse>", Klasse);
                FindAndReplace(wordApp, doc, "<heute>", DateTime.Now.ToShortDateString());
                FindAndReplace(wordApp, doc, "<betreff>", art == "M" ? "Mitteilung über den Leistungsstand" : "Gefährdung der Versetzung");
                FindAndReplace(wordApp, doc, "<absatz1>", GetAbsatz1(art));
                FindAndReplace(wordApp, doc, "<fächer>", RenderFächer(art));
                FindAndReplace(wordApp, doc, "<absatz2>", GetAbsatz2(art));
                FindAndReplace(wordApp, doc, "<absatz3>", GetAbsatz3());
                FindAndReplace(wordApp, doc, "<klassenleitung>", Klassenleitung);
                FindAndReplace(wordApp, doc, "<klassenlehrerIn>", KlassenleitungMw == "Herr" ? "Klassenlehrer" : "Klassenlehrerin");
                FindAndReplace(wordApp, doc, "<hinweis>", GetHinweis());
                FindAndReplace(wordApp, doc, "<footer>", "");

                doc.ExportAsFixedFormat(fileName.Replace(".docx","") + ".pdf", WdExportFormat.wdExportFormatPDF, false, WdExportOptimizeFor.wdExportOptimizeForOnScreen,
                    WdExportRange.wdExportAllDocument, 1, 1, WdExportItem.wdExportDocumentContent, true, true,
                    WdExportCreateBookmarks.wdExportCreateHeadingBookmarks, true, true, false, ref oMissing);
                doc.Save();
                doc.Close();
                System.Runtime.InteropServices.Marshal.ReleaseComObject(doc);
                doc = null;
                GC.Collect();
                wordApp.Quit();
            }            
        }

        private object GetAnredeLerncoaching()
        {
            string x = "";

            x += "Liebe" + (GeschlechtMw == "M" ? "r " : " ") + Vorname;

            if (!Volljaehrig)
            {
                x += ",\r\nliebe Erziehungsberechtigte";
            }

            x += "!";
            return x;
        }

        private string RenderGefährdungNeu()
        {
            string x = "Neu hinzukommende Gefährdung: ";

            foreach (var item in (from f in DefizitäreLeistungen where f.NeueDefizitLeistung select f).ToList())
            {
                //x += " " + item.KürzelUntis + "(" + (from g in Global.Noten where item.NoteJetzt == g.Stufe select g.Klartext).FirstOrDefault() + "),";
            }
            return x.TrimEnd(',');
        }

        private object GetAbsatz3()
        {
            return "Wir laden Sie zu einem Beratungsgespräch ein. Stimmen Sie bitte den Gesprächster- min mit " + (KlassenleitungMw == "Herr" ? "dem Klassenlehrer" : "der Klassenlehrerin") + " " + Klassenleitung + " (" + KlassenleitungMail + ") ab.";
        }

        private object GetAbsatz2(string art)
        {
            if (art == "M")
            {
                return "abweichend von " + ((from f in DefizitäreLeistungen where f.NeueDefizitLeistung || f.NochmaligeVerschlechterungAuf6 select f).Count() > 1 ? "den" : "der") + " im letzten Zeugnis erteilten Note" + ((from f in DefizitäreLeistungen where f.NeueDefizitLeistung || f.NochmaligeVerschlechterungAuf6 select f).Count() > 1 ? "n" : "") + " verschlechtert hat. Stellt sich eine weitere nicht ausreichende Leistung ein, ist die Versetzung gefährdet.";
            }
            if (art == "G")
            {
                return "abweichend von " + ((from f in DefizitäreLeistungen where f.NeueDefizitLeistung || f.NochmaligeVerschlechterungAuf6 select f).Count() > 1 ? "den" : "der") + " im letzten Zeugnis erteilten Note" + ((from f in DefizitäreLeistungen where f.NeueDefizitLeistung || f.NochmaligeVerschlechterungAuf6 select f).Count() > 1 ? "n" : "") + " verschlechtert " + ((from f in DefizitäreLeistungen where f.NeueDefizitLeistung || f.NochmaligeVerschlechterungAuf6 select f).Count() > 1 ? "haben" : "hat") + ".";
            }
            else
            {
                return "abweichend von der im letzten Zeugnis erteilten Note nur noch ungenügend ist.";
            }
        }

        private object GetAbsatz1(string art)
        {
            if (!Volljaehrig)
            {
                if (Geschlecht.ToLower() == "m")
                {
                    return "Sie werden darüber unterrichtet, dass sich die Leistung" + ((from d in DefizitäreLeistungen where d.NeueDefizitLeistung || d.NochmaligeVerschlechterungAuf6 select d).Count() > 1 ? "en" : "") + " Ihres Sohnes " + Vorname + ", Klasse " + Klasse + ", in " + ((from d in DefizitäreLeistungen where d.NeueDefizitLeistung || d.NochmaligeVerschlechterungAuf6 select d).Count() > 1 ? "den Fächern" : "dem Fach");
                }
                else
                {
                    return "Sie werden darüber unterrichtet, dass sich die Leistung" + ((from d in DefizitäreLeistungen where d.NeueDefizitLeistung || d.NochmaligeVerschlechterungAuf6 select d).Count() > 1 ? "en" : "") + " Ihrer Tochter " + Vorname + ", Klasse " + Klasse + ", in " + ((from d in DefizitäreLeistungen where d.NeueDefizitLeistung || d.NochmaligeVerschlechterungAuf6 select d).Count() > 1 ? "den Fächern" : "dem Fach");
                }
            }
            else
            {
                return "Sie werden darüber unterrichtet, dass sich Ihre Leistung" + ((from d in DefizitäreLeistungen where d.NeueDefizitLeistung || d.NochmaligeVerschlechterungAuf6 select d).Count() > 1 ? "en" : "") + " in " + ((from f in DefizitäreLeistungen where f.NeueDefizitLeistung || f.NochmaligeVerschlechterungAuf6 select f).Count() > 1 ? "den Fächern" : "dem Fach");
            }
        }

        private object GetIhreTochterIhrSohn()
        {
            if (!Volljaehrig)
            {
                if (Geschlecht.ToLower() == "m")
                {
                    return "die Leistung Ihres Sohnes " + Vorname + ", Klasse " + Klasse + ", ";
                }
                else
                {
                    return "die Leistung Ihrer Tochter " + Vorname + ", Klasse " + Klasse + ", ";
                }
            }
            else
            {
                return @"Ihre 
Leistung";
            }
        }

        private object GetHinweis()
        {
            if (!Volljaehrig)
            {
                if (Geschlecht.ToLower() == "m")
                {
                    return "Ihr Sohn die Klasse zurzeit wiederholt,";
                }
                else
                {
                    return "Ihre Tochter die Klasse zurzeit wiederholt,";
                }
            }
            else
            {
                return "Sie die Klasse zurzeit wiederholen,";
            }
        }

        private object GetAnrede()
        {
            if (Volljaehrig)
            {
                if (Geschlecht.ToLower() == "m")
                {
                    return "Sehr geehrter Herr " + Vorname + " " + Nachname + ",";
                }
                else
                {
                    return "Sehr geehrte Frau " + Vorname + " " + Nachname + ",";
                }
            }
            else
            {
                return "Sehr geehrte Erziehungsberechtigte,";
            }
        }

        private string RenderFächer(string art)
        {
            string x = "";

            if (art == "V")
            {
                foreach (var dl in (from d in DefizitäreLeistungen where d.NoteHalbjahr == 5 where d.NoteJetzt == 6 select d).ToList())
                {
                    x += " " + dl.BezeichnungImZeugnis.TrimEnd(Environment.NewLine.ToCharArray()) + " (" + NoteKlartext(dl.NoteJetzt) + ")\r\n";
                }
            }
            else
            {
                foreach (var dl in (from d in DefizitäreLeistungen where d.NeueDefizitLeistung || d.NochmaligeVerschlechterungAuf6 select d).ToList())
                {
                    x += " " + dl.BezeichnungImZeugnis.TrimEnd(Environment.NewLine.ToCharArray()) + " (" + NoteKlartext(dl.NoteJetzt) + ")\r\n";
                }
            }
            return x.Replace(" **)", "");
        }

        private string NoteKlartext(int noteJetzt)
        {
            if (noteJetzt == 6)
            {
                return "ungenügend";
            }
            if (noteJetzt == 5)
            {
                return "mangelhaft";
            }
            if (noteJetzt == 4)
            {
                return "ausreichend";
            }
            if (noteJetzt == 3)
            {
                return "befriedigend";
            }
            if (noteJetzt == 2)
            {
                return "gut";
            }
            if (noteJetzt == 1)
            {
                return "sehr gut";
            }
            return "fehler";
        }

        private static void FindAndReplace(Application app, Document doc, object findText, object replaceWithText)
        {
            //options
            object matchCase = false;
            object matchWholeWord = true;
            object matchWildCards = false;
            object matchSoundsLike = false;
            object matchAllWordForms = false;
            object forward = true;
            object format = false;
            object matchKashida = false;
            object matchDiacritics = false;
            object matchAlefHamza = false;
            object matchControl = false;
            object read_only = false;
            object visible = true;
            object replace = 2;
            object wrap = 1;
            //execute find and replace
            try
            {
                // Der neue Text darf nur 255 Zeichen lang sein.

                if (replaceWithText.ToString().Length < 255)
                {
                    app.Selection.Find.Execute(ref findText, ref matchCase, ref matchWholeWord,
                    ref matchWildCards, ref matchSoundsLike, ref matchAllWordForms, ref forward, ref wrap, ref format, ref replaceWithText, ref replace,
                    ref matchKashida, ref matchDiacritics, ref matchAlefHamza, ref matchControl);
                }
                else
                {
                    object empty = "";
                    Bookmark bm = doc.Bookmarks["faecher"];
                    Range range = bm.Range;
                    range.Text = replaceWithText.ToString();
                    doc.Bookmarks.Add("faecher", range);
                    app.Selection.Find.Execute(ref findText, ref matchCase, ref matchWholeWord,
                    ref matchWildCards, ref matchSoundsLike, ref matchAllWordForms, ref forward, ref wrap, ref format, ref empty, ref replace,
                    ref matchKashida, ref matchDiacritics, ref matchAlefHamza, ref matchControl);
                }
            }
            catch (Exception ex)
            {
                Console.WriteLine(ex);
                Console.ReadKey();
            }
        }
    }
}