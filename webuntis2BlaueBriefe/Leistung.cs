// Published under the terms of GPLv3 Stefan Bäumer 2023

using System;
using System.Linq;

namespace webuntis2BlaueBriefe
{
    public class Leistung
    {
        public DateTime Datum { get; internal set; }
        public string Name { get; internal set; }
        public string Klasse { get; internal set; }
        public string Fach { get; internal set; }
        public string Prüfungsart { get; internal set; }
        public int NoteJetzt { get; internal set; }
        public string Bemerkung { get; internal set; }
        public string Benutzer { get; internal set; }
        public int SchlüsselExtern { get; internal set; }
        public int LeistungId { get; internal set; }
        public bool ReligionAbgewählt { get; internal set; }
        public int NoteHalbjahr { get; internal set; }
        public string Schuljahr { get; internal set; }
        public string Gliederung { get; internal set; }
        public bool HatBemerkung { get; internal set; }
        public int Jahrgang { get; internal set; }
        public string Nachname { get; internal set; }
        public string Vorname { get; internal set; }
        public int LehrkraftAtlantisId { get; internal set; }
        public string Bereich { get; internal set; }
        public DateTime Geburtsdatum { get; internal set; }
        public bool Volljährig { get; internal set; }
        public string Gesamtpunkte { get; internal set; }
        public string Tendenz { get; internal set; }
        public string EinheitNP { get; internal set; }
        public string HzJz { get; internal set; }
        public string Anlage { get; internal set; }
        public string Zeugnisart { get; internal set; }
        public string BezeichnungImZeugnis { get; internal set; }
        public DateTime Konferenzdatum { get; internal set; }
        public bool NeueDefizitLeistung { get; internal set; }
        public bool NochmaligeVerschlechterungAuf6 { get; internal set; }
        public string KlasseAtlantis { get; internal set; }
    }
}