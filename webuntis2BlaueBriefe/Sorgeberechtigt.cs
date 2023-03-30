// Published under the terms of GPLv3 Stefan Bäumer 2023
namespace webuntis2BlaueBriefe
{
    public class Sorgeberechtigt
    {
        public string Vorname { get; private set; }
        public string Nachname { get; private set; }
        public string Strasse { get; private set; }
        public string Plz { get; private set; }
        public string Ort { get; private set; }

        public Sorgeberechtigt(string vorname, string nachname, string strasse, string plz, string ort)
        {
            Vorname = vorname;
            Nachname = nachname;
            Strasse = strasse;
            Plz = plz;
            Ort = ort;
        }

        public Sorgeberechtigt(string plz, string strasse, string ort)
        {
            Plz = plz;
            Strasse = strasse;
            Ort = ort;
        }
    }
}