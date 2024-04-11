// Published under the terms of GPLv3 Stefan Bäumer 2023

using System;
using System.Collections.Generic;
using System.Data.SqlClient;
using System.IO;
using System.Linq;

namespace webuntis2BlaueBriefe
{
    internal class Global
    {
        public static string ConnectionStringAtlantis = @"Dsn=Atlantis17u;uid=DBA";
        //public static string ConnectionStringUntis = @"Data Source=SQL01\UNTIS;Initial Catalog=master;Integrated Security=True";
        public const string ConnectionStringUntis = @"Data Source=SQL01\UNTIS;Initial Catalog=master;Integrated Security=True";
        public static string BlaueBriefe
        { get { return "Mahnung gem. §50 (4) SchulG (Blauer Brief)"; } }

        public static string AktSjUntis
        {
            get
            {
                int sj = (DateTime.Now.Month >= 8 ? DateTime.Now.Year : DateTime.Now.Year - 1);
                return sj.ToString() + (sj + 1);
            }
        }

        public static string AktSjAtlantis
        {
            get
            {
                List<string> aktSj = new List<string>
                {
                    (DateTime.Now.Month >= 8 ? DateTime.Now.Year : DateTime.Now.Year - 1).ToString(),
                    (DateTime.Now.Month >= 8 ? DateTime.Now.Year + 1 - 2000 : DateTime.Now.Year - 2000).ToString()
                };
                return aktSj[0] + "/" + aktSj[1];
            }
        }
        
        public static string SafeGetString(SqlDataReader reader, int colIndex)
        {
            if (!reader.IsDBNull(colIndex))
                return reader.GetString(colIndex);
            return string.Empty;
        }

        internal static string List2String(List<string> interessierendeKlassen, char delimiter)
        {
            var s = "";
            foreach (var item in interessierendeKlassen)
            {
                s += item + delimiter;
            }
            return s.TrimEnd(delimiter);
        }

        internal static void Write(string folder, string v)
        {
            File.AppendAllText(folder + "\\Protokoll.txt", v);            
        }

        internal static void WriteLine(string folder, string v)
        {
            File.AppendAllText(folder + "\\Protokoll.txt", v + Environment.NewLine);
        }
    }
}