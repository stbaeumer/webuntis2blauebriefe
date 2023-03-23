// Published under the terms of GPLv3 Stefan Bäumer 2023
using System;
using System.Collections.Generic;
using System.Data.SqlClient;

namespace webuntis2BlaueBriefe
{
    public class Lehrers : List<Lehrer>
    {
        public Lehrers()
        {
        }

        public Lehrers(Periodes periodes)
        {
            using (SqlConnection sqlConnection = new SqlConnection(Global.ConnectionStringUntis))
            {
                try
                {
                    string queryString = @"SELECT DISTINCT 
Teacher.Teacher_ID, 
Teacher.Name, 
Teacher.Longname, 
Teacher.FirstName,
Teacher.Email,
Teacher.Flags,
Teacher.Title,
Teacher.ROOM_ID,
Teacher.Text2,
Teacher.Text3,
Teacher.PlannedWeek
FROM Teacher 
WHERE (((SCHOOLYEAR_ID)= " + Global.AktSjUntis + ") AND  ((TERM_ID)=" + periodes.Count + ") AND ((Teacher.SCHOOL_ID)=177659) AND (((Teacher.Deleted)='false'))) ORDER BY Teacher.Name;";

                    SqlCommand odbcCommand = new SqlCommand(queryString, sqlConnection);
                    sqlConnection.Open();
                    SqlDataReader sqlDataReader = odbcCommand.ExecuteReader();

                    while (sqlDataReader.Read())
                    {
                        Lehrer lehrer = new Lehrer()
                        {
                            IdUntis = sqlDataReader.GetInt32(0),
                            Kürzel = Global.SafeGetString(sqlDataReader, 1),
                            Nachname = Global.SafeGetString(sqlDataReader, 2),
                            Vorname = Global.SafeGetString(sqlDataReader, 3),
                            Mail = Global.SafeGetString(sqlDataReader, 4),
                            Anrede = Global.SafeGetString(sqlDataReader, 5) == "n" ? "Herr" : Global.SafeGetString(sqlDataReader, 5) == "W" ? "Frau" : "",
                            Titel = Global.SafeGetString(sqlDataReader, 6),
                            Raum = "",
                            Funktion = Global.SafeGetString(sqlDataReader, 8),
                            Dienstgrad = Global.SafeGetString(sqlDataReader, 9)
                        };

                        if (!lehrer.Mail.EndsWith("@berufskolleg-borken.de") && lehrer.Kürzel != "LAT" && lehrer.Kürzel != "?")
                            Console.WriteLine("Untis2Exchange Fehlermeldung: Der Lehrer " + lehrer.Kürzel + " hat keine Mail-Adresse in Untis. Bitte in Untis eintragen.");
                        if (lehrer.Anrede == "" && lehrer.Kürzel != "LAT" && lehrer.Kürzel != "?")
                            Console.WriteLine("Untis2Exchange Fehlermeldung: Der Lehrer " + lehrer.Kürzel + " hat kein Geschlecht in Untis. Bitte in Untis eintragen.");

                        this.Add(lehrer);
                    };

                    Console.WriteLine(("Lehrer " + ".".PadRight(this.Count / 150, '.')).PadRight(48, '.') + (" " + this.Count).ToString().PadLeft(4), '.');

                    sqlDataReader.Close();
                }
                catch (Exception ex)
                {
                    Console.WriteLine(ex.ToString());
                }
                finally
                {
                    sqlConnection.Close();
                }
            }
        }
    }
}