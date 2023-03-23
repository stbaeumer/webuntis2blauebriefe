// Published under the terms of GPLv3 Stefan Bäumer 2023

using System;
using System.Collections.Generic;
using System.Data.SqlClient;
using System.Linq;

namespace webuntis2BlaueBriefe
{
    public class Klasses : List<Klasse>
    {
        public Lehrers Lehrers { get; set; }

        public Klasses(Lehrers lehrers, Periodes periodes, Leistungen dlw)
        {
            Lehrers = lehrers;

            using (SqlConnection sqlConnection = new SqlConnection(Global.ConnectionStringUntis))
            {
                try
                {
                    string queryString = @"SELECT 
Class.CLASS_ID, 
Class.Name, 
Class.TeacherIds, 
Class.Longname, 
Teacher.Name,
Class.ClassLevel, 
Class.PERIODS_TABLE_ID,
Department.Name,
Class.TimeRequest,
Class.ROOM_ID,
Class.Text
FROM (Class LEFT JOIN Department ON Class.DEPARTMENT_ID = Department.DEPARTMENT_ID) LEFT JOIN Teacher ON Class.TEACHER_ID = Teacher.TEACHER_ID
WHERE (((Class.SCHOOL_ID)=177659) AND ((Class.TERM_ID)=" + periodes.Count + ") AND ((Class.Deleted)='false') AND ((Class.TERM_ID)=" + periodes.Count + ") AND ((Class.SCHOOLYEAR_ID)=" + Global.AktSjUntis + ") AND ((Department.SCHOOL_ID)=177659) AND ((Department.SCHOOLYEAR_ID)=" + Global.AktSjUntis + ") AND ((Teacher.SCHOOL_ID)=177659) AND ((Teacher.SCHOOLYEAR_ID)=" + Global.AktSjUntis + ") AND ((Teacher.TERM_ID)=" + periodes.Count + "))ORDER BY Class.Name ASC; ";

                    SqlCommand odbcCommand = new SqlCommand(queryString, sqlConnection);
                    sqlConnection.Open();
                    SqlDataReader sqlDataReader = odbcCommand.ExecuteReader();

                    while (sqlDataReader.Read())
                    {
                        List<Lehrer> klassenleitungen = new List<Lehrer>();

                        foreach (var item in (Global.SafeGetString(sqlDataReader, 2)).Split(','))
                        {
                            klassenleitungen.Add((from l in lehrers
                                                  where l.IdUntis.ToString() == item
                                                  select l).FirstOrDefault());
                        }

                        var klasseName = Global.SafeGetString(sqlDataReader, 1);

                        Klasse klasse = new Klasse()
                        {
                            IdUntis = sqlDataReader.GetInt32(0),
                            NameUntis = klasseName,
                            Klassenleitungen = klassenleitungen,
                            Jahrgang = Global.SafeGetString(sqlDataReader, 5),
                            Bereichsleitung = Global.SafeGetString(sqlDataReader, 7),
                            Beschreibung = Global.SafeGetString(sqlDataReader, 3),                            
                            Url = "https://www.berufskolleg-borken.de/bildungsgange/" + Global.SafeGetString(sqlDataReader, 10)
                        };

                        if ((from d in dlw where d.Klasse == klasse.NameUntis select d).Any())
                        {
                            this.Add(klasse);
                        }
                    };

                    Console.WriteLine(("Klassen " + ".".PadRight(this.Count / 150, '.')).PadRight(48, '.') + (" " + this.Count).ToString().PadLeft(4), '.');

                    sqlDataReader.Close();
                }
                catch (Exception ex)
                {
                    Console.WriteLine(ex.ToString());
                    throw new Exception(ex.ToString());
                }
                finally
                {
                    sqlConnection.Close();
                }
            }
        }
    }
}