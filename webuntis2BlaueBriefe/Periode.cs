// Published under the terms of GPLv3 Stefan Bäumer 2023

using System;

namespace webuntis2BlaueBriefe
{
    public class Periode
    {
        public Periode()
        {
        }

        public int IdUntis { get; internal set; }
        public string Name { get; internal set; }
        public string Langname { get; internal set; }
        public DateTime Von { get; internal set; }
        public DateTime Bis { get; internal set; }
    }
}