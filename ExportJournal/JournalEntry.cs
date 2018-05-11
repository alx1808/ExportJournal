using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace ExportJournal
{
    internal class JournalEntry
    {
        public string Betreff { get; set; }
        public DateTime BeginntAm { get; set; }
        public int Dauer { get; set; }
        private string _Text = "";
        public string Text
        {
            get { return _Text; }
            set { _Text = value; }
        }
        
        public string Kategorien { get; set; }
    }
}
