using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Windows.Forms;
using System.IO;
using Outlook = Microsoft.Office.Interop.Outlook;
using System.Diagnostics;

namespace ExportJournal
{
    public partial class Form1 : Form
    {
        internal const string MDB_PATH = @"C:\Users\a.ausweger\Documents\Büro\honorare";

        public Form1()
        {
            InitializeComponent();

        }

        private void cmdExport_Click(object sender, EventArgs e)
        {
            try
            {
                this.Enabled = false;

                var templateMdb = GetTemplateMdb();
                if (templateMdb == default(string)) return;

                var newMdb = GetNewMdbName(templateMdb);
                if (newMdb == default(string)) return;

                File.Copy(templateMdb, newMdb);

                var entries = GetJournalEntries();
                if (entries.Count == 0) return;

                var mdbWriter = new MdbWriter() { MdbName = newMdb };

                DeleteExistingJournal(mdbWriter);

                mdbWriter.CreateJournal();

                try
                {
                    mdbWriter.InsertRows(entries);
                }
                catch (Exception ex)
                {
                    MessageBox.Show(ex.Message);
                }

                // ...and start a viewer.
                Process.Start(newMdb);

                this.Close();
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message);
            }
            finally
            {
                this.Enabled = true;
            }
        }

        private void DeleteExistingJournal(MdbWriter mdbWriter)
        {
            var journalTableName = mdbWriter.GetTableNames().FirstOrDefault(x => string.Compare(x, "Journal", StringComparison.OrdinalIgnoreCase) == 0);
            if (journalTableName != null)
            {
                mdbWriter.DeleteJournal();
            }
        }

        private List<JournalEntry> GetJournalEntries()
        {
            List<JournalEntry> entries = new List<JournalEntry>();

            Outlook.Application o = new Outlook.Application();
            Outlook._NameSpace ns = (Outlook._NameSpace)o.GetNamespace("MAPI");
            Outlook.MAPIFolder f = ns.GetDefaultFolder(Outlook.OlDefaultFolders.olFolderJournal);

            foreach (var item in f.Items)
            {

                Outlook.JournalItem journal = item as Outlook.JournalItem;
                if (journal != null)
                {
                    entries.Add(new JournalEntry() { Betreff = journal.Subject, BeginntAm = journal.CreationTime, Dauer = journal.Duration, Text = journal.Body, Kategorien = journal.Categories});
                }
            }
            return entries;
        }

        private string GetNewMdbName(string templateName)
        {
            var incStr = templateName.Substring(templateName.Length - 8, 2);
            int inc;
            if (!int.TryParse(incStr, out inc))
            {
                MessageBox.Show(this, string.Format("Ungültiger Honorarnotenname: '{0}'!", Path.GetFileName(templateName)), Properties.Resources.MsgboxCaption);
                return default(string);
            }

            inc++;
            var newName =  templateName.Substring(0, templateName.Length - 8) + inc.ToString().PadLeft(2,'0') + templateName.Substring(templateName.Length - 6, 6);
            if (File.Exists(newName))
            {
                MessageBox.Show(this, string.Format("Honorar '{0}' existiert bereits!", Path.GetFileName(newName)), Properties.Resources.MsgboxCaption);
                return default(string);
            }

            return newName;
        }

        private string GetTemplateMdb()
        {
            OpenFileDialog ofd = new OpenFileDialog();
            ofd.Filter = "Access mdbs|*.mdb";
            ofd.InitialDirectory = MDB_PATH;
            var res = ofd.ShowDialog(this);
            if (res == System.Windows.Forms.DialogResult.Cancel)
            {
                return default(string);
            }
            else
            {
                return ofd.FileName;
            }

        }

        private void cmdCancel_Click(object sender, EventArgs e)
        {
            this.Close();
        }


    }
}
