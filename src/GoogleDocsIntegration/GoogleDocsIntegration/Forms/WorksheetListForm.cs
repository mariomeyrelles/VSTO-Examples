using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Windows.Forms;
using Google.GData.Client;
using Google.GData.Spreadsheets;

namespace GoogleDocsIntegration.Forms
{
    public partial class WorksheetListForm : Form
    {
        public WorksheetListForm()
        {
            InitializeComponent();
        }

       

        private void WorksheetListForm_Load(object sender, EventArgs e)
        {

            //Check if user/pass are null. If the are null, show the login form.

            if (string.IsNullOrEmpty(App.CurrentUser) ||
                string.IsNullOrEmpty(App.CurrentPassword))
            {
                var loginForm = new LoginForm();
                loginForm.Show();
                this.Close();
                return;
            }


            // Instantiate a SpreadsheetQuery object to retrieve spreadsheets.
            SpreadsheetQuery query = new SpreadsheetQuery();

            // Make a request to the API and get all spreadsheets.
            SpreadsheetFeed feed = App.Service.Query(query);

            if (feed.Entries.Count == 0)
            {
                MessageBox.Show("There are no spreadsheets created by this user.", "Google Spreadsheets API", MessageBoxButtons.OK, MessageBoxIcon.Error);
                this.Close();
            }

            //iterate over each worksheet found and fill the list view with title, author and link
            foreach (var entry in feed.Entries)
            {
                ListViewItem item = new ListViewItem(new string[]{entry.Title.Text,entry.Authors[0].Name, entry.AlternateUri.Content});
                this.listView1.Items.Add(item);
            }

        }

        private void listView1_SelectedIndexChanged(object sender, EventArgs e)
        {
            listView2.Clear();
            if (listView1.SelectedItems.Count == 0)
            {
                return;
            }
            
            var documentTitle = this.listView1.SelectedItems[0].SubItems[0].Text;

            var query = new SpreadsheetQuery();
            query.Title = documentTitle;

            //looking for a exact worksheet of the given selected title. Returns 1 match.
            var selectedSpreadsheet = (SpreadsheetEntry)App.Service.Query(query).Entries[0];

            //retrieve all worksheets available in the selected spreadsheet.
            var allWorksheetsFeed = selectedSpreadsheet.Worksheets;

            foreach (var worksheet in allWorksheetsFeed.Entries)
            {
                var listItem = new ListViewItem(worksheet.Title.Text){Tag = worksheet};
                this.listView2.Items.Add(listItem);
            }

        }

        private void button1_Click(object sender, EventArgs e)
        {
            if (listView2.SelectedItems.Count == 0)
            {
                return;
            }

            var currentWks = listView2.SelectedItems[0].Tag;
            App.CurrentWorksheet = (WorksheetEntry) currentWks;
            App.SetSidePanel();


            this.Close();
        }
    }
}
