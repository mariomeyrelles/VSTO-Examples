using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Windows.Forms;
using Google.GData.Spreadsheets;

namespace GoogleDocsIntegration.Forms
{
    public partial class LoginForm : Form
    {
        public LoginForm()
        {
            InitializeComponent();
        }

        private void label3_Click(object sender, EventArgs e)
        {

        }

        private void button1_Click(object sender, EventArgs e)
        {
            App.Service.setUserCredentials(this.txtUser.Text, this.txtPass.Text);

            // Instantiate a SpreadsheetQuery object to retrieve spreadsheets.
            SpreadsheetQuery query = new SpreadsheetQuery();

            // Make a request to the API and get all spreadsheets.
            SpreadsheetFeed feed = App.Service.Query(query);

            if (feed.Entries.Count == 0)
            {
                MessageBox.Show("There are no spreadsheets created by this user.","Google Spreadsheets API", MessageBoxButtons.OK, MessageBoxIcon.Error);
                return;
            }


            App.CurrentUser = this.txtUser.Text;
            App.CurrentPassword = this.txtPass.Text;


            var wksListForm = new WorksheetListForm();
            
            wksListForm.Show();
            this.Close();
            



        }

        private void LoginForm_Load(object sender, EventArgs e)
        {
            this.txtUser.Text = "mariomeyrelles@gmail.com";
            this.txtPass.Text = "aquarius.12";
        }
    }
}
