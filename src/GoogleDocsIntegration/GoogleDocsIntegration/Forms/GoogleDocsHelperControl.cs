using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Diagnostics;
using System.Drawing;
using System.Data;
using System.Linq;
using System.Text;
using System.Windows.Forms;
using Google.GData.Client;
using Google.GData.Spreadsheets;

namespace GoogleDocsIntegration.Forms
{
    public partial class GoogleDocsHelperControl : UserControl
    {
        public GoogleDocsHelperControl()
        {
            InitializeComponent();
        }

      

        private void GoogleDocsHelperControl_Load(object sender, EventArgs e)
        {
            var currentWks = App.CurrentWorksheet.Title.Text;
            this.lblActiveWks.Text = currentWks;
        }

        private void btnRetrieveData_Click(object sender, EventArgs e)
        {
            var wks = App.CurrentWorksheet;


            // Define the URL to request the list feed of the worksheet.
            var listFeedLink = wks.Links.FindService(GDataSpreadsheetsNameTable.ListRel, null);

            // Fetch the list feed of the worksheet.
            var listQuery = new ListQuery(listFeedLink.HRef.ToString());
            var listFeed = App.Service.Query(listQuery);

            var dataTable = new DataTable();


            foreach (var entry in listFeed.Entries[0].ExtensionElements)
            {
                var columnName = entry.XmlName;
                dataTable.Columns.Add(columnName);
            }

            // Iterate through each row, printing its header and cell values.
            foreach (ListEntry row in listFeed.Entries)
            {
                var dataRow = dataTable.NewRow();

                // Iterate over the columns, and print each cell value
                foreach (ListEntry.Custom element in row.Elements) //elements = columns!
                {
                    Debug.WriteLine(element.LocalName + "; " + element.Value);
                    dataRow[element.LocalName] = element.Value;
                   
                }

                dataTable.Rows.Add(dataRow);
            }

            
            
            Globals.Sheet1.tblGoogleData.SetDataBinding(dataTable,"");
        }
    }
}
