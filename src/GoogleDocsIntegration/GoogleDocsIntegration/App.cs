using System;
using Google.GData.Client;
using Google.GData.Spreadsheets;
using GoogleDocsIntegration.Forms;

namespace GoogleDocsIntegration
{
    public class App
    {
        private static SpreadsheetsService _service;
        public static SpreadsheetsService Service
        {
            get { return _service ?? (_service = new SpreadsheetsService("Spreadsheet-GData-Sample-App")); }
        }

        public static string CurrentUser { get; set; }

        public static string CurrentPassword { get; set; }

        public static WorksheetEntry CurrentWorksheet { get; set; }


        public static void SetSidePanel()
        {
            Globals.ThisWorkbook.ActionsPane.Controls.Add(new GoogleDocsHelperControl());
        }
    }
}