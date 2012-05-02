using System;
using System.Collections.Generic;
using System.Data;
using System.Linq;
using System.Text;
using System.Windows.Forms;
using System.Xml.Linq;
using Microsoft.Office.Tools.Excel;
using Microsoft.VisualStudio.Tools.Applications.Runtime;
using Excel = Microsoft.Office.Interop.Excel;
using Office = Microsoft.Office.Core;

namespace Intro
{
    public partial class Sheet1
    {
        private void Sheet1_Startup(object sender, System.EventArgs e)
        {
        }

        private void Sheet1_Shutdown(object sender, System.EventArgs e)
        {
        }

        #region VSTO Designer generated code

        /// <summary>
        /// Required method for Designer support - do not modify
        /// the contents of this method with the code editor.
        /// </summary>
        private void InternalStartup()
        {
            this.button1.Click += new System.EventHandler(this.button1_Click);
            this.Startup += new System.EventHandler(this.Sheet1_Startup);
            this.Shutdown += new System.EventHandler(this.Sheet1_Shutdown);

        }

        #endregion

        
        private void button1_Click(object sender, EventArgs e)
        {
            Intro1();
        }

        public void Intro1()
        {
            var initialRange = Range["C5"];
            var checks = new List<Check>();
            var row = 0;

            while (initialRange.Offset[row, 0].Value != null)
            {
                var item = new Check();
                item.CheckNumber = (long) initialRange.Offset[row, 0].Value;
                item.CustomerName = initialRange.Offset[row, 1].Value;
                item.Amount= initialRange.Offset[row, 2].Value;
                item.DueDate = initialRange.Offset[row, 3].Value;

                checks.Add(item);

                row++;
            }

            initialRange = Range["I9"];
            row = 0;


            while (initialRange.Offset[row, 0].Value != null)
            {
                var date = initialRange.Offset[row, 0].Value;
                var month = date.Month;
                double monthlyAmount = 0;

                foreach (var cheque in checks)
                {
                    if (cheque.DueDate.Month == month)
                        monthlyAmount += cheque.Amount;
                }

                initialRange.Offset[row, 1].Value = monthlyAmount;

                row++;
            }

        }
    }
}
