using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using GoogleDocsIntegration.Forms;
using Microsoft.Office.Tools.Ribbon;

namespace GoogleDocsIntegration
{
    public partial class Ribbon1
    {
        private void Ribbon1_Load(object sender, RibbonUIEventArgs e)
        {

        }

       
        private void button1_Click(object sender, RibbonControlEventArgs e)
        {
            var wksForm = new WorksheetListForm();
            wksForm.Show();
        }
    }
}
