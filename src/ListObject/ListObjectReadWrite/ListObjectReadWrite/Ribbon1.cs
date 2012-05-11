using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using Microsoft.Office.Tools.Ribbon;

namespace ListObjectReadWrite
{
    public partial class Ribbon1
    {
        private ReadWriteTests _tests;

        private void Ribbon1_Load(object sender, RibbonUIEventArgs e)
        {
            _tests = new ReadWriteTests();
        }

        private void button1_Click(object sender, RibbonControlEventArgs e)
        {
            _tests.FailedReadMethod();
        }

        private void button2_Click(object sender, RibbonControlEventArgs e)
        {
            _tests.ReadFaster();
        }

        private void button3_Click(object sender, RibbonControlEventArgs e)
        {
            _tests.Write();
        }

        private void button4_Click(object sender, RibbonControlEventArgs e)
        {
            _tests.WriteFaster();
        }
    }
}
