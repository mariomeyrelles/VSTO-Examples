using System;
using System.Collections;
using System.Collections.Generic;
using System.Diagnostics;
using System.Linq;
using System.Text;
using System.Windows.Forms;
using Microsoft.Office.Interop.Excel;
using ListObject = Microsoft.Office.Tools.Excel.ListObject;
using Office = Microsoft.Office.Core;

namespace ListObjectReadWrite
{
    public class ReadWriteTests
    {
        public List<SalesOrderDetail> Read()
        {

            var salesDetails = new List<SalesOrderDetail>();

            Stopwatch watch = new Stopwatch();
            watch.Start();

            for (int i = 1; i <= Globals.Sheet1.tblSalesOrderDetails.ListRows.Count; i++)
            {
                var salesDetail = new SalesOrderDetail();

                salesDetail.SalesOrderDetailID = (int) Globals.Sheet1.tblSalesOrderDetails.ListRows[i].Range[1, 1].Value2;

                salesDetails.Add(salesDetail);
            }

            watch.Stop();

            MessageBox.Show("Elapsed time: " + watch.Elapsed.TotalSeconds);

            return salesDetails;
        }

        public List<SalesOrderDetail> ReadFaster()
        {
            Stopwatch watch = new Stopwatch();
            watch.Start();

            var salesDetails = new List<SalesOrderDetail>();
            object[,] rawData = Globals.Sheet1.tblSalesOrderDetails.Range.Value2;

            for (var row = 2; row <= rawData.GetLength(0); row++)
            {
                var salesDetail = new SalesOrderDetail();

                salesDetail.SalesOrderDetailID = Convert.ToInt32(rawData[row, 1]);

                salesDetails.Add(salesDetail);
            }

            watch.Stop();

            MessageBox.Show("Elapsed time: " + watch.Elapsed.TotalSeconds);

            return salesDetails;
        }



    }
}