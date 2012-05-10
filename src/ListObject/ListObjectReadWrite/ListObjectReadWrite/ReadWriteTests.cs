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
        private List<SalesOrderDetail> _salesDetails = new List<SalesOrderDetail>(122000);

        public void Read()
        {
            _salesDetails.Clear();
            _salesDetails = new List<SalesOrderDetail>(122000);
            
            var watch = new Stopwatch();
            watch.Start();

            for (var i = 1; i <= Globals.Sheet1.tblSalesOrderDetails.ListRows.Count; i++)
            {
                var salesDetail = new SalesOrderDetail();

                salesDetail.SalesOrderID = (int)Globals.Sheet1.tblSalesOrderDetails.ListRows[i].Range[1, 1].Value2;
                salesDetail.SalesOrderDetailID = (int)Globals.Sheet1.tblSalesOrderDetails.ListRows[i].Range[1, 2].Value2;
                salesDetail.CarrierTrackingNumber = Globals.Sheet1.tblSalesOrderDetails.ListRows[i].Range[1, 3].Value2;
                salesDetail.OrderQty = (int)Globals.Sheet1.tblSalesOrderDetails.ListRows[i].Range[1, 4].Value2;
                salesDetail.ProductID = (int)Globals.Sheet1.tblSalesOrderDetails.ListRows[i].Range[1, 5].Value2;
                salesDetail.SpecialOfferID = (int)Globals.Sheet1.tblSalesOrderDetails.ListRows[i].Range[1, 6].Value2;
                salesDetail.UnitPrice = (double)Globals.Sheet1.tblSalesOrderDetails.ListRows[i].Range[1, 7].Value2;
                salesDetail.UnitPriceDiscount = (double)Globals.Sheet1.tblSalesOrderDetails.ListRows[i].Range[1, 8].Value2;
                salesDetail.LineTotal = (double)Globals.Sheet1.tblSalesOrderDetails.ListRows[i].Range[1, 9].Value2;
                salesDetail.ModifiedDate = DateTime.FromOADate(Globals.Sheet1.tblSalesOrderDetails.ListRows[i].Range[1, 10].Value2);

                _salesDetails.Add(salesDetail);
            }

            watch.Stop();
            //about 220 secs
            MessageBox.Show("Elapsed time in sec: " + watch.Elapsed.TotalSeconds);

        }

        public void ReadFaster()
        {
            _salesDetails.Clear();
            _salesDetails = new List<SalesOrderDetail>(122000);

            var sb = new StringBuilder();

            var watch = new Stopwatch();
            watch.Start();

            
            Globals.ThisWorkbook.ThisApplication.EnableEvents = false;

            object[,] rawData = Globals.Sheet1.tblSalesOrderDetails.Range.Value2;

            watch.Stop();

            //about 13 secs
            sb.AppendLine("Elapsed Time (getting data from Excel) in sec: " + watch.Elapsed.TotalSeconds);

            watch.Restart();
            
            for (var row = 2; row <= rawData.GetLength(0); row++)
            {
                var salesDetail = new SalesOrderDetail();

                salesDetail.SalesOrderID = Convert.ToInt32(rawData[row, 1]);
                salesDetail.SalesOrderDetailID =  Convert.ToInt32(rawData[row, 2]);
                salesDetail.CarrierTrackingNumber =  Convert.ToString(rawData[row, 3]);
                salesDetail.OrderQty = Convert.ToInt32(rawData[row, 4]);
                salesDetail.ProductID = Convert.ToInt32(rawData[row, 5]);
                salesDetail.SpecialOfferID = Convert.ToInt32(rawData[row, 6]);
                salesDetail.UnitPrice = Convert.ToDouble(rawData[row, 7]);
                salesDetail.UnitPriceDiscount = Convert.ToDouble(rawData[row, 8]);
                salesDetail.LineTotal = Convert.ToDouble(rawData[row, 9]);
                salesDetail.ModifiedDate = DateTime.FromOADate((double)rawData[row, 10]);

                _salesDetails.Add(salesDetail);
            }


            watch.Stop();

            //milliseconds only.
            sb.AppendLine("Elapsed Time (processing data inside .NET) in sec: " + watch.Elapsed.TotalSeconds);

            Globals.ThisWorkbook.ThisApplication.EnableEvents = true;
            
            MessageBox.Show(sb.ToString());

        }

        public void Write()
        {
            if (_salesDetails == null || _salesDetails.Count == 0)
            {
                MessageBox.Show("Please load data first");
                return;
            }

            var watch = new Stopwatch();
            watch.Start();

            Globals.ThisWorkbook.ThisApplication.EnableEvents = false;
            Globals.Sheet1.tblSalesOrderDetails.SetDataBinding(_salesDetails, "", "SalesOrderID", "SalesOrderDetailID",
                                                               "CarrierTrackingNumber", "OrderQty", "ProductID",
                                                               "SpecialOfferID", "UnitPrice", "UnitPriceDiscount",
                                                               "LineTotal", "ModifiedDate");

            Globals.Sheet1.tblSalesOrderDetails.Disconnect();

            watch.Stop();
            Globals.ThisWorkbook.ThisApplication.EnableEvents = true;

            //about 100 secs
            MessageBox.Show("Elapsed time in sec: " + watch.Elapsed.TotalSeconds);
        }

        public void WriteFaster()
        {
            if (_salesDetails == null || _salesDetails.Count == 0)
            {
                MessageBox.Show("Please load data first");
                return;
            }

            var watch = new Stopwatch();
            watch.Start();

            Globals.ThisWorkbook.ThisApplication.EnableEvents = false;


            object[,] arrayOfSales = new object[_salesDetails.Count,10];

            for (int i = 0; i < _salesDetails.Count; i++)
            {
                arrayOfSales[i, 0] = _salesDetails[i].SalesOrderDetailID;
                arrayOfSales[i, 1] = _salesDetails[i].SalesOrderID;
                arrayOfSales[i, 2] = _salesDetails[i].CarrierTrackingNumber;
                arrayOfSales[i, 3] = _salesDetails[i].OrderQty;
                arrayOfSales[i, 4] = _salesDetails[i].ProductID;
                arrayOfSales[i, 5] = _salesDetails[i].SpecialOfferID;
                arrayOfSales[i, 6] = _salesDetails[i].UnitPrice;
                arrayOfSales[i, 7] = _salesDetails[i].UnitPriceDiscount;
                arrayOfSales[i, 8] = _salesDetails[i].LineTotal;
                arrayOfSales[i, 9] = _salesDetails[i].ModifiedDate;
            }

            Globals.Sheet1.tblSalesOrderDetails.DataBodyRange.Value2 = arrayOfSales;



            Globals.ThisWorkbook.ThisApplication.EnableEvents = true;

            watch.Stop();

            //about 16 secs
            MessageBox.Show("Elapsed time in sec: " + watch.Elapsed.TotalSeconds);
        }


    }
}