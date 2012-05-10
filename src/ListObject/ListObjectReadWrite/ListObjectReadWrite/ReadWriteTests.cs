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

        public void ReadSlower()
        {
            _salesDetails.Clear();
            _salesDetails = new List<SalesOrderDetail>(122000);

            Stopwatch watch = new Stopwatch();
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
            //about 300 secs
            MessageBox.Show("Elapsed time in sec: " + watch.Elapsed.TotalSeconds);

        }

        public void ReadFaster()
        {
            _salesDetails.Clear();
            _salesDetails = new List<SalesOrderDetail>(122000);

            var watch = new Stopwatch();
            watch.Start();

            object[,] rawData = Globals.Sheet1.tblSalesOrderDetails.Range.Value2;
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

            //about 16 secs
            MessageBox.Show("Elapsed time in sec: " + watch.Elapsed.TotalSeconds);

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


            Globals.Sheet1.tblSalesOrderDetails.SetDataBinding(_salesDetails, "", "SalesOrderID", "SalesOrderDetailID",
                                                               "CarrierTrackingNumber", "OrderQty", "ProductID",
                                                               "SpecialOfferID", "UnitPrice", "UnitPriceDiscount",
                                                               "LineTotal", "ModifiedDate");

            Globals.Sheet1.tblSalesOrderDetails.Disconnect();

            watch.Stop();

            //about 100 secs
            MessageBox.Show("Elapsed time in sec: " + watch.Elapsed.TotalSeconds);
        }


    }
}