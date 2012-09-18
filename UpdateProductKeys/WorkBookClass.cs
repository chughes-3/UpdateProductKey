using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using Excel = Microsoft.Office.Interop.Excel;
using System.Diagnostics;
using System.Windows.Forms;

namespace UpdateProductKeys
{
    class WorkBookClass
    {
        Excel.Application xlApp;
        Excel.Workbook xlWBook = null;
        Excel.Worksheet xlWsheet = null;
        Excel.Range xlRangeWorking;
        internal class RowData
        {
            public string mMfr;
            public string mSerNo;
            public string mOsPK;
            public string mStatus;
            public string mRsult;
        }
        internal List<RowData> rowList = new List<RowData>();
        List<string> colHdrsExpected = new List<string>() { "mr_manufacturer", "mr_serial_number", "OS_Product_Key", "Current Status", "Result" };
        internal WorkBookClass()
        {
            try
            {
                xlApp = System.Runtime.InteropServices.Marshal.GetActiveObject("Excel.Application") as Excel.Application;
            }
            catch (Exception)
            {
                xlApp = new Excel.Application();
            }
            Microsoft.Office.Core.FileDialog fd = this.xlApp.get_FileDialog(Microsoft.Office.Core.MsoFileDialogType.msoFileDialogOpen);
            fd.AllowMultiSelect = false;
            fd.Filters.Clear();
            fd.Filters.Add("Excel Files", "*.xls;*.xlsx");
            fd.Filters.Add("All Files", "*.*");
            fd.Title = "Tax-Aide Update Product Keys";
            if (fd.Show() != 0)
            {
                fd.Execute();
            }
            xlApp.Visible = true;
            xlWBook = xlApp.ActiveWorkbook;
            xlWsheet = xlWBook.ActiveSheet;
            Excel.Range colHdrsRng = xlWsheet.Range["A1:E1"];
            object[,] colHdrsObj = new object[5, 1];
            colHdrsObj = colHdrsRng.Value2;
            var colHdrsStr = colHdrsObj.Cast<string>();
            //Debug.WriteLine(string.Join(Environment.NewLine + "\t",colHdrsStr.Select(x => x.ToString())));
            var result = colHdrsStr.SequenceEqual(colHdrsExpected);
            if (!result)
            {
                MessageBox.Show("The Column Headers in this spreadsheet do not conform to specification.\r\nIs the correct spreadsheet open?\r\n\r\nExiting", "Tax-Aide Update Product Keys");
                DisposeX();
            }
            //pull spreadsheet data across for faster performance
            xlRangeWorking = xlWsheet.UsedRange;
            object[,] sysObj = new object[5, xlRangeWorking.Rows.Count];
            sysObj = xlRangeWorking.Value2;
            rowList = (from idx1 in Enumerable.Range(1, sysObj.GetLength(0))
                          select new RowData { mMfr = (string)sysObj[idx1, 1], mSerNo = (string)sysObj[idx1, 2], mOsPK = (string)sysObj[idx1, 3], mStatus = (string)sysObj[idx1, 4], mRsult = (string)sysObj[idx1, 5] }).ToList<RowData>();
            rowList.Insert(0, new RowData());   //to get numbering rows the same as excel blank record at zero
            //Debug.WriteLine(sysData.GetType().ToString() + "  count elements " + sysData.Count());
            //Debug.WriteLine(sysData.ElementAt(1).mMfr);
            Debug.WriteLine(string.Join(Environment.NewLine + "\t",from row in rowList select row.mMfr + "  " + row.mSerNo ) );
        }

        private void DisposeX()
        {
            this.xlApp.ActiveWorkbook.Close();
            xlApp = null;
            Environment.Exit(1);
        }
        internal void Dispose()
        {
            this.xlApp.ActiveWorkbook.Close();
            xlApp = null;
        }
    }
}
