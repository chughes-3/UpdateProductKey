using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using Excel = Microsoft.Office.Interop.Excel;
using System.Diagnostics;
using System.Windows.Forms;
using System.Data;

namespace UpdateProductKeys
{
    class WorkBookClass
    {
        Excel.Application xlApp;
        Excel.Workbook xlWBook = null;
        Excel.Worksheet xlWsheet = null;
        Excel.Range xlRangeWorking;
        internal Excel.Range userMessageCell;
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
            if (fd.Show() == -1)
                try
                {
                    fd.Execute();
                }
                catch (Exception)
                {
                    MessageBox.Show("There is a problem opening the excel file.\r\nPlease close any open Excel applications and start the program again.");
                    Environment.Exit(1);
                }
            else
                Environment.Exit(1);
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
            object test = new object();
            test = xlWsheet.Range["B11"].Value2;
            Console.WriteLine(test.ToString());
            object[,] sysObj = new object[5, xlRangeWorking.Rows.Count];
            sysObj = xlRangeWorking.Value2;
            rowList = (from idx1 in Enumerable.Range(1, sysObj.GetLength(0))
                       select new RowData { mMfr = (string)sysObj[idx1, 1], mSerNo = (sysObj[idx1, 2] != null) ? sysObj[idx1, 2].ToString() : "", mOsPK = (string)sysObj[idx1, 3], mStatus = (string)sysObj[idx1, 4], mRsult = (string)sysObj[idx1, 5] }).ToList<RowData>();
            rowList.RemoveAt(0);    //Remove Column headers from list
            //Debug.WriteLine(sysData.GetType().ToString() + "  count elements " + sysData.Count());
            //Debug.WriteLine(sysData.ElementAt(1).mMfr);
            Debug.WriteLine(string.Join(Environment.NewLine + "\t", from row in rowList select row.mMfr + "  " + row.mSerNo));
            //Change topic setup status column to receive text strings
            xlWsheet.Range["D1:E1"].ColumnWidth = 40;
            userMessageCell = xlWsheet.Range["D1"];
            userMessageCell.Value2 = "Querying the database for these systems";
            userMessageCell.Font.Italic = true;
            userMessageCell.Font.Color = System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.BlueViolet);
        }

        internal void ProcessRows(MySql dbAccess)
        {
            
            userMessageCell.Value2 = "Updating Data";
            foreach (var row in rowList)
            {
                var qry = from DataRow record in dbAccess.systemlicEx.Rows
                          where (string)record["mr_serial_number"] == row.mSerNo  && (string)record["mr_manufacturer"] == row.mMfr
                          select record;
                Console.WriteLine("count = " + qry.Count().ToString() + "   entry= ");
                switch (qry.Count())
                {
                    case 0:
                        row.mStatus = "No Entry in License Database";
                        InvDbAnalysis();
                        continue;
                    case 1:
                        break;
                    default:
                        row.mStatus = "Multiple Entries in License Database - major DB error";
                        continue;
                }
                DataTable dbResult = qry.CopyToDataTable();
                DataRowCollection dbResultRows = dbResult.Rows;
                Func<bool> notValid0000000 = () => ((sbyte)dbResultRows[0]["valid"] & 1) == 0;
                Func<bool> forceKeyUpload0 = () => ((sbyte)dbResultRows[0]["valid"] & 4) == 4;
                Func<bool> forcKeyUplodOEM = () => ((string)dbResultRows[0]["product_code"]) == "OEM:SLP";
                Func<bool> keyfordownload0 = () => ((sbyte)dbResultRows[0]["valid"] & 2) == 2;
                Func<bool> nlkfordownload0 = () => ((string)dbResultRows[0]["product_code"]) == "NLK";
                Action statusNotValid0000 = () => row.mStatus = "Flagged \"Not Valid\" in License Database";
                Action statusKeyUpload000 = () => row.mStatus = "Upload Product Key - VLK(Techsoup??)";
                Action statusKeyUploadOEM = () => row.mStatus = "Upload OEM COA Product Key";
                Action statKeyForDownload = () => row.mStatus = "Non-NLK Key available for download";
                Action statNLKForDownload = () => row.mStatus = "NLK available for download";
                LogicTable.Start()
                    .Condition(notValid0000000, "T----")
                    .Condition(forceKeyUpload0, "FTTFF")
                    .Condition(forcKeyUplodOEM, "FFTFF")
                    .Condition(keyfordownload0, "FFFTF")
                    .Condition(nlkfordownload0, "FFF-T")
                    .Action(statusNotValid0000, "X    ")
                    .Action(statusKeyUploadOEM, "  X  ")
                    .Action(statusKeyUpload000, " X   ")
                    .Action(statKeyForDownload, "   X ")
                    .Action(statNLKForDownload, "    X");
            }

            
        }

        private void InvDbAnalysis()
        {
        }

        internal void UpdateWSheet()
        {
            object[,] objData = new object[rowList.Count, 5];
            Excel.Range tableIn = xlWsheet.Range["A2:E" + (rowList.Count + 1).ToString()];
            Debug.WriteLine("output range = " + "A2:E" + (rowList.Count + 1).ToString());
            for (int i = 0; i < rowList.Count; i++)
            {
                objData[i, 0] = rowList[i].mMfr; objData[i, 1] = rowList[i].mSerNo; objData[i, 2] = rowList[i].mOsPK; objData[i, 3] = rowList[i].mStatus; objData[i, 4] = rowList[i].mRsult;
            }
            tableIn.ClearContents();
            tableIn.Value2 = objData;
            userMessageCell.Font.Color = userMessageCell.Offset[0, 1].Font.Color;
            userMessageCell.Font.Italic = false;
            userMessageCell.Value2 = "Current Status";
        }
        private void DisposeX()
        {
            if (xlApp.ActiveWorkbook != null)
            {
                this.xlApp.ActiveWorkbook.Close(); 
            }
            xlApp = null;
            Environment.Exit(1);
        }
        internal void Dispose()
        {
            if (xlApp.ActiveWorkbook != null)
            {
                this.xlApp.ActiveWorkbook.Close(); 
            }
            xlApp = null;
        }
    }
}
