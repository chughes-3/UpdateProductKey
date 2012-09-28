using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using Excel = Microsoft.Office.Interop.Excel;
using System.Diagnostics;
using System.Windows.Forms;
using System.Data;
using System.IO.MemoryMappedFiles;

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
            userMessageCell.Value2 = "Updating local copy of license database entries";
            foreach (var row in rowList)
            {
                var qry = from DataRow record in dbAccess.systemlicEx.Rows
                          where ((record["mr_serial_number"] != DBNull.Value) ? (string)record["mr_serial_number"] : "") == row.mSerNo && ((record["mr_manufacturer"] != DBNull.Value) ? (string)record["mr_manufacturer"] : "") == row.mMfr
                          select new { rec = record, indx = dbAccess.systemlicEx.Rows.IndexOf(record) };
                switch (qry.Count())
                {
                    case 0:
                        row.mRsult = "No Entry in License Database";
                        UpdateWSheetRow(rowList.IndexOf(row));
                        InvDbAnalysis();
                        continue;
                    case 1:
                        break;
                    default:
                        row.mRsult = "Multiple Entries in License Database - major DB error";
                        UpdateWSheetRow(rowList.IndexOf(row));
                        continue;
                }
                //at this point have only 1 record which is correct and desirable
                DataRow rowlicTable = qry.ElementAtOrDefault(0).rec;    //row in license table to be processed
                int licTableRow = qry.ElementAtOrDefault(0).indx;   //index of that row - needed for later updating of the database after row processing
                Func<bool> notValid0000000 = () => ((sbyte)rowlicTable["valid"] & 1) == 0;
                Func<bool> forceKeyUpload0 = () => ((sbyte)rowlicTable["valid"] & 4) == 4;
                Func<bool> forcKeyUplodOEM = () => ((string)rowlicTable["product_code"]) == "OEM:SLP";
                Func<bool> keyfordownload0 = () => ((sbyte)rowlicTable["valid"] & 2) == 2;
                Func<bool> nlkfordownload0 = () => ((string)rowlicTable["product_code"]) == "NLK";
                Action statusNotValid0000 = () => row.mRsult = "Flagged \"Not Valid\" in License Database";
                Action statusKeyUpload000 = () => ProcessNewKey(row, "Upload Product Key - VLK(Techsoup??)", rowlicTable);
                Action statusKeyUploadOEM = () => ProcessNewKey(row, "Upload OEM COA Product Key", rowlicTable);
                Action statKeyForDownload = () => { string l5 = GetExistKeyLast5(rowlicTable); ProcessNewKey(row, "Product Key Ending " + l5 + " available for download", rowlicTable); };
                Action statNLKForDownload = () => ProcessNewKey(row, "NLK available for download", rowlicTable);
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
                dbAccess.systemlicEx.Rows[licTableRow]["valid"] = rowlicTable["valid"];
                dbAccess.systemlicEx.Rows[licTableRow]["product_code"] = rowlicTable["product_code"];
                UpdateWSheetRow(rowList.IndexOf(row));
            }
        }

        private string GetExistKeyLast5(DataRow rowDataBase)
        {
            string decKey = EncodeDecodeKey((string)rowDataBase["product_code"], "-d");
            return decKey.Substring(24);
        }

        private void ProcessNewKey(RowData row, string mstatusMess, DataRow rowDataBase)
        {
            //check key with pid chaecker here to be added after reg exp check
            row.mStatus = mstatusMess;
            System.Text.RegularExpressions.Regex keyPattern = new System.Text.RegularExpressions.Regex("^([BCDFGHJKMPQRTVWXY2346789]{5}-){4}[BCDFGHJKMPQRTVWXY2346789]{5}$");
            System.Text.RegularExpressions.Regex keyPattern1 = new System.Text.RegularExpressions.Regex("^[BCDFGHJKMPQRTVWXY2346789]{25}$");
            if (row.mOsPK == null || !(keyPattern.IsMatch(row.mOsPK) || keyPattern1.IsMatch(row.mOsPK)))
            {
                row.mRsult = "Specified Product Key has incorrect form or alphanumeric characters";
                return;
            }
            string encKey = EncodeDecodeKey(row.mOsPK, "-e");  //Encrypt the key
            rowDataBase["product_code"] = encKey;
            rowDataBase["valid"] = ((sbyte)rowDataBase["valid"] & 249) + 2; //resets the 4 bit and sets the 2 bit
            row.mRsult = "Product Key Updated ";
            //Console.ReadKey();
        }

        private static string EncodeDecodeKey(string str, string param)
        {//Calls KeyCrypt with either -e or -d as a parameter and gets a string back and forth
            string encDecKey;
            using (MemoryMappedFile mmf = MemoryMappedFile.CreateNew("EncDecString", 80))
            {
                System.Threading.Mutex mutex = new System.Threading.Mutex(true, "EncMemShare");
                using (MemoryMappedViewStream stream = mmf.CreateViewStream())
                {
                    System.IO.StreamWriter wtr = new System.IO.StreamWriter(stream);
                    wtr.Write(str);
                    wtr.Flush();
                }
                mutex.ReleaseMutex();
                string codeBase = System.Reflection.Assembly.GetExecutingAssembly().CodeBase;
                UriBuilder uri = new UriBuilder(codeBase);
                string path = Uri.UnescapeDataString(uri.Path);
                path = System.IO.Path.GetDirectoryName(path);   //now path is path to current executing assembly which is where KeyCrypt will be on install
                string keyCryptPath = System.IO.Path.Combine(path, "KeyCrypt.exe");
                Process p = Process.Start(keyCryptPath, param);
                p.WaitForExit();
                mutex.WaitOne();    //this process now waits until gets mutex back which means encryption is done
                //Console.ReadKey();  //wait to do pkencrypt
                using (MemoryMappedViewStream stream = mmf.CreateViewStream())
                {
                    System.IO.StreamReader rdr = new System.IO.StreamReader(stream);
                    encDecKey = rdr.ReadToEnd();
                    encDecKey = encDecKey.Substring(0, encDecKey.IndexOf("\0"));
                }
                mutex.ReleaseMutex();
                mutex.Dispose();
            }
            return encDecKey;
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
        void UpdateWSheetRow(int rowNo)
        {
            object[] objData = new object[5];
            Excel.Range rowRng = xlWsheet.Range[xlRangeWorking.Cells[rowNo + 2, 1], xlRangeWorking.Cells[rowNo + 2, 5]];
            objData[0] = rowList[rowNo].mMfr; objData[1] = rowList[rowNo].mSerNo; objData[2] = rowList[rowNo].mOsPK; objData[3] = rowList[rowNo].mStatus; objData[4] = rowList[rowNo].mRsult;
            rowRng.Value2 = objData;
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

        internal void RemoveUserMess()
        {
            userMessageCell.Font.Color = userMessageCell.Offset[0, 1].Font.Color;
            userMessageCell.Font.Italic = false;
            userMessageCell.Value2 = "Current Status";
        }
    }
}
