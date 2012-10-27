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
        internal Excel.Range resultMessageCell;
        internal StringBuilder sqlQryResultInventory = new StringBuilder();
        internal StringBuilder sqlQryResultLicense = new StringBuilder();
        internal class RowData
        {
            public string mMfr;
            public string mSerNo;
            public string mOsPK;
            public string mStatus;
            public string mRsult;
        }
        internal List<RowData> rowList = new List<RowData>();
        PidChecker pidCheck;
        List<string> colHdrsExpected = new List<string>() { "mr_manufacturer", "mr_serial_number", "OS_Product_Key", "Current Status", "Result" };
        internal WorkBookClass(string production)
        {
            pidCheck = new PidChecker(production);  //instantiation for later use in checking keys
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
            //First make sure that working range is one area and has no blank rows at the end.
            if (xlWsheet.UsedRange.Areas.Count != 1)
            {
                MessageBox.Show("The used range on this spreadsheet is not contiguous, something is wrong.\r\nIs the correct spreadsheet open?\r\n\r\nExiting", "Tax-Aide Update Product Keys");
                DisposeX();
            }
            //xlApp.ScreenUpdating = true;
            xlRangeWorking = xlWsheet.UsedRange;
            while (xlApp.WorksheetFunction.CountA(xlRangeWorking.Offset[xlRangeWorking.Rows.Count - 1, 0].Resize[1, Type.Missing]) == 0)
            {//eliminates blank rows at end of working range from the working range
                xlRangeWorking = xlRangeWorking.Resize[xlRangeWorking.Rows.Count - 1, Type.Missing];
            }
            resultMessageCell = xlWsheet.Range[xlRangeWorking.Cells[(1 + xlRangeWorking.Rows.Count), 1], xlRangeWorking.Cells[(1 + xlRangeWorking.Rows.Count), 1]];
            //Console.WriteLine(xlRangeWorking.Row.ToString() + "  rowCnt= " + xlRangeWorking.Rows.Count.ToString());
            //pull spreadsheet data across for faster performance
            object[,] sysObj = new object[5, xlRangeWorking.Rows.Count];
            sysObj = xlRangeWorking.Value2;
            rowList = (from idx1 in Enumerable.Range(1, sysObj.GetLength(0))
                       select new RowData { mMfr = (string)sysObj[idx1, 1], mSerNo = (sysObj[idx1, 2] != null) ? sysObj[idx1, 2].ToString() : "", mOsPK = (string)sysObj[idx1, 3], mStatus = (string)sysObj[idx1, 4], mRsult = (string)sysObj[idx1, 5] }).ToList<RowData>();
            rowList.RemoveAt(0);    //Remove Column headers from list
            //Debug.WriteLine(sysData.GetType().ToString() + "  count elements " + sysData.Count());
            //Debug.WriteLine(sysData.ElementAt(1).mMfr);
            Debug.WriteLine(string.Join(Environment.NewLine + "\t", from row in rowList select row.mMfr + "  " + row.mSerNo));
            //Change topic setup status column to receive text strings
            try
            {
                xlWsheet.Range["D1:E1"].ColumnWidth = 40;
            }
            catch (Exception e)
            {
                MessageBox.Show("Excel appears to be in a mode of not taking programmatic input. Please close all open Excel copies and restart\r\nThe error message was " + e.Message);
                DisposeX();
            }
            userMessageCell = xlWsheet.Range["D1"];
            userMessageCell.Value2 = "Querying the database for these systems";
            userMessageCell.Font.Italic = true;
            userMessageCell.Font.Color = System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.BlueViolet);
        }

        internal void ProcessRows(MySql dbAccess)
        {
            userMessageCell.Value2 = "Updating local copy of license and inventory database entries";
            foreach (var row in rowList)
            {
                if (dbAccess.systemlicEx == null)
                {
                    row.mRsult = "No Entry in License Database";
                    UpdateWSheetRow(rowList.IndexOf(row));
                    InvDbAnalysis(dbAccess, row);
                    continue;
                }
                var qry = from DataRow record in dbAccess.systemlicEx.Rows
                          where ((record["mr_serial_number"] != DBNull.Value) ? (string)record["mr_serial_number"] : "") == row.mSerNo && ((record["mr_manufacturer"] != DBNull.Value) ? (string)record["mr_manufacturer"] : "") == row.mMfr
                          select new { rec = record, indx = dbAccess.systemlicEx.Rows.IndexOf(record) };
                switch (qry.Count())
                {
                    case 0:
                        row.mRsult = "No Entry in License Database";
                        UpdateWSheetRow(rowList.IndexOf(row));
                        InvDbAnalysis(dbAccess, row);
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
                Action statusKeyUpload000 = () => ProcessNewKeyLic(row, "Upload Product Key - VLK(Techsoup??)", rowlicTable);
                Action statusKeyUploadOEM = () => ProcessNewKeyLic(row, "Upload OEM COA Product Key", rowlicTable);
                Action statKeyForDownload = () => { string l5 = GetExistKeyLast5(rowlicTable); ProcessNewKeyLic(row, "Product Key Ending " + l5 + " available for download", rowlicTable); };
                Action statNLKForDownload = () => ProcessNewKeyLic(row, "NLK available for download", rowlicTable);
                LogicTable.Start()
                    .Condition(notValid0000000, "TFFFF")
                    .Condition(forceKeyUpload0, "-TTFF")
                    .Condition(forcKeyUplodOEM, "-FTFF")
                    .Condition(keyfordownload0, "-FFTF")
                    .Condition(nlkfordownload0, "-FF-T")
                    .Action(statusNotValid0000, "X    ")
                    .Action(statusKeyUploadOEM, "  X  ")
                    .Action(statusKeyUpload000, " X   ")
                    .Action(statKeyForDownload, "   X ")
                    .Action(statNLKForDownload, "    X");
                dbAccess.systemlicEx.Rows[licTableRow]["valid"] = rowlicTable["valid"];
                dbAccess.systemlicEx.Rows[licTableRow]["product_code"] = rowlicTable["product_code"];
                UpdateWSheetRow(rowList.IndexOf(row));
                if (sqlQryResultLicense.Length > 0)
                    sqlQryResultLicense.Append(" OR ");
                sqlQryResultLicense.Append("(`mr_manufacturer` = " + "'" + row.mMfr + "' AND " + "`mr_serial_number` = " + "'" + row.mSerNo + "')");
            }
        }

        private void InvDbAnalysis(MySql dbAccess, RowData row)
        {
            var qry = from DataRow record in dbAccess.invLicCombined.Rows
                      where ((record["mr_serial_number"] != DBNull.Value) ? (string)record["mr_serial_number"] : "") == row.mSerNo && ((record["mr_manufacturer"] != DBNull.Value) ? (string)record["mr_manufacturer"] : "") == row.mMfr
                      select new { rec = record, indx = dbAccess.invLicCombined.Rows.IndexOf(record) };
            switch (qry.Count())
            {
                case 0:
                    row.mRsult = "No Entry in Inventory or License Database";
                    UpdateWSheetRow(rowList.IndexOf(row));
                    return;
                case 1:
                    break;
                default:
                    row.mRsult = "Multiple Entries in Inventory Database - major DB error";
                    UpdateWSheetRow(rowList.IndexOf(row));
                    return;
            }
            //at this point have only 1 record which is correct and desirable
            DataRow rowInvTable = qry.ElementAtOrDefault(0).rec;    //row in license table to be processed
            int invTableRow = qry.ElementAtOrDefault(0).indx;   //index of that row - needed for later updating of the database after row processing
            Func<bool> testIfActive000 = () => ((sbyte)rowInvTable["active"] & 1) == 0;
            Func<bool> qualified4Nlk00 = () => ((sbyte)rowInvTable["active"] & 2) == 2;
            Func<bool> localW7LicAvail = () => ((sbyte)rowInvTable["inv_flags"] & 4) == 4;
            Func<bool> locW7withKtype0 = () => (string.IsNullOrWhiteSpace((string)rowInvTable["os_product_key_type"])) == false;
            //Func<bool> locW7withlast5c = () => (string.IsNullOrWhiteSpace((string)rowInvTable["os_partial_product_key"])) == false;
            Action errMessageNotActiv = () => { row.mRsult = "Flagged as Not Active System in Inventory database"; };
            Action errMessNotW7Eligab = () => { row.mStatus = "System in Inventory DB, but is not eligible for NLK Win 7 installation"; ProcessNewKeyInv(row, "System Status updated in Inventory DB to allow key download and install", rowInvTable); };
            Action statusPartKeyExist = () => { row.mStatus = "Windows 7, Key Type is " + rowInvTable["os_product_key_type"] + " with last 5 of " + rowInvTable["os_partial_product_key"]; ProcessNewKeyInv(row, "Inventory DB Product Key updated and ready for download", rowInvTable); };
            Action statusWindowsXP000 = () => { row.mStatus = "Windows XP or No data for Windows 7"; ProcessNewKeyInv(row, "Inventory DB Product Key added and is ready for download", rowInvTable); };
            //Action statusNlkEligible0 = () => { row.mStatus = "NLK Eligible"; ProcessNewKeyInv(row, "Product Key added and is ready for download", rowInvTable); };
            LogicTable.Start()
                .Condition(testIfActive000, "TFFF")
                .Condition(qualified4Nlk00, "-F-T")
                .Condition(localW7LicAvail, "-FT-")
                .Condition(locW7withKtype0, "---F")
                .Action(errMessageNotActiv, "X   ")
                .Action(errMessNotW7Eligab, " X  ")   // W7 Home premium will end on this action
                .Action(statusPartKeyExist, "  X ")
                .Action(statusWindowsXP000, "   X");
            dbAccess.invLicCombined.Rows[invTableRow]["active"] = rowInvTable["active"];
            dbAccess.invLicCombined.Rows[invTableRow]["inv_flags"] = rowInvTable["inv_flags"];
            dbAccess.invLicCombined.Rows[invTableRow]["os_product_key"] = rowInvTable["os_product_key"];
            UpdateWSheetRow(rowList.IndexOf(row));
            if (sqlQryResultInventory.Length != 0) 
                sqlQryResultInventory.Append(" OR ");
            sqlQryResultInventory.Append("(`mr_manufacturer` = " + "'" + row.mMfr + "' AND " + "`mr_serial_number` = " + "'" + row.mSerNo + "')");
        }

        private string GetExistKeyLast5(DataRow rowDataBase)
        {
            string decKey = EncodeDecodeKey((string)rowDataBase["product_code"], "-d");
            return decKey.Substring(24);
        }

        private void ProcessNewKeyLic(RowData row, string mstatusMess, DataRow rowLicDataBase)
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
            string tempstor = userMessageCell.Value2;
            userMessageCell.Value2 = "Checking Product Key this will take a while";
            string resKey = pidCheck.CheckProductKey(row.mOsPK);
            if (resKey != "Valid")
            {
                row.mRsult = resKey;
                return;
            }
            userMessageCell.Value2 = tempstor;
            string encKey = EncodeDecodeKey(row.mOsPK, "-e");  //Encrypt the key
            rowLicDataBase["product_code"] = encKey;
            rowLicDataBase["valid"] = ((sbyte)rowLicDataBase["valid"] & 249) + 2; //resets the 4 bit and sets the 2 bit
            row.mRsult = "Product Key Updated in the License Database";
            //Console.ReadKey();
        }

        void ProcessNewKeyInv(RowData row, string mResultMess, DataRow rowInvDataBase)
        {
            System.Text.RegularExpressions.Regex keyPattern = new System.Text.RegularExpressions.Regex("^([BCDFGHJKMPQRTVWXY2346789]{5}-){4}[BCDFGHJKMPQRTVWXY2346789]{5}$");
            System.Text.RegularExpressions.Regex keyPattern1 = new System.Text.RegularExpressions.Regex("^[BCDFGHJKMPQRTVWXY2346789]{25}$");
            if (row.mOsPK == null || !(keyPattern.IsMatch(row.mOsPK) || keyPattern1.IsMatch(row.mOsPK)))
            {
                row.mRsult = "Specified Product Key has incorrect form or alphanumeric characters";
                return;
            }
            string tempstor = userMessageCell.Value2;
            userMessageCell.Value2 = "Checking Product Key this will take a while";
            string resKey = pidCheck.CheckProductKey(row.mOsPK);
            if (resKey != "Valid")
            {
                row.mRsult = resKey;
                return;
            }
            userMessageCell.Value2 = tempstor;
            string encKey = EncodeDecodeKey(row.mOsPK, "-e");  //Encrypt the key
            rowInvDataBase["os_product_key"] = encKey;
            rowInvDataBase["active"] = ((sbyte)rowInvDataBase["active"] & 253); //resets NLK bit in case it is set
            rowInvDataBase["inv_flags"] = ((sbyte)rowInvDataBase["inv_flags"] & 251) + 4;   //sets the 4 bit indicating key is available
            rowInvDataBase["os_partial_product_key"] = row.mOsPK.Substring(24);
            row.mRsult = mResultMess;
        }

        private string EncodeDecodeKey(string str, string param)
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

        internal void ResultUserMess(string numLicRec, string numInvRec)
        {
            userMessageCell.Value2 = "Updating License database from local copy";
            resultMessageCell.Value2 = "Number of License table records updated = " + numLicRec;
            userMessageCell.Value2 = "Updating Inventory database from local copy";
            resultMessageCell.Offset[1,Type.Missing].Value2 = "Number of Inventory table records updated = " + numInvRec;
            resultMessageCell.Offset[2, Type.Missing].Value2 = "Below are license and inventory table search strings to check updated records, copy/paste the ENTIRE cell into the SQL Query \"Where\" clause box on the query page";
            resultMessageCell.Offset[3, Type.Missing].Value2 = sqlQryResultLicense.ToString();
            resultMessageCell.Offset[4, Type.Missing].Value2 = sqlQryResultInventory.ToString();
            RemoveUserMess();
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
