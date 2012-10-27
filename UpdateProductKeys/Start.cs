using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Diagnostics;

namespace UpdateProductKeys
{
    class Start
    {
        static void Main(string[] args)
        {
            // If start with no args will use the practice db. Must start with the argument -p to use the production db PID Chekcing not done in PracticeDB
            Debug.Listeners.Add(new TextWriterTraceListener(Console.Out));
            Debug.AutoFlush = true;
            Debug.Indent();
            Debug.WriteLine("Starting excel access");
            WorkBookClass xcel = new WorkBookClass((args.Length > 0) ? args[0] : string.Empty);   // opens spreadsheet brings spreadsheet data across into a List(rowData) 
            MySql dbAccess = new MySql((args.Length > 0) ? args[0] : string.Empty ); // initializes connection to database
            //process rows xcel method have to pass datbase object so that can do database method calls from xcel
            dbAccess.GetDataSet(xcel.rowList);
            xcel.ProcessRows(dbAccess); //passing dbaccess so method can access license and inventory datatables
            Debug.WriteLine("about to update sheet");
            //add lines at bottom of spreadsheet with update summary and table inquiry SQL
            xcel.ResultUserMess(dbAccess.UpdateSysLicTable().ToString(), dbAccess.UpdateInventoryTable().ToString());
            dbAccess.CloseConn();
            xcel.Dispose();
        }
    }
}
