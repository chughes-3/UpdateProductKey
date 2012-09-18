﻿using System;
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
            Debug.Listeners.Add(new TextWriterTraceListener(Console.Out));
            Debug.AutoFlush = true;
            Debug.Indent();
            Debug.WriteLine("starting debug");
            WorkBookClass xcel = new WorkBookClass();   // opens spreadsheet brings spreadsheet data across into a List(rowData) 
            MySql dbAccess = new MySql(); // initializes connection to database
            //process rows xcel method have to pass datbase object so that can do database method calls from xcel
            dbAccess.GetDataSet(xcel.rowList);
            dbAccess.CloseConn();
            xcel.Dispose();
        }
    }
}
