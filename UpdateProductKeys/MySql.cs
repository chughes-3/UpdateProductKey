using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using MySql.Data.MySqlClient;
using System.Data;

namespace UpdateProductKeys
{
    class MySql
    {
        MySqlConnection dbConn;
        MySqlDataAdapter dbDataAdptr;
        internal DataTable invLicCombined = new DataTable();
        internal DataTable systemlicEx;
        internal MySql(string production)
        {
            Console.WriteLine("Starting sql");
            string dbConnString;
            if (production != "-p")
                dbConnString = System.Configuration.ConfigurationManager.ConnectionStrings["PracticeDB"].ToString();
            else
                dbConnString = System.Configuration.ConfigurationManager.ConnectionStrings["ProductionDB"].ToString();
            dbConn = new MySqlConnection(dbConnString);
            //dbConn.Open();
            //string stm = "SELECT VERSION()";
            //MySqlCommand cmd = new MySqlCommand(stm, dbConn);
            //string version = Convert.ToString(cmd.ExecuteScalar());
            //Console.WriteLine("MySQL version : {0}", version);
        }
        internal void CloseConn()
        {
            if (dbConn != null)
                dbConn.Close();
        }

        internal void GetDataSet(List<WorkBookClass.RowData> xcelRowList)
        {
            // See notes in OneNote for full explanation of below or google on sql parameter IN clause
            string cmdText =
                "SELECT `inv_db_id`, a.lic_db_id, `active`, `inv_flags`, a.mr_manufacturer, a.mr_serial_number, `os_product_key_type`, `os_partial_product_key`, `os_product_key`, `valid`, `product_code` " +
                "FROM inventory2012 AS a " +
                "LEFT JOIN systemlic AS b ON a.mr_manufacturer = b.mr_manufacturer  AND a.mr_serial_number = b.mr_serial_number " +
                " WHERE CONCAT( a.mr_manufacturer, a.mr_serial_number )" +
                " IN ({0})" +
                " UNION " +
                "SELECT `inv_db_id`, b.lic_db_id, `active`, `inv_flags`, b.mr_manufacturer, b.mr_serial_number, `os_product_key_type`, `os_partial_product_key`,  `os_product_key`, `valid`, `product_code` " +
                "FROM inventory2012 AS a " +
                "RIGHT JOIN systemlic AS b ON a.mr_manufacturer = b.mr_manufacturer AND a.mr_serial_number = b.mr_serial_number " +
                "WHERE CONCAT( b.mr_manufacturer, b.mr_serial_number ) " +
                "IN ({1}) ";
            //In string above will format {0} and {1} with tag names like @tag0,@tag1,@tag2 etc 1 for each row of spreadsheet
            string[] paramNames = xcelRowList.Select(
                (s, i) => "@tag" + i.ToString()
                    ).ToArray();
            string inClause = string.Join(",", paramNames);
            MySqlCommand selectCmd = new MySqlCommand(string.Format(cmdText, inClause, inClause));
            //At this point have full select command string now create parameter list of name value @tag0 is concatanation of row 1 mfr + ser etc
            for (int i = 0; i < paramNames.Length; i++)
            {
                selectCmd.Parameters.AddWithValue(paramNames[i], xcelRowList[i].mMfr + xcelRowList[i].mSerNo);
            }
            selectCmd.Connection = dbConn;
            dbDataAdptr = new MySqlDataAdapter(selectCmd);
            dbDataAdptr.SelectCommand.CommandType = System.Data.CommandType.Text;
            try
            {
                dbDataAdptr.Fill(invLicCombined);
            }
            catch (MySqlException ex)
            {
                Console.WriteLine("Error: {0}", ex.ToString());
            }
            var qry = from DataRow row in invLicCombined.Rows where ((int)row["lic_db_id"] != 0 && row["valid"] != DBNull.Value) select row; //||  this is the license valid column if it is DBNull then no entry was in the license table for this system. Also if lic_db_id in inventory table = 0 then this is not a valid license db record.
            if (qry.Count() > 0)
            {
                systemlicEx = qry.CopyToDataTable();  //define the LINQ qry then qry.copytotable is syntax for creating datatable from linq per msdn
                systemlicEx.Columns.Remove("inv_db_id");    //remove inv_db_id
                systemlicEx.Columns.Remove("active");   //remove active flags
                systemlicEx.Columns.Remove("inv_flags"); //remove inv_flags
                systemlicEx.Columns.Remove("os_product_key");    //removes os_product_key from inventory table
                systemlicEx.Columns.Remove("os_product_key_type");    //remove os_product_key_type
                systemlicEx.Columns.Remove("os_partial_product_key");    //removes os_product_key from inventory table
                systemlicEx.AcceptChanges();    //Changes to data from now on will be flagged for later updating of on line database
                //At this point systemlicEX has records from systemlic table and the invLicCombined has records for all inventory table entries. The extra columns can be ignored since the dataAdpater update sql statement will be hand crafted to use just the columns of interest. It was really unnecessary to remove the extra columns from systemlicEX - it was done for easy reading in console window.
                OutTable2Console(systemlicEx);
            }

        }

        internal int UpdateSysLicTable()
        {
            if (systemlicEx == null)
            {
                Console.WriteLine("Null System Lic");
                return 0;
            }
            string mySqlCmd = "UPDATE `systemlic` " +
                "SET `valid` = @valid, `product_code` = @product_code " +
                "WHERE (`lic_db_id` = @lic_db_id)";
            MySqlCommand cmd = new MySqlCommand(mySqlCmd, dbConn);
            cmd.Parameters.Add("@valid", MySqlDbType.Byte, 1, "valid");   //valid is tinyint = signed byte
            cmd.Parameters.Add("@product_code", MySqlDbType.VarChar, 80, "product_code");
            cmd.Parameters.Add("@lic_db_id", MySqlDbType.UInt32, 15, "lic_db_id");  // size is ignored for int32
            dbDataAdptr.UpdateCommand = cmd;
            var output = dbDataAdptr.Update(systemlicEx);
            Console.WriteLine("License Rows Updated = " + output);
            return output;
        }

        internal int UpdateInventoryTable()
        {
            string mySqlCmd = "UPDATE `inventory2012` " +
                "SET `active` = @active, `inv_flags` = @inv_flags, `os_product_key` = @os_product_key, `os_product_key_type` = 'manualUpdate', `os_partial_product_key` = @os_partial_product_key " +
                "WHERE (`inv_db_id` = @inv_db_id)";
            MySqlCommand cmd = new MySqlCommand(mySqlCmd, dbConn);
            cmd.Parameters.Add("@active", MySqlDbType.Byte, 1, "active");   //valid is tinyint = signed byte
            cmd.Parameters.Add("@inv_flags", MySqlDbType.Byte, 1, "inv_flags");   //valid is tinyint = signed byte
            cmd.Parameters.Add("@os_product_key", MySqlDbType.VarChar, 80, "os_product_key");
            cmd.Parameters.Add("@inv_db_id", MySqlDbType.UInt32, 15, "inv_db_id");  // size is ignored for int32
            cmd.Parameters.Add("@os_partial_product_key", MySqlDbType.VarChar, 5, "os_partial_product_key");
            dbDataAdptr.UpdateCommand = cmd;
            var output = dbDataAdptr.Update(invLicCombined);
            Console.WriteLine("Inventory Rows Updated = " + output);
            return output;
        }

        internal void getUpLicDataset()
        {

            string mySqlCmd =
                " SELECT lic_db_id, mr_manufacturer, mr_serial_number,  `valid`, `product_code` " +
                "FROM systemlic " +
                "WHERE CONCAT( mr_manufacturer, mr_serial_number ) " +
                "IN ('Dell Inc.CR742N1', 'TOSHIBA7B357968Q', 'Compaq9X35KQDZD33W', 'Dell Inc.3T1ddffe', 'Dell Inc.3T182L1', 'Dell Inc.18Z82L1', 'IBML3HZH46') ";
            MySqlDataAdapter dbDataAdptr = new MySqlDataAdapter(mySqlCmd, dbConn);
            dbDataAdptr.SelectCommand.CommandType = System.Data.CommandType.Text;
            MySqlCommandBuilder cmdBld = new MySqlCommandBuilder(dbDataAdptr);
            DataTable syslic = new DataTable();
            try
            {
                dbDataAdptr.Fill(syslic);
            }
            catch (MySqlException ex)
            {
                Console.WriteLine("Error: {0}", ex.ToString());
            }
            OutTable2Console(syslic);
            Console.WriteLine("valid = " + syslic.Columns["valid"].DataType.ToString() + "   " + syslic.Columns["product_code"].DataType.ToString());
            syslic.PrimaryKey = new DataColumn[1] { syslic.Columns["lic_db_id"] };
            syslic.Rows[3]["valid"] = 7;
            syslic.Rows[1]["product_code"] = "";
            OutTable2Console(syslic);
            var ouput = dbDataAdptr.Update(syslic);
            Console.WriteLine(cmdBld.GetUpdateCommand(true).CommandText);
            // UPDATE `systemlic` SET `mr_manufacturer` = @mr_manufacturer, `mr_serial_number` = @mr_serial_number, `valid` = @valid, `product_code` = @product_code WHERE ((`l ic_db_id` = @Original_lic_db_id) AND (`mr_manufacturer` = @Original_mr_manufactu rer) AND (`mr_serial_number` = @Original_mr_serial_number) AND (`valid` = @Origi nal_valid) AND (`product_code` = @Original_product_code))
            Console.WriteLine("updateOutput = " + ouput);
            Console.ReadKey();
        }
        private static void OutTable2Console(DataTable invLicCombined)
        {
            DataRow[] currentRows = invLicCombined.Select(null, null, DataViewRowState.CurrentRows);
            foreach (DataRow dr in currentRows)
            {
                foreach (DataColumn column in invLicCombined.Columns)
                    Console.Write("\t{0}", dr[column]); Console.WriteLine(" ");
                Console.WriteLine("NEXT ROW");
            }
        }

    }
}
