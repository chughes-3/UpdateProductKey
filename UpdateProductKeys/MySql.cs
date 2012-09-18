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
        internal MySql()
        {
            Console.WriteLine("statring sql");
            Console.WriteLine(System.Configuration.ConfigurationManager.ConnectionStrings["PracticeDB"].ToString());
            string dbConnString = System.Configuration.ConfigurationManager.ConnectionStrings["PracticeDB"].ToString();
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
            //'mr_manufacturer', 'mr_serial_number', 'OS_Product_Key'
            string mySqlCmd = 
                "SELECT `inv_db_id`, a.lic_db_id, `inv_flags`, a.mr_manufacturer, a.mr_serial_number, `os_product_key_type`, `os_product_key`, `valid`, `product_code` " + 
                "FROM inventory2012 AS a " + 
                "LEFT JOIN systemlic AS b ON a.mr_manufacturer = b.mr_manufacturer  AND a.mr_serial_number = b.mr_serial_number " + 
                " WHERE CONCAT( a.mr_manufacturer, a.mr_serial_number )" + 
                " IN ('Dell Inc.CR742N1','TOSHIBA7B357968Q','Compaq9X35KQDZD33W','Dell Inc.3T1ddffe','Dell Inc.3T182L1','Dell Inc.18Z82L1','IBML3HZH46')" + 
                " UNION " +
                " SELECT`inv_db_id` , b.lic_db_id,`inv_flags` , b.mr_manufacturer, b.mr_serial_number, `os_product_key_type`, `os_product_key`, `valid`, `product_code` " + 
                "FROM inventory2012 AS a " + 
                "RIGHT JOIN systemlic AS b ON a.mr_manufacturer = b.mr_manufacturer AND a.mr_serial_number = b.mr_serial_number " + 
                "WHERE CONCAT( b.mr_manufacturer, b.mr_serial_number ) " + 
                "IN ('Dell Inc.CR742N1', 'TOSHIBA7B357968Q', 'Compaq9X35KQDZD33W', 'Dell Inc.3T1ddffe', 'Dell Inc.3T182L1', 'Dell Inc.18Z82L1', 'IBML3HZH46') ";
            MySqlDataAdapter dbDataAdptr = new MySqlDataAdapter(mySqlCmd, dbConn);
            dbDataAdptr.SelectCommand.CommandType = System.Data.CommandType.Text;
            DataTable invLicCombined = new DataTable();
            try
            {
                dbDataAdptr.Fill(invLicCombined);
            }
            catch (MySqlException ex)
            {
                Console.WriteLine("Error: {0}", ex.ToString());
            }
            var qry = from DataRow row in invLicCombined.Rows where ( (int)row[1] != 0) select row; //row[7] != DBNull.Value ||  this is the license valid column if it is DBNull then no entry was in the license table for this system. Also if lic_db_id in inventory table = 0 then this is not a valid license db record.
            DataTable systemlicEx = qry.CopyToDataTable();  //define the LINQ qry then qry.copytotable is syntax for creating datatable from linq per msdn
            systemlicEx.Columns.RemoveAt(6);    //removes os_product_key from inventory table
            systemlicEx.Columns.RemoveAt(5);    //remove os_product_key_type
            systemlicEx.Columns.RemoveAt(2);//remove inv_flags
            systemlicEx.Columns.RemoveAt(0);    //remove inv_db_id
            systemlicEx.AcceptChanges();
            //At this point systemlicEX has records from systemlic table and the invLicCombined has records for all inventory table entries. The extra columns can be ignored since the dataAdpater update sql statement will be hand crafted to use just the columns of interest. It was really unnecessary to remove the extra columns from systemlicEX - it was done for easy reading in console window.
            OutTable2Console(systemlicEx);
            systemlicEx.Rows[3]["valid"] = 3;
            systemlicEx.Rows[3]["product_code"] = "NLK";
            systemlicEx.Rows[2]["product_code"] = "NLK";

            systemlicEx.Rows[2]["valid"] = 3;
            OutTable2Console(systemlicEx);
            mySqlCmd = "UPDATE `systemlic` " +
                "SET `valid` = @valid, `product_code` = @product_code " +
                "WHERE (`lic_db_id` = @lic_db_id)";

            MySqlCommand cmd = new MySqlCommand(mySqlCmd, dbConn);
            cmd.Parameters.Add("@valid", MySqlDbType.Byte,1,"valid");   //valid is tinyint = signed byte
            cmd.Parameters.Add("@product_code", MySqlDbType.VarChar, 80, "product_code");
            cmd.Parameters.Add("@lic_db_id", MySqlDbType.UInt32, 15, "lic_db_id");  // size is ignored for int32
            dbDataAdptr.UpdateCommand = cmd;
            var output = dbDataAdptr.Update(systemlicEx);
            Console.WriteLine("output = " + output);
            Console.ReadKey();
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
            syslic.PrimaryKey = new DataColumn[1] {syslic.Columns["lic_db_id"]};
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
