using Newtonsoft.Json;
using System;
using System.Collections.Generic;
using System.Data.Common;
using System.Data.OleDb;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace TestyDanychTestowych
{
    class Excel2JsonCoverter
    {
          // textBox2.Text= TestDataHelper2.mainHelper.WyliczNRB(textBox1.Text);
          //  TestDataHelper2.mainHelper mh = new TestDataHelper2.mainHelper();
          //  textBox2.Text = mh.GenerujNrKonta();
          //  mh.get_all_locales();
        public void Convert()
        {
            string pathToExcel = @"C:\Users\mathisse\Desktop\Banki2Json.xlsx";
            string sheetName = "Arkusz1";
            string destinationPath = @"C:\Users\mathisse\Desktop\Banki2Json_file.json";

            //Use this connection string if you have Office 2007+ drivers installed and 
            //your data is saved in a .xlsx file
            string connectionString = String.Format(@"
                Provider=Microsoft.ACE.OLEDB.12.0;
                Data Source={0};
                Extended Properties=""Excel 12.0 Xml;HDR=YES"" ", pathToExcel);

            //Creating and opening a data connection to the Excel sheet 
            using (var conn = new OleDbConnection(connectionString))
            {
                conn.Open();

                var cmd = conn.CreateCommand();
                cmd.CommandText = String.Format(
                    @"SELECT * FROM [{0}$]",
                    sheetName);

                using (var rdr = cmd.ExecuteReader())
                {

                    //LINQ query - when executed will create anonymous objects for each row
                    var query =
                        from DbDataRecord row in rdr
                        select new
                        {
                            nrb = row[0].ToString(),
                            detail = new details (
                                row[1].ToString().TrimEnd(),
                                row[2].ToString().TrimEnd(),
                                row[3].ToString().TrimEnd(),
                                row[4].ToString().TrimEnd(),
                                row[5].ToString().TrimEnd())
                            }
                        ;

                    //Generates JSON from the LINQ query
                    var json = JsonConvert.SerializeObject(query, Formatting.Indented);

                    //Write the file to the destination path    

                    File.WriteAllText(destinationPath, json, Encoding.UTF8);
                    
                }
            }
        }
    }
}
