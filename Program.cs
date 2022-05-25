using System;
using System.Data;
using System.Data.OleDb;
using System.IO;
using System.Net;
using System.Text;
using System.Diagnostics;

namespace ExampleBugCellA1
{
    class Program
    {
        static void Main(string[] args)
        {

            const string USER_CREDENTIALS = "22876";
            const string PSW_CREDENTIALS = "xzHvihqwl1";
            const string FTP_URL = "ftp://ftp.partseurope.eu";
            const string FILE_NAME = "pricefile_AKRAPOVIC_v3.xls";
            const string FILE_URL = FTP_URL + "/"+ FILE_NAME;
            const string NAME_SHEET = "Pricelist";
            string LocalFullPathFile = Path.GetTempPath() + FILE_NAME;


            Console.WriteLine();
            Console.WriteLine("For read excel files you need to install Microsoft.ACE.OLEDB.12.0");
            Console.WriteLine("https://www.microsoft.com/en-US/download/details.aspx?id=13255");
            Console.WriteLine("Press the return to comtinue or '0 + return' to end");
            if (Console.ReadLine() == "0")
                Environment.Exit(0);


            Console.WriteLine("DownLoad the file from FTP repository");
            Console.WriteLine("Press the return key to start");
            Console.ReadLine();

            //DowloadFile
            using (WebClient request = new WebClient())
            {
                System.Threading.Thread.Sleep(1000);
                request.Credentials = new NetworkCredential( USER_CREDENTIALS,PSW_CREDENTIALS);
                byte[] fileData = request.DownloadData(FILE_URL);
                
                using (FileStream filetemp = File.Create(LocalFullPathFile))
                {
                    filetemp.Write(fileData, 0, fileData.Length);
                    filetemp.Close();
                }
            }
            Console.WriteLine();
            Console.WriteLine("I will show you then columns in the file");
            Console.WriteLine("Pay attention to the name of columns");
            Console.WriteLine("Press the return key to continue");
            Console.ReadLine();
            //Open the file
            String Provider = "Provider=Microsoft.ACE.OLEDB.12.0;Data Source=" +
                               LocalFullPathFile +
                               ";Extended Properties='Excel 8.0;HDR=YES';";
            using (OleDbConnection con = new OleDbConnection(Provider))
            {
                con.Open();
                OleDbCommand oconn = new OleDbCommand("Select * From [" + NAME_SHEET + "$]", con);
                OleDbDataAdapter sda = new OleDbDataAdapter(oconn);

                DataTable data = new DataTable();

                sda.Fill(data);

                data = data.DefaultView.ToTable();

                Console.WriteLine();
                foreach (DataColumn col in data.Columns)
                    Console.WriteLine(col.ColumnName +" / " + col.Ordinal );

            }

            Console.WriteLine();
            Console.WriteLine("The column A1 is lost");
            Console.WriteLine("Open the file and rename the cell A1");
            Console.WriteLine("Press the return key to Open the file");
            Console.ReadLine();
            Process.Start("explorer.exe", LocalFullPathFile);
            Console.WriteLine("Press the return key to continue");
            Console.ReadLine();
            //Open the file manually and change the name of cellA1
            using (OleDbConnection con = new OleDbConnection(Provider))
            {
                con.Open();
                OleDbCommand oconn = new OleDbCommand("Select * From [" + NAME_SHEET + "$]", con);
                OleDbDataAdapter sda = new OleDbDataAdapter(oconn);

                DataTable data = new DataTable();

                sda.Fill(data);

                data = data.DefaultView.ToTable();

                Console.WriteLine();
                foreach (DataColumn col in data.Columns)
                    Console.WriteLine(col.ColumnName + " / " + col.Ordinal);

            }
            Console.WriteLine();
            Console.WriteLine("Now compare the names of columns");
            Console.WriteLine("Press the return key to end");
            Console.ReadLine();
        }
    }
}
