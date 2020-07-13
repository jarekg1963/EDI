using Renci.SshNet;
using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Data.SqlClient;
using System.Drawing;
using System.Globalization;
using System.IO;
using System.IO.Compression;
using System.Linq;
using System.Net;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using Excel = Microsoft.Office.Interop.Excel;

namespace WindowsFormsApp3
{
    public partial class Form1 : Form
    {
        public Form1()
        {

            InitializeComponent();
        }

        public const string susername = "SN_Tool";
        public const string spassword = "46qD8N4w";
        public string shost = "ftp://172.26.59.13";
        public string connetionStringTarget = @"Data Source=DESKTOP-M9PRPPC\MSSQLSERVER01;Initial Catalog=TMS;Integrated Security=True;";

        public string katalogprogramu = "";

        private void button1_Click(object sender, EventArgs e)
        {

            // zmienne pomocnicze 
            string strDzien;
            string strMc;
            double licznik = 0;
            string zmSoldtoName;
            string zmShiptoName;
            string zmShiptoAddress;

            // source 
            string connetionStringSource;
            string sqlSource;
            SqlCommand Commandsource;

            SqlDataReader DataReadersource;
            SqlConnection cnnSource;
         
            connetionStringSource = "Server=172.26.61.125;Initial Catalog=DWH;User ID=CCM_DCMon_sql_usr;Password=LolkaBolka+!1234";
            cnnSource = new SqlConnection(connetionStringSource);
            cnnSource.Open();
            sqlSource = "select * from dbo.SN";
            Commandsource = new SqlCommand(sqlSource, cnnSource);
            DataReadersource = Commandsource.ExecuteReader();
        
            SqlConnection cnnTarget;
            SqlDataAdapter adapterTarget = new SqlDataAdapter();
            string sqlTarget;
            cnnTarget = new SqlConnection(connetionStringTarget);
            cnnTarget.Open();

            while (DataReadersource.Read())
            {
                licznik++;
                Console.WriteLine(licznik);
                zmShiptoName = DataReadersource.GetValue(5).ToString().Replace("'", "");
                zmShiptoAddress = DataReadersource.GetValue(6).ToString().Replace("'", "");
                zmSoldtoName = DataReadersource.GetValue(12).ToString().Replace("'", "");



                #region Data
                DateTime zmDateDateScanned = (DateTime)DataReadersource.GetValue(2);
                int dzien = zmDateDateScanned.Day;
                if (dzien < 10)
                {
                    strDzien = "0" + dzien.ToString();
                }
                else
                {
                    strDzien = dzien.ToString();
                }
                int mc = zmDateDateScanned.Month;
                if (mc < 10)
                {
                    strMc = "0" + mc.ToString();
                }
                else
                {
                    strMc = mc.ToString();
                }
                int rok = zmDateDateScanned.Year;
                string zmData = rok.ToString() + "- " + strMc + "-" + strDzien;
                #endregion
                sqlTarget = $" insert into SN (SerialNumber, Material, DateScanned , SalesOrganisation , ShiptoNumber , ShiptoName , ShiptoAddress , Warehouse , WarehouseName , SAPDeliveryNumber , SoldtoNumber,  SoldtoName , FileID, PalletID ) " +
                    $"values (  \'  {  DataReadersource.GetValue(0)  }  \', \'  {  DataReadersource.GetValue(1)  }  \'  ,  \' {  zmData  } \'    , \'  {  DataReadersource.GetValue(3)  }  \'  , \'  {  DataReadersource.GetValue(4)  }  \' ," +
                    $" \' { zmShiptoName  }  \',  \' {  zmShiptoAddress  }  \' ,  \'  {  DataReadersource.GetValue(7)  }  \' ,  \'  {  DataReadersource.GetValue(8)  }  \' ,  \'  {  DataReadersource.GetValue(9)  }  \'," +
                    $" \'  {  DataReadersource.GetValue(10)  }  \',  \'  {  zmSoldtoName  }  \',  \'  {  DataReadersource.GetValue(12)  }  \' ,  \'  {  DataReadersource.GetValue(13)  }  \') ";

                adapterTarget.InsertCommand = new SqlCommand(sqlTarget, cnnTarget);
                adapterTarget.InsertCommand.ExecuteNonQuery();
            }

            cnnTarget.Close();
            cnnSource.Close();

        }



        private void Form1_Load(object sender, EventArgs e)
        {
            // TODO: This line of code loads data into the 'tMSDataSet.SN' table. You can move, or remove it, as needed.


        }

        private void button2_Click(object sender, EventArgs e)
        {
            string sciezkaplik = @"c:\Praca\DC\Gorzow\serialextract-130720.txt";
            string SerialNumberCeva = "";
            string MaterialCeva = "";
            string SalesOrganisationCeva = "";
            string ShiptoNumberCeva = "";
            string WarehouseCeva;
            string WarehouseNameCeva;
            string SAPDeliveryNumberCeva = "";
            string SoldtoNumberCeva = "";
            string DateScannedCevaString;
            string ShiptoNameCeva;
            string ShiptoAddressCeva;
            string SoldtoNameCeva;
            int FileIDCeva;
            string PalletIDCeva;
            string sqlCeva;
            int w1;
            int w2;
            int w3;
            int w4;
            int w5;
            int w6;
            int dlugoscLini;

            SqlConnection cnnTarget;
            cnnTarget = new SqlConnection(connetionStringTarget);
            SqlDataAdapter adapterTarget = new SqlDataAdapter();
            cnnTarget.Open();
            int licznik = 0;

            foreach (string line in File.ReadLines(sciezkaplik, Encoding.UTF8))
            {
                licznik++;
                if (licznik > 1)
                {
                    w1 = line.Trim().IndexOf("\t", 1);
                    w2 = line.Trim().IndexOf("\t", w1 + 1);
                    w3 = line.Trim().IndexOf("\t", w2 + 1);
                    w4 = line.Trim().IndexOf("\t", w3 + 1);
                    w5 = line.Trim().IndexOf("\t", w4 + 1);
                    w6 = line.Trim().IndexOf("\t", w5 + 1);
                    dlugoscLini = line.Trim().Length;


                    SerialNumberCeva = line.Trim().Substring(w4 + 1, w1 - 1).Trim();
                    MaterialCeva = line.Trim().Substring(w5 + 1, dlugoscLini - w5 - 1).Trim();
                    SalesOrganisationCeva = line.Trim().Substring(0, 4).Trim();
                    ShiptoNumberCeva = line.Trim().Substring(w2 + 1, w3 - w2).Trim();
                    SAPDeliveryNumberCeva = line.Trim().Substring(5, w1 - 5).Trim();
                    DateScannedCevaString = line.Trim().Substring(w3 + 1, w4 - w3).Trim();
                    ShiptoNameCeva = "";
                    ShiptoAddressCeva = "";
                    SoldtoNumberCeva = "";
                    WarehouseCeva = "TPV_Gorzow";
                    WarehouseNameCeva = "TPV Displays Polska Sp. z o.o.";
                    SoldtoNameCeva = "";
                    FileIDCeva = 1;
                    PalletIDCeva = "";

                    sqlCeva = $" insert into SN_JG (SerialNumber, Material, DateScanned , SalesOrganisation , ShiptoNumber , ShiptoName , ShiptoAddress , Warehouse , WarehouseName , SAPDeliveryNumber , SoldtoNumber,  SoldtoName , FileID, PalletID ) " +
                       $"values (  \'  { SerialNumberCeva } \', \'  {  MaterialCeva  }  \'  ,  \' {  DateScannedCevaString  } \'    , \'  {  SalesOrganisationCeva  }  \'  , \'  {  ShiptoNumberCeva  }  \' ," +
                       $" \' { ShiptoNameCeva  }  \',  \' {  ShiptoAddressCeva  }  \' ,  \'  {  WarehouseCeva  }  \' ,  \'  {  WarehouseNameCeva   }  \' ,  \'  {  SAPDeliveryNumberCeva  }  \'," +
                       $" \'  {  ShiptoNumberCeva  }  \',  \'  {  SoldtoNameCeva  }  \',  \'  {  FileIDCeva  }  \' ,  \'  {  PalletIDCeva  }  \') ";

                    adapterTarget.InsertCommand = new SqlCommand(sqlCeva, cnnTarget);
                    adapterTarget.InsertCommand.ExecuteNonQuery();
                    Console.WriteLine(sqlCeva);
                }
            }
        }

        private void button3_Click(object sender, EventArgs e)
        {
            string sciezkaplik = @"c:\Praca\DC\Venlo\serialextract-070720.txt";
            string SerialNumberCeva = "";
            string MaterialCeva = "";
            string SalesOrganisationCeva = "";
            string ShiptoNumberCeva = "";
            string WarehouseCeva;
            string WarehouseNameCeva;
            string SAPDeliveryNumberCeva = "";
            string SoldtoNumberCeva = "";
            string DateScannedCevaString;
            string ShiptoNameCeva;
            string ShiptoAddressCeva;
            string SoldtoNameCeva;
            int FileIDCeva;
            string PalletIDCeva;
            int licznik = 0;

            string sqlCeva;

            int w1;
            int w2;
            int w3;
            int w4;
            int w5;
            int w6;
            int dlugoscLini;

            SqlConnection cnnTarget;
            cnnTarget = new SqlConnection(connetionStringTarget);
            SqlDataAdapter adapterTarget = new SqlDataAdapter();
            cnnTarget.Open();

            foreach (string line in File.ReadLines(sciezkaplik, Encoding.UTF8))
            {
                licznik++;
                if (licznik > 1)
                {
                    w1 = line.Trim().IndexOf("\t", 1);
                    w2 = line.Trim().IndexOf("\t", w1 + 1);
                    w3 = line.Trim().IndexOf("\t", w2 + 1);
                    w4 = line.Trim().IndexOf("\t", w3 + 1);
                    w5 = line.Trim().IndexOf("\t", w4 + 1);
                    w6 = line.Trim().IndexOf("\t", w5 + 1);
                    dlugoscLini = line.Trim().Length;

                    if (w6 == -1)
                    {
                        MaterialCeva = line.Trim().Substring(w5 + 1, dlugoscLini - w5 - 1).Trim();
                    }
                    else
                    {
                        MaterialCeva = line.Trim().Substring(w6, dlugoscLini - w6).Trim();
                    }


                    SerialNumberCeva = line.Trim().Substring(w4 + 1, w1 - 1).Trim();

                    SalesOrganisationCeva = line.Trim().Substring(0, 4).Trim();
                    ShiptoNumberCeva = line.Trim().Substring(w2 + 1, w3 - w2).Trim();
                    SAPDeliveryNumberCeva = line.Trim().Substring(5, w1 - 5).Trim();
                    DateScannedCevaString = line.Trim().Substring(w3 + 1, w4 - w3).Trim();
                    ShiptoNameCeva = "";
                    ShiptoAddressCeva = "";
                    SoldtoNumberCeva = "";
                    WarehouseCeva = "Venlo";
                    WarehouseNameCeva = "Venlo";
                    SoldtoNameCeva = "";
                    FileIDCeva = 2;
                    PalletIDCeva = "";

                    sqlCeva = $" insert into SN_JG (SerialNumber, Material, DateScanned , SalesOrganisation , ShiptoNumber , ShiptoName , ShiptoAddress , Warehouse , WarehouseName , SAPDeliveryNumber , SoldtoNumber,  SoldtoName , FileID, PalletID ) " +
                       $"values (  \'  { SerialNumberCeva } \', \'  {  MaterialCeva  }  \'  ,  \' {  DateScannedCevaString  } \'    , \'  {  SalesOrganisationCeva  }  \'  , \'  {  ShiptoNumberCeva  }  \' ," +
                       $" \' { ShiptoNameCeva  }  \',  \' {  ShiptoAddressCeva  }  \' ,  \'  {  WarehouseCeva  }  \' ,  \'  {  WarehouseNameCeva   }  \' ,  \'  {  SAPDeliveryNumberCeva  }  \'," +
                       $" \'  {  ShiptoNumberCeva  }  \',  \'  {  SoldtoNameCeva  }  \',  \'  {  FileIDCeva  }  \' ,  \'  {  PalletIDCeva  }  \') ";

                    adapterTarget.InsertCommand = new SqlCommand(sqlCeva, cnnTarget);
                    adapterTarget.InsertCommand.ExecuteNonQuery();
                    Console.WriteLine(sqlCeva);
                }
            }
        }

        private void button4_Click(object sender, EventArgs e)
        {
            Excel.Application xlApp = new Excel.Application();
            Excel.Workbook xlWorkbook = xlApp.Workbooks.Open(@"c:\Praca\DC\Daganzo\SN outbound.xls");
            Excel._Worksheet xlWorksheet = xlWorkbook.Sheets[1];
            Excel.Range xlRange = xlWorksheet.UsedRange;

            int rowCount = xlRange.Rows.Count;
            int colCount = xlRange.Columns.Count;

            #region zmienne do budowy select Daganzo
            string SerialNumberDaganzo;
            DateTime DateScannedDaganzo;

            string MaterialDaganzo;
            string SAPDeliveryNumberDaganzo;
            string ShiptoNameDaganzo;

            // do wyjasnienia 
            string SalesOrganisationDaganzo;
            string ShiptoNumberDaganzo;
            string ShiptoAddressDaganzo;
            string WarehouseDaganzo;
            string WarehouseNameDaganzo;
            string SoldtoNumberDaganzo;
            string SoldtoNameDaganzo;
            string FileIDDaganzo;
            string PalletIDDaganzo;
            string sqlDaganzo;

            #endregion

            // czesc wspolna dla wszystkich procedur 
            SqlConnection cnnTarget;
            cnnTarget = new SqlConnection(connetionStringTarget);
            SqlDataAdapter adapterTarget = new SqlDataAdapter();
            cnnTarget.Open();

            for (int i = 2; i <= rowCount; i++)
            {
                string dtString = xlRange.Cells[i, 7].Value2.ToString();
                DateScannedDaganzo = DateTime.Parse(ConvertToDateTime(dtString));

                string DateScannedDaganzoString = DateScannedDaganzo.ToString("M/dd/yyyy", CultureInfo.InvariantCulture);
                if (String.IsNullOrEmpty(xlRange.Cells[i, 5].Value2))
                    SerialNumberDaganzo = "";
                else SerialNumberDaganzo = xlRange.Cells[i, 5].Value2.ToString();
                MaterialDaganzo = xlRange.Cells[i, 2].Value2.ToString();
                ShiptoNameDaganzo = xlRange.Cells[i, 9].Value2.ToString();
                SAPDeliveryNumberDaganzo = xlRange.Cells[i, 1].Value2.ToString();
                SalesOrganisationDaganzo = "";
                ShiptoNumberDaganzo = "";
                ShiptoNameDaganzo = "";
                ShiptoAddressDaganzo = "";
                WarehouseDaganzo = "";
                WarehouseNameDaganzo = "";
                SoldtoNumberDaganzo = "";
                SoldtoNameDaganzo = "";
                FileIDDaganzo = "";
                PalletIDDaganzo = "";

                sqlDaganzo = $" insert into SN_JG (SerialNumber, Material, DateScanned , SalesOrganisation , ShiptoNumber , ShiptoName , ShiptoAddress , Warehouse , WarehouseName , SAPDeliveryNumber , SoldtoNumber,  SoldtoName , FileID, PalletID ) " +
                      $"values (  \'  { SerialNumberDaganzo } \', \'  {  MaterialDaganzo  }  \'  ,  \'{DateScannedDaganzoString}\'    , \'  {  SalesOrganisationDaganzo  }  \'  , \'  {  ShiptoNumberDaganzo  }  \' ," +
                      $" \' { ShiptoNameDaganzo  }  \',  \' {  ShiptoAddressDaganzo  }  \' ,  \'  {  WarehouseDaganzo  }  \' ,  \'  {  WarehouseNameDaganzo   }  \' ,  \'  {  SAPDeliveryNumberDaganzo  }  \'," +
                      $" \'  {  ShiptoNumberDaganzo  }  \',  \'  {  SoldtoNameDaganzo  }  \',  \'  {  FileIDDaganzo  }  \' ,  \'  {  PalletIDDaganzo  }  \') ";

                adapterTarget.InsertCommand = new SqlCommand(sqlDaganzo, cnnTarget);
                adapterTarget.InsertCommand.ExecuteNonQuery();

            }

            MessageBox.Show("Finished OK ");


        }

        public static string ConvertToDateTime(string strExcelDate)
        {
            double excelDate;
            try
            {
                excelDate = Convert.ToDouble(strExcelDate);
            }
            catch
            {
                return strExcelDate;
            }
            if (excelDate < 1)
            {
                throw new ArgumentException("Excel dates cannot be smaller than 0.");
            }
            DateTime dateOfReference = new DateTime(1900, 1, 1);
            if (excelDate > 60d)
            {
                excelDate = excelDate - 2;
            }
            else
            {
                excelDate = excelDate - 1;
            }
            return dateOfReference.AddDays(excelDate).ToShortDateString();
        }

        private void button5_Click(object sender, EventArgs e)
        {
            Excel.Application xlApp = new Excel.Application();
            Excel.Workbook xlWorkbook = xlApp.Workbooks.Open(@"c:\Praca\DC\Batta\Shipped SNs WK27.xlsx");
            Excel._Worksheet xlWorksheet = xlWorkbook.Sheets[1];
            Excel.Range xlRange = xlWorksheet.UsedRange;

            int rowCount = xlRange.Rows.Count;
            int colCount = xlRange.Columns.Count;

            #region zmienne do budowy select Bata
            string SerialNumberBatta;
            DateTime DateScannedBatta;

            string MaterialBatta;
            string SAPDeliveryNumberBatta;
            string ShiptoNameBatta;

            // do wyjasnienia 
            string SalesOrganisationBatta;
            string ShiptoNumberBatta;
            string ShiptoAddressBatta;
            string WarehouseBatta;
            string WarehouseNameBatta;
            string SoldtoNumberBatta;
            string SoldtoNameBatta;
            string FileIDBatta;
            string PalletIDBatta;
            string sqlBatta;

            #endregion

           
            SqlConnection cnnTarget;
       
            cnnTarget = new SqlConnection(connetionStringTarget);
            SqlDataAdapter adapterTarget = new SqlDataAdapter();
            cnnTarget.Open();

            for (int i = 2; i <= rowCount; i++)
            {
                string dtString = xlRange.Cells[i, 10].Value2.ToString();
                DateScannedBatta = DateTime.Parse(ConvertToDateTime(dtString));

                string DateScannedBattaString = DateScannedBatta.ToString("M/dd/yyyy", CultureInfo.InvariantCulture);
                if (String.IsNullOrEmpty(xlRange.Cells[i, 7].Value2))
                    SerialNumberBatta = "";
                else SerialNumberBatta = xlRange.Cells[i, 7].Value2.ToString();
                MaterialBatta = xlRange.Cells[i, 5].Value2.ToString();
                // ?? ShiptoNameBatta = xlRange.Cells[i, 9].Value2.ToString();
                SAPDeliveryNumberBatta = xlRange.Cells[i, 21].Value2.ToString().Substring(5, 10);
                SalesOrganisationBatta = xlRange.Cells[i, 21].Value2.ToString().Substring(0, 4);
                ShiptoNumberBatta = "";
                ShiptoNameBatta = "";
                ShiptoAddressBatta = "";
                WarehouseBatta = "";
                WarehouseNameBatta = "";
                SoldtoNumberBatta = "";
                SoldtoNameBatta = "";
                FileIDBatta = "";
                PalletIDBatta = "";

                sqlBatta = $" insert into SN_JG (SerialNumber, Material, DateScanned , SalesOrganisation , ShiptoNumber , ShiptoName , ShiptoAddress , Warehouse , WarehouseName , SAPDeliveryNumber , SoldtoNumber,  SoldtoName , FileID, PalletID ) " +
                      $"values (  \'  { SerialNumberBatta } \', \'  {  MaterialBatta  }  \'  ,  \'{DateScannedBattaString}\'    , \'  {  SalesOrganisationBatta  }  \'  , \'  {  ShiptoNumberBatta  }  \' ," +
                      $" \' { ShiptoNameBatta  }  \',  \' {  ShiptoAddressBatta  }  \' ,  \'  {  WarehouseBatta  }  \' ,  \'  {  WarehouseNameBatta   }  \' ,  \'  {  SAPDeliveryNumberBatta  }  \'," +
                      $" \'  {  ShiptoNumberBatta  }  \',  \'  {  SoldtoNameBatta  }  \',  \'  {  FileIDBatta  }  \' ,  \'  {  PalletIDBatta  }  \') ";

                adapterTarget.InsertCommand = new SqlCommand(sqlBatta, cnnTarget);
                adapterTarget.InsertCommand.ExecuteNonQuery();


            }
            MessageBox.Show("Skonczone OK");
        }

   
        private void ftpDownload(string sourceFile, string destinationFile)
        {

            FtpWebRequest request =
                (FtpWebRequest)WebRequest.Create(sourceFile);
            request.Credentials = new NetworkCredential(susername, spassword);
            request.Method = WebRequestMethods.Ftp.DownloadFile;
            request.UseBinary = true;

            using (Stream ftpStream = request.GetResponse().GetResponseStream())
            using (Stream fileStream = File.Create(destinationFile))
            {
                ftpStream.CopyTo(fileStream);
            }
        }



        private void ftpUpload(string sourceFile, string destinationFile)
        {

        
            FileInfo objFile = new FileInfo(sourceFile);
            FtpWebRequest objFTPRequest;
            objFTPRequest = (FtpWebRequest)FtpWebRequest.Create(new Uri(destinationFile));
            objFTPRequest.Credentials = new NetworkCredential(susername, spassword);
            // not closed after a command is executed.
            objFTPRequest.KeepAlive = false;

            // Set the data transfer type.
            objFTPRequest.UseBinary = true;

            // Set content length
            objFTPRequest.ContentLength = objFile.Length;

            // Set request method
            objFTPRequest.Method = WebRequestMethods.Ftp.UploadFile;

            int intBufferLength = 16 * 1024;
            byte[] objBuffer = new byte[intBufferLength];


            FileStream objFileStream = objFile.OpenRead();
            try
            {
                // Get Stream of the file
                Stream objStream = objFTPRequest.GetRequestStream();

                int len = 0;

                while ((len = objFileStream.Read(objBuffer, 0, intBufferLength)) != 0)
                {
                    // Write file Content 
                    objStream.Write(objBuffer, 0, len);

                }

                objStream.Close();
                objFileStream.Close();
            }
            catch (Exception ex)
            {
                throw ex;
            }
        }

     

        public static void DeleteFTPFile(string sFile)
        {
            FtpWebRequest clsRequest = (System.Net.FtpWebRequest)WebRequest.Create(sFile);


            clsRequest.Credentials = new NetworkCredential(susername, spassword);
            clsRequest.Method = WebRequestMethods.Ftp.DeleteFile;

            string result = string.Empty;
            FtpWebResponse response = (FtpWebResponse)clsRequest.GetResponse();
            long size = response.ContentLength;
            Stream datastream = response.GetResponseStream();
            StreamReader sr = new StreamReader(datastream);
            result = sr.ReadToEnd();
            sr.Close();
            datastream.Close();
            response.Close();
        }

        private void button11_Click(object sender, EventArgs e)
        {


            string plik = spradzPlikNaFTP(@"/SN_Tool/GZ");

            if (plik.Length > 0)
            {
                string plikKataloglokalnyplik = @"c:\DC\Gz\" + plik;
                string plikKatalogNaftp = @"ftp://172.26.59.13/SN_Tool/GZ/" + plik;
                string plikKatalogNaftpBackup = @"ftp://172.26.59.13/SN_Backup/GZ/" + plik;
                // Copy from ftp to local 

                ftpDownload(plikKatalogNaftp, plikKataloglokalnyplik);

                wczytajCeveDoBazy(plikKataloglokalnyplik);
                ftpUpload(plikKataloglokalnyplik, plikKatalogNaftpBackup);
                DeleteFTPFile(plikKatalogNaftp);
                File.Delete(plikKataloglokalnyplik);
            }
            else
            {
                MessageBox.Show("spaday ");
            }

        }

        private string spradzPlikNaFTP(string skatalog)
        {
            try
            {

                FtpWebRequest request = (FtpWebRequest)WebRequest.Create(shost + skatalog);
                request.Method = WebRequestMethods.Ftp.ListDirectory;

                request.Credentials = new NetworkCredential(susername, spassword);
                FtpWebResponse response = (FtpWebResponse)request.GetResponse();
                Stream responseStream = response.GetResponseStream();
                StreamReader reader = new StreamReader(responseStream);
                string names = reader.ReadToEnd();

                reader.Close();
                response.Close();

                return names;


            }
            catch (Exception)
            {
                throw;
            }
        }

        private void wczytajCeveDoBazy(string sciezkaplik)
        {
            //string sciezkaplik = @"c:\Praca\DC\Gorzow\serialextract-130720.txt";
            string SerialNumberCeva = "";
            string MaterialCeva = "";
            string SalesOrganisationCeva = "";
            string ShiptoNumberCeva = "";
            string WarehouseCeva;
            string WarehouseNameCeva;
            string SAPDeliveryNumberCeva = "";
            string SoldtoNumberCeva = "";
            string DateScannedCevaString;
            string ShiptoNameCeva;
            string ShiptoAddressCeva;
            string SoldtoNameCeva;
            int FileIDCeva;
            string PalletIDCeva;

            string sqlCeva;

            int w1;
            int w2;
            int w3;
            int w4;
            int w5;
            int w6;
            int dlugoscLini;
    
            SqlConnection cnnTarget;
            cnnTarget = new SqlConnection(connetionStringTarget);
            SqlDataAdapter adapterTarget = new SqlDataAdapter();
            cnnTarget.Open();
            int licznik = 0;

            foreach (string line in File.ReadLines(sciezkaplik, Encoding.UTF8))
            {
                licznik++;
                if (licznik > 1)
                {
                    w1 = line.Trim().IndexOf("\t", 1);
                    w2 = line.Trim().IndexOf("\t", w1 + 1);
                    w3 = line.Trim().IndexOf("\t", w2 + 1);
                    w4 = line.Trim().IndexOf("\t", w3 + 1);
                    w5 = line.Trim().IndexOf("\t", w4 + 1);
                    w6 = line.Trim().IndexOf("\t", w5 + 1);
                    dlugoscLini = line.Trim().Length;


                    SerialNumberCeva = line.Trim().Substring(w4 + 1, w1 - 1).Trim();
                    MaterialCeva = line.Trim().Substring(w5 + 1, dlugoscLini - w5 - 1).Trim();
                    SalesOrganisationCeva = line.Trim().Substring(0, 4).Trim();
                    ShiptoNumberCeva = line.Trim().Substring(w2 + 1, w3 - w2).Trim();
                    SAPDeliveryNumberCeva = line.Trim().Substring(5, w1 - 5).Trim();
                    DateScannedCevaString = line.Trim().Substring(w3 + 1, w4 - w3).Trim();
                    ShiptoNameCeva = "";
                    ShiptoAddressCeva = "";
                    SoldtoNumberCeva = "";
                    WarehouseCeva = "TPV_Gorzow";
                    WarehouseNameCeva = "TPV Displays Polska Sp. z o.o.";
                    SoldtoNameCeva = "";
                    FileIDCeva = 1;
                    PalletIDCeva = "";

                    sqlCeva = $" insert into SN_JG (SerialNumber, Material, DateScanned , SalesOrganisation , ShiptoNumber , ShiptoName , ShiptoAddress , Warehouse , WarehouseName , SAPDeliveryNumber , SoldtoNumber,  SoldtoName , FileID, PalletID ) " +
                       $"values (  \'  { SerialNumberCeva } \', \'  {  MaterialCeva  }  \'  ,  \' {  DateScannedCevaString  } \'    , \'  {  SalesOrganisationCeva  }  \'  , \'  {  ShiptoNumberCeva  }  \' ," +
                       $" \' { ShiptoNameCeva  }  \',  \' {  ShiptoAddressCeva  }  \' ,  \'  {  WarehouseCeva  }  \' ,  \'  {  WarehouseNameCeva   }  \' ,  \'  {  SAPDeliveryNumberCeva  }  \'," +
                       $" \'  {  ShiptoNumberCeva  }  \',  \'  {  SoldtoNameCeva  }  \',  \'  {  FileIDCeva  }  \' ,  \'  {  PalletIDCeva  }  \') ";

                    adapterTarget.InsertCommand = new SqlCommand(sqlCeva, cnnTarget);
                    adapterTarget.InsertCommand.ExecuteNonQuery();
                    Console.WriteLine(sqlCeva);
                }
            }
        }

        private void wczytajVenloDoBazy(string sciezkaplik)
        {

            string SerialNumberCeva = "";
            string MaterialCeva = "";
            string SalesOrganisationCeva = "";
            string ShiptoNumberCeva = "";
            string WarehouseCeva;
            string WarehouseNameCeva;
            string SAPDeliveryNumberCeva = "";
            string SoldtoNumberCeva = "";
            string DateScannedCevaString;
            string ShiptoNameCeva;
            string ShiptoAddressCeva;
            string SoldtoNameCeva;
            int FileIDCeva;
            string PalletIDCeva;
            int licznik = 0;

            string sqlCeva;

            int w1;
            int w2;
            int w3;
            int w4;
            int w5;
            int w6;
            int dlugoscLini;

            SqlConnection cnnTarget;
            cnnTarget = new SqlConnection(connetionStringTarget);
            SqlDataAdapter adapterTarget = new SqlDataAdapter();
            cnnTarget.Open();

            foreach (string line in File.ReadLines(sciezkaplik, Encoding.UTF8))
            {
                licznik++;
                if (licznik > 1)
                {
                    w1 = line.Trim().IndexOf("\t", 1);
                    w2 = line.Trim().IndexOf("\t", w1 + 1);
                    w3 = line.Trim().IndexOf("\t", w2 + 1);
                    w4 = line.Trim().IndexOf("\t", w3 + 1);
                    w5 = line.Trim().IndexOf("\t", w4 + 1);
                    w6 = line.Trim().IndexOf("\t", w5 + 1);
                    dlugoscLini = line.Trim().Length;

                    if (w6 == -1)
                    {
                        MaterialCeva = line.Trim().Substring(w5 + 1, dlugoscLini - w5 - 1).Trim();
                    }
                    else
                    {
                        MaterialCeva = line.Trim().Substring(w6, dlugoscLini - w6).Trim();
                    }

                    SerialNumberCeva = line.Trim().Substring(w4 + 1, w1 - 1).Trim();

                    SalesOrganisationCeva = line.Trim().Substring(0, 4).Trim();
                    ShiptoNumberCeva = line.Trim().Substring(w2 + 1, w3 - w2).Trim();
                    SAPDeliveryNumberCeva = line.Trim().Substring(5, w1 - 5).Trim();
                    DateScannedCevaString = line.Trim().Substring(w3 + 1, w4 - w3).Trim();
                    ShiptoNameCeva = "";
                    ShiptoAddressCeva = "";
                    SoldtoNumberCeva = "";
                    WarehouseCeva = "Venlo";
                    WarehouseNameCeva = "Venlo";
                    SoldtoNameCeva = "";
                    FileIDCeva = 2;
                    PalletIDCeva = "";

                    sqlCeva = $" insert into SN_JG (SerialNumber, Material, DateScanned , SalesOrganisation , ShiptoNumber , ShiptoName , ShiptoAddress , Warehouse , WarehouseName , SAPDeliveryNumber , SoldtoNumber,  SoldtoName , FileID, PalletID ) " +
                       $"values (  \'  { SerialNumberCeva } \', \'  {  MaterialCeva  }  \'  ,  \' {  DateScannedCevaString  } \'    , \'  {  SalesOrganisationCeva  }  \'  , \'  {  ShiptoNumberCeva  }  \' ," +
                       $" \' { ShiptoNameCeva  }  \',  \' {  ShiptoAddressCeva  }  \' ,  \'  {  WarehouseCeva  }  \' ,  \'  {  WarehouseNameCeva   }  \' ,  \'  {  SAPDeliveryNumberCeva  }  \'," +
                       $" \'  {  ShiptoNumberCeva  }  \',  \'  {  SoldtoNameCeva  }  \',  \'  {  FileIDCeva  }  \' ,  \'  {  PalletIDCeva  }  \') ";

                    adapterTarget.InsertCommand = new SqlCommand(sqlCeva, cnnTarget);
                    adapterTarget.InsertCommand.ExecuteNonQuery();
                    Console.WriteLine(sqlCeva);
                }
            }
        }

        private void button12_Click(object sender, EventArgs e)
        {
            string plik = spradzPlikNaFTP(@"/SN_Tool/Venlo");
            if (plik.Length > 0)
            {
                string plikKataloglokalnyplik = @"c:\DC\Venlo\" + plik;
                string plikKatalogNaftp = @"ftp://172.26.59.13/SN_Tool/Venlo/" + plik;
                string plikKatalogNaftpBackup = @"ftp://172.26.59.13/SN_Backup/Venlo/" + plik;
                ftpDownload(plikKatalogNaftp, plikKataloglokalnyplik);
                wczytajVenloDoBazy(plikKataloglokalnyplik);
                ftpUpload(plikKataloglokalnyplik, plikKatalogNaftpBackup);
                DeleteFTPFile(plikKatalogNaftp);
                File.Delete(plikKataloglokalnyplik);

            }
            else
            {

            }
        }


        private void wczytajDaganzoDoBazy(string sciezkaplik)
        {
            Excel.Application xlApp = new Excel.Application();
            Excel.Workbook xlWorkbook = xlApp.Workbooks.Open(sciezkaplik);
            Excel._Worksheet xlWorksheet = xlWorkbook.Sheets[1];
            Excel.Range xlRange = xlWorksheet.UsedRange;

            int rowCount = xlRange.Rows.Count;
            int colCount = xlRange.Columns.Count;

            #region zmienne do budowy select Daganzo
            string SerialNumberDaganzo;
            DateTime DateScannedDaganzo;

            string MaterialDaganzo;
            string SAPDeliveryNumberDaganzo;
            string ShiptoNameDaganzo;

            // do wyjasnienia 
            string SalesOrganisationDaganzo;
            string ShiptoNumberDaganzo;
            string ShiptoAddressDaganzo;
            string WarehouseDaganzo;
            string WarehouseNameDaganzo;
            string SoldtoNumberDaganzo;
            string SoldtoNameDaganzo;
            string FileIDDaganzo;
            string PalletIDDaganzo;
            string sqlDaganzo;

            #endregion

            // czesc wspolna dla wszystkich procedur 
            SqlConnection cnnTarget;
            cnnTarget = new SqlConnection(connetionStringTarget);
            SqlDataAdapter adapterTarget = new SqlDataAdapter();
            cnnTarget.Open();

            for (int i = 2; i <= rowCount; i++)
            {
                string dtString = xlRange.Cells[i, 7].Value2.ToString();
                DateScannedDaganzo = DateTime.Parse(ConvertToDateTime(dtString));

                string DateScannedDaganzoString = DateScannedDaganzo.ToString("M/dd/yyyy", CultureInfo.InvariantCulture);
                if (String.IsNullOrEmpty(xlRange.Cells[i, 5].Value2))
                    SerialNumberDaganzo = "";
                else SerialNumberDaganzo = xlRange.Cells[i, 5].Value2.ToString();
                MaterialDaganzo = xlRange.Cells[i, 2].Value2.ToString();
                ShiptoNameDaganzo = xlRange.Cells[i, 9].Value2.ToString();
                SAPDeliveryNumberDaganzo = xlRange.Cells[i, 1].Value2.ToString();
                SalesOrganisationDaganzo = "";
                ShiptoNumberDaganzo = "";
                ShiptoNameDaganzo = "";
                ShiptoAddressDaganzo = "";
                WarehouseDaganzo = "";
                WarehouseNameDaganzo = "";
                SoldtoNumberDaganzo = "";
                SoldtoNameDaganzo = "";
                FileIDDaganzo = "";
                PalletIDDaganzo = "";

                sqlDaganzo = $" insert into SN_JG (SerialNumber, Material, DateScanned , SalesOrganisation , ShiptoNumber , ShiptoName , ShiptoAddress , Warehouse , WarehouseName , SAPDeliveryNumber , SoldtoNumber,  SoldtoName , FileID, PalletID ) " +
                      $"values (  \'  { SerialNumberDaganzo } \', \'  {  MaterialDaganzo  }  \'  ,  \'{DateScannedDaganzoString}\'    , \'  {  SalesOrganisationDaganzo  }  \'  , \'  {  ShiptoNumberDaganzo  }  \' ," +
                      $" \' { ShiptoNameDaganzo  }  \',  \' {  ShiptoAddressDaganzo  }  \' ,  \'  {  WarehouseDaganzo  }  \' ,  \'  {  WarehouseNameDaganzo   }  \' ,  \'  {  SAPDeliveryNumberDaganzo  }  \'," +
                      $" \'  {  ShiptoNumberDaganzo  }  \',  \'  {  SoldtoNameDaganzo  }  \',  \'  {  FileIDDaganzo  }  \' ,  \'  {  PalletIDDaganzo  }  \') ";

                adapterTarget.InsertCommand = new SqlCommand(sqlDaganzo, cnnTarget);
                adapterTarget.InsertCommand.ExecuteNonQuery();

            }

            MessageBox.Show("Finished OK ");


        }

        private void button6_Click(object sender, EventArgs e)
        {
            string plik = spradzPlikNaFTP(@"/SN_Tool/Daganzo");
            if (plik.Length > 0)
            {
                string plikKataloglokalnyplik = @"c:\DC\Daganzo\" + plik;
                string plikKatalogNaftp = @"ftp://172.26.59.13/SN_Tool/Daganzo/" + plik;
                string plikKatalogNaftpBackup = @"ftp://172.26.59.13/SN_Backup/Daganzo/" + plik;
                ftpDownload(plikKatalogNaftp, plikKataloglokalnyplik);
                wczytajVenloDoBazy(plikKataloglokalnyplik);
                ftpUpload(plikKataloglokalnyplik, plikKatalogNaftpBackup);
                DeleteFTPFile(plikKatalogNaftp);
                File.Delete(plikKataloglokalnyplik);

            }
            else
            {

            }
        }

        private void button8_Click(object sender, EventArgs e)
        {
            string plik = spradzPlikNaFTP(@"/SN_Tool/Batta");
            if (plik.Length > 0)
            {
                string plikKataloglokalnyplik = @"c:\DC\Batta\" + plik;
                string plikKatalogNaftp = @"ftp://172.26.59.13/SN_Tool/Batta/" + plik;
                string plikKatalogNaftpBackup = @"ftp://172.26.59.13/SN_Backup/Batta/" + plik;
                ftpDownload(plikKatalogNaftp, plikKataloglokalnyplik);
                wczytajVenloDoBazy(plikKataloglokalnyplik);
                ftpUpload(plikKataloglokalnyplik, plikKatalogNaftpBackup);
                DeleteFTPFile(plikKatalogNaftp);
                File.Delete(plikKataloglokalnyplik);
            }
            else
            {

            }
        }

        private void wczytajBattaDoBazy(string sciezkaplik)
        {
            Excel.Application xlApp = new Excel.Application();
            Excel.Workbook xlWorkbook = xlApp.Workbooks.Open(sciezkaplik);
            Excel._Worksheet xlWorksheet = xlWorkbook.Sheets[1];
            Excel.Range xlRange = xlWorksheet.UsedRange;

            int rowCount = xlRange.Rows.Count;
            int colCount = xlRange.Columns.Count;

            #region zmienne do budowy select Bata
            string SerialNumberBatta;
            DateTime DateScannedBatta;

            string MaterialBatta;
            string SAPDeliveryNumberBatta;
            string ShiptoNameBatta;

            // do wyjasnienia 
            string SalesOrganisationBatta;
            string ShiptoNumberBatta;
            string ShiptoAddressBatta;
            string WarehouseBatta;
            string WarehouseNameBatta;
            string SoldtoNumberBatta;
            string SoldtoNameBatta;
            string FileIDBatta;
            string PalletIDBatta;
            string sqlBatta;

            #endregion

            SqlConnection cnnTarget;         
            cnnTarget = new SqlConnection(connetionStringTarget);
            SqlDataAdapter adapterTarget = new SqlDataAdapter();
            cnnTarget.Open();

            for (int i = 2; i <= rowCount; i++)
            {
                string dtString = xlRange.Cells[i, 10].Value2.ToString();
                DateScannedBatta = DateTime.Parse(ConvertToDateTime(dtString));

                string DateScannedBattaString = DateScannedBatta.ToString("M/dd/yyyy", CultureInfo.InvariantCulture);
                if (String.IsNullOrEmpty(xlRange.Cells[i, 7].Value2))
                    SerialNumberBatta = "";
                else SerialNumberBatta = xlRange.Cells[i, 7].Value2.ToString();
                MaterialBatta = xlRange.Cells[i, 5].Value2.ToString();
                // ?? ShiptoNameBatta = xlRange.Cells[i, 9].Value2.ToString();
                SAPDeliveryNumberBatta = xlRange.Cells[i, 21].Value2.ToString().Substring(5, 10);
                SalesOrganisationBatta = xlRange.Cells[i, 21].Value2.ToString().Substring(0, 4);
                ShiptoNumberBatta = "";
                ShiptoNameBatta = "";
                ShiptoAddressBatta = "";
                WarehouseBatta = "";
                WarehouseNameBatta = "";
                SoldtoNumberBatta = "";
                SoldtoNameBatta = "";
                FileIDBatta = "";
                PalletIDBatta = "";

                sqlBatta = $" insert into SN_JG (SerialNumber, Material, DateScanned , SalesOrganisation , ShiptoNumber , ShiptoName , ShiptoAddress , Warehouse , WarehouseName , SAPDeliveryNumber , SoldtoNumber,  SoldtoName , FileID, PalletID ) " +
                      $"values (  \'  { SerialNumberBatta } \', \'  {  MaterialBatta  }  \'  ,  \'{DateScannedBattaString}\'    , \'  {  SalesOrganisationBatta  }  \'  , \'  {  ShiptoNumberBatta  }  \' ," +
                      $" \' { ShiptoNameBatta  }  \',  \' {  ShiptoAddressBatta  }  \' ,  \'  {  WarehouseBatta  }  \' ,  \'  {  WarehouseNameBatta   }  \' ,  \'  {  SAPDeliveryNumberBatta  }  \'," +
                      $" \'  {  ShiptoNumberBatta  }  \',  \'  {  SoldtoNameBatta  }  \',  \'  {  FileIDBatta  }  \' ,  \'  {  PalletIDBatta  }  \') ";

                adapterTarget.InsertCommand = new SqlCommand(sqlBatta, cnnTarget);
                adapterTarget.InsertCommand.ExecuteNonQuery();


            }
            MessageBox.Show("Skonczone OK");
        }

        private void button9_Click(object sender, EventArgs e)
        {
            wpiszDoLoga("Operacja jeden");
        }

        private void wpiszDoLoga(string trescOperacji)
        {
          //  DateTime steraz = DateTime.Now;

            string steraz = DateTime.Now.ToString("M/dd/yyyy HH:mm:ss", CultureInfo.InvariantCulture);

            SqlConnection cnnTarget;
            cnnTarget = new SqlConnection(connetionStringTarget);
            SqlDataAdapter adapterTarget = new SqlDataAdapter();
            cnnTarget.Open();
            string sqlImportLog = $"INSERT INTO ImportLog (DataLog, Operation) VALUES( \'{steraz}\' , \'{trescOperacji}\' ) ";


            adapterTarget.InsertCommand = new SqlCommand(sqlImportLog, cnnTarget);
            adapterTarget.InsertCommand.ExecuteNonQuery();
            cnnTarget.Close();
        }
    }
}
