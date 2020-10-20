using NPOI.HSSF.UserModel;
using NPOI.SS.UserModel;
using NPOI.XSSF.UserModel;
using Org.BouncyCastle.Asn1;
using Renci.SshNet;
using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Data.SqlClient;
using System.Diagnostics;
using System.Drawing;
using System.Globalization;
using System.IO;
using System.IO.Compression;
using System.Linq;
using System.Net;
using System.Text;
using System.Threading;
using System.Threading.Tasks;
using System.Windows.Forms;
using WinSCP;
using Timer = System.Windows.Forms.Timer;

using Excel = Microsoft.Office.Interop.Excel;


namespace WindowsFormsApp3
{


    public partial class Form1 : Form
    {
        public Form1()
        {

            InitializeComponent();
        }


        private static FileStream fs = new FileStream(@"c:\temp\mcb.txt", FileMode.OpenOrCreate, FileAccess.Write);
        private static StreamWriter m_streamWriter = new StreamWriter(fs);
        private const string Caption = "Interval elapsed.  Continue running?";
        public Timer Timer = new Timer();
        int lastHour = DateTime.Now.Hour;

        // public const string RodzajFTP = "sFtp";

        public const string RodzajFTP = "Ftp";
        public const string susername = "SN_Tool";
        public const string spassword = "46qD8N4w";
        System.Windows.Forms.Timer t = new System.Windows.Forms.Timer();

        public string shost = "ftp://172.26.59.13";
        // public string connetionStringTarget = @"Data Source=DESKTOP-M9PRPPC\MSSQLSERVER01;Initial Catalog=TMS;Integrated Security=True;";
        public string connetionStringTarget = @"Data Source= 172.17.80.141;Initial Catalog=SN_Tool;User Id=sn_tool;Password=F7E@Ln!j_viFoSW*R6bi";

        public string katalogprogramu = "";
        public static bool leci = false;


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
            this.label5.Text = connetionStringTarget;
            // wczytajPlikizFtp();

        }

        private void button2_Click(object sender, EventArgs e)
        {

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

        private void ftpWinScpDownload(string sourceFile, string destinationFile)
        {
            SessionOptions FtpsessionOptions = new SessionOptions
            {
                Protocol = Protocol.Ftp,
                HostName = "172.26.59.13",
                UserName = "SN_Tool",
                Password = "46qD8N4w",
                FtpSecure = FtpSecure.Explicit,
                TlsHostCertificateFingerprint = "1e:91:90:86:47:16:96:7d:12:c4:ac:3f:0f:04:98:c2:3c:78:a5:0c",
            };

            SessionOptions sFtpsessionOptions = new SessionOptions
            {
                Protocol = Protocol.Sftp,
                HostName = "128.127.93.50",
                UserName = "SN_Tool",
                Password = "46qD8N4w",
                SshHostKeyFingerprint = "ssh-rsa 2048 GYol3dF9i8Br6BvsU469/Lx3qdUA18gNwIuaN/DORdE=",
            };
            using (WinSCP.Session session = new WinSCP.Session())
            {
                // Connect
                if (RodzajFTP == "sFtp")
                {
                    session.Open(sFtpsessionOptions);
                }

                if (RodzajFTP == "Ftp")
                {
                    session.Open(FtpsessionOptions);
                }
                // Upload files
                TransferOptions transferOptions = new TransferOptions();
                transferOptions.TransferMode = TransferMode.Binary;
                TransferOperationResult transferResult;
                transferResult =
                    session.GetFiles(sourceFile, destinationFile, false, transferOptions);
                // Throw on any error
                transferResult.Check();
                // Print results
                foreach (TransferEventArgs transfer in transferResult.Transfers)
                {
                    Console.WriteLine("Upload of {0} succeeded", transfer.FileName);
                }
            }

        }

        private void ftpWinScpUpload(string sourceFile, string destinationFile)
        {
            SessionOptions FtpsessionOptions = new SessionOptions
            {

                Protocol = Protocol.Ftp,
                HostName = "172.26.59.13",
                UserName = "SN_Tool",
                Password = "46qD8N4w",
                FtpSecure = FtpSecure.Explicit,
                TlsHostCertificateFingerprint = "1e:91:90:86:47:16:96:7d:12:c4:ac:3f:0f:04:98:c2:3c:78:a5:0c",

            };


            SessionOptions sFtpsessionOptions = new SessionOptions
            {

                Protocol = Protocol.Sftp,
                HostName = "128.127.93.50",
                UserName = "SN_Tool",
                Password = "46qD8N4w",
                SshHostKeyFingerprint = "ssh-rsa 2048 GYol3dF9i8Br6BvsU469/Lx3qdUA18gNwIuaN/DORdE=",

            };
            using (WinSCP.Session session = new WinSCP.Session())
            {
                // Connect
                if (RodzajFTP == "sFtp")
                {
                    session.Open(sFtpsessionOptions);
                }


                if (RodzajFTP == "Ftp")
                {
                    session.Open(FtpsessionOptions);
                }
                // Upload files
                TransferOptions transferOptions = new TransferOptions();
                transferOptions.TransferMode = TransferMode.Binary;
                TransferOperationResult transferResult;
                transferResult =
                    session.PutFiles(sourceFile, destinationFile, false, transferOptions);
                // Throw on any error
                transferResult.Check();
                // Print results
                foreach (TransferEventArgs transfer in transferResult.Transfers)
                {
                    Console.WriteLine("Upload of {0} succeeded", transfer.FileName);
                }
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


            string plik = spradzPlikNaFTP(@"/SN_Tool/GZ").Trim();



            if (plik.Length > 0)
            {
                string plikKataloglokalnyplik = @"c:\DC\Gz\" + plik;
                string plikKatalogNaftp = @"ftp://172.26.59.13/SN_Tool/GZ/" + plik;
                string plikKatalogNaftpBackup = @"ftp://172.26.59.13/SN_Backup/GZ/" + plik;
                // Copy from ftp to local 

                ftpDownload(plikKatalogNaftp, plikKataloglokalnyplik);

                wpiszDoLoga("Skopiowano z ftp ", "CEVA GZ", plik, "TV");
                Wait(3000);
                zamknijProcessExcela();
                wczytajCeveDoBazy(plikKataloglokalnyplik);
                wpiszDoLoga("Wczytano do bazy ", "CEVA GZ", plik, "TV");
                Wait(3000);
                zamknijProcessExcela();
                ftpUpload(plikKataloglokalnyplik, plikKatalogNaftpBackup);
                wpiszDoLoga("Skopiowano do Backup ftp ", "CEVA GZ", plik, "TV");
                DeleteFTPFile(plikKatalogNaftp);

                wpiszDoLoga("Skasowano ftp ", "CEVA GZ", plik, "TV");
                File.Delete(plikKataloglokalnyplik);
                wpiszDoLoga("Skasowano lokalnie ", "CEVA GZ", plik, "TV");
                MessageBox.Show(" skonczone OK ");
            }
            else
            {
                MessageBox.Show("Brak pliku na ftp ");
            }

        }

        private string sprawdzPlikWinScpFtp(string skatalog)
        {
            string sFilename = "";


            SessionOptions FtpsessionOptions = new SessionOptions
            {

                Protocol = Protocol.Ftp,
                HostName = "172.26.59.13",
                UserName = "SN_Tool",
                Password = "46qD8N4w",
                FtpSecure = FtpSecure.Explicit,
                TlsHostCertificateFingerprint = "1e:91:90:86:47:16:96:7d:12:c4:ac:3f:0f:04:98:c2:3c:78:a5:0c",

            };


            SessionOptions sFtpsessionOptions = new SessionOptions
            {

                Protocol = Protocol.Sftp,
                HostName = "128.127.93.50",
                UserName = "SN_Tool",
                Password = "46qD8N4w",
                SshHostKeyFingerprint = "ssh-rsa 2048 GYol3dF9i8Br6BvsU469/Lx3qdUA18gNwIuaN/DORdE=",

            };


            ////////////////////////////////////




            using (WinSCP.Session session = new WinSCP.Session())
            {
                // Connect


                if (RodzajFTP == "sFtp")
                {
                    session.Open(sFtpsessionOptions);
                }


                if (RodzajFTP == "Ftp")
                {
                    session.Open(FtpsessionOptions);
                }

                IEnumerable<RemoteFileInfo> fileInfos =
                        session.EnumerateRemoteFiles(
                            skatalog, null,
                            EnumerationOptions.EnumerateDirectories |
                                EnumerationOptions.AllDirectories);
                foreach (RemoteFileInfo fileInfo in fileInfos)
                {

                    sFilename = fileInfo.FullName.ToString();
                }

                return sFilename;

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
            string DateScannedCevaStringRaw;
            string plikDoZapisu = sciezkaplik.Trim();

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
                    DateScannedCevaStringRaw = line.Trim().Substring(w3 + 1, w4 - w3).Trim();
                    DateScannedCevaString = DateScannedCevaStringRaw.Substring(3, 2) + "/" + DateScannedCevaStringRaw.Substring(0, 2) + "/" + DateScannedCevaStringRaw.Substring(6, 4);
                    ShiptoNameCeva = "";
                    ShiptoAddressCeva = "";
                    SoldtoNumberCeva = "";
                    WarehouseCeva = "TPV_Gorzow";
                    WarehouseNameCeva = "TPV Displays Polska Sp. z o.o.";
                    SoldtoNameCeva = "";
                    FileIDCeva = 1;
                    PalletIDCeva = plikDoZapisu;

                    sqlCeva = $" insert into SN_JG (SerialNumber, Material, DateScanned , SalesOrganisation , ShiptoNumber , ShiptoName , ShiptoAddress , Warehouse , WarehouseName , SAPDeliveryNumber , SoldtoNumber,  SoldtoName , FileID, PalletID ) " +
                       $"values (  \'  { SerialNumberCeva } \', \'  {  MaterialCeva  }  \'  ,  \' {  DateScannedCevaString  } \'    , \'  {  SalesOrganisationCeva  }  \'  , \'  {  ShiptoNumberCeva  }  \' ," +
                       $" \' { ShiptoNameCeva  }  \',  \' {  ShiptoAddressCeva  }  \' ,  \'  {  WarehouseCeva  }  \' ,  \'  {  WarehouseNameCeva   }  \' ,  \'  {  SAPDeliveryNumberCeva  }  \'," +
                       $" \'  {  ShiptoNumberCeva  }  \',  \'  {  SoldtoNameCeva  }  \',  \'  {  FileIDCeva  }  \' ,  \'  {  PalletIDCeva  }  \') ";

                    adapterTarget.InsertCommand = new SqlCommand(sqlCeva, cnnTarget);
                    adapterTarget.InsertCommand.ExecuteNonQuery();
                    Console.WriteLine(licznik);
                }
            }
        }

        private void wczytajVenloDoBazy(string sciezkaplik)
        {
            string SerialNumberVenlo = "";
            string MaterialVenlo = "";
            string SalesOrganisationVenlo = "";
            string ShiptoNumberVenlo = "";
            string WarehouseVenlo;
            string WarehouseNameVenlo;
            string SAPDeliveryNumberVenlo = "";
            string SoldtoNumberVenlo = "";
            string DateScannedVenloString;
            string ShiptoNameVenlo;
            string ShiptoAddressVenlo;
            string SoldtoNameVenlo;
            int FileIDVenlo;
            string PalletIDVenlo;
            int licznik = 0;
            string DateScannedVenloStringRaw;
            string plikDoZapisu = sciezkaplik.Trim();


            string sqlVenlo;

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
                if (licznik > 2)
                {
                    w1 = line.Trim().IndexOf(";", 1);
                    w2 = line.Trim().IndexOf(";", w1 + 1);
                    w3 = line.Trim().IndexOf(";", w2 + 1);
                    w4 = line.Trim().IndexOf(";", w3 + 1);
                    w5 = line.Trim().IndexOf(";", w4 + 1);
                    w6 = line.Trim().IndexOf(";", w5 + 1);
                    dlugoscLini = line.Trim().Length;

                    SerialNumberVenlo = line.Trim().Substring(w5 + 1, dlugoscLini - w5 - 1).Trim();

                    SalesOrganisationVenlo = line.Trim().Substring(0, 4).Trim();

                    SAPDeliveryNumberVenlo = line.Trim().Substring(w2 + 1, w3 - w2 - 1).Trim();
                    DateScannedVenloString = line.Trim().Substring(w3 + 1, w4 - w3).Trim();
                    MaterialVenlo = line.Trim().Substring(w4 + 1, w5 - w4 - 1).Trim();
                    DateScannedVenloStringRaw = line.Trim().Substring(w1 + 1, w2 - w1).Trim();
                    DateScannedVenloString = DateScannedVenloStringRaw.Substring(4, 2) + "/" + DateScannedVenloStringRaw.Substring(6, 2) + "/" + DateScannedVenloStringRaw.Substring(0, 4);
                    ShiptoNameVenlo = "";
                    ShiptoAddressVenlo = "";
                    SoldtoNumberVenlo = "";
                    WarehouseVenlo = "Venlo";
                    WarehouseNameVenlo = "Venlo";
                    SoldtoNameVenlo = "";
                    FileIDVenlo = 2;

                    if (plikDoZapisu.Trim().Length > 46)
                    { PalletIDVenlo = plikDoZapisu.Trim().Substring(0, 46); }
                    else
                    {
                        PalletIDVenlo = plikDoZapisu.Trim();
                    }

                    sqlVenlo = $" insert into SN_JG (SerialNumber, Material, DateScanned , SalesOrganisation , ShiptoNumber , ShiptoName , ShiptoAddress , Warehouse , WarehouseName , SAPDeliveryNumber , SoldtoNumber,  SoldtoName , FileID, PalletID ) " +
                       $"values (  \'  { SerialNumberVenlo } \', \'  {  MaterialVenlo  }  \'  ,  \' {  DateScannedVenloString  } \'    , \'  {  SalesOrganisationVenlo  }  \'  , \'  {  ShiptoNumberVenlo  }  \' ," +
                       $" \' { ShiptoNameVenlo  }  \',  \' {  ShiptoAddressVenlo  }  \' ,  \'  {  WarehouseVenlo  }  \' ,  \'  {  WarehouseNameVenlo   }  \' ,  \'  {  SAPDeliveryNumberVenlo  }  \'," +
                       $" \'  {  ShiptoNumberVenlo  }  \',  \'  {  SoldtoNameVenlo  }  \',  \'  {  FileIDVenlo  }  \' ,  \'  {  PalletIDVenlo  }  \') ";

                    adapterTarget.InsertCommand = new SqlCommand(sqlVenlo, cnnTarget);
                    adapterTarget.InsertCommand.ExecuteNonQuery();

                    Console.WriteLine(licznik.ToString());

                }
            }

        }

        private void button12_Click(object sender, EventArgs e)
        {
            string plik = spradzPlikNaFTP(@"/SN_Tool/Venlo").Trim();

            if (plik.Length > 0)
            {
                string plikKataloglokalnyplik = @"c:\DC\Venlo\" + plik;
                string plikKatalogNaftp = @"ftp://172.26.59.13/SN_Tool/Venlo/" + plik;
                string plikKatalogNaftpBackup = @"ftp://172.26.59.13/SN_Backup/Venlo/" + plik;
                ftpDownload(plikKatalogNaftp, plikKataloglokalnyplik);
                Wait(3000);
                wpiszDoLoga("Skopiowano na lokalny ", "Venlo", plik, "TV");
                zamknijProcessExcela();
                wczytajVenloDoBazy(plikKataloglokalnyplik);
                Wait(3000);
                wpiszDoLoga("Wczytano do bazy ", "Venlo", plik, "TV");
                zamknijProcessExcela();
                ftpUpload(plikKataloglokalnyplik, plikKatalogNaftpBackup);
                wpiszDoLoga("Skopiowano do backup ftp ", "Venlo", plik, "TV");
                DeleteFTPFile(plikKatalogNaftp);
                wpiszDoLoga("Skasowano na ftp ", "Venlo", plik, "TV");
                File.Delete(plikKataloglokalnyplik);
                wpiszDoLoga("Skasowano  lokalnego ", "Venlo", plik, "TV");
                MessageBox.Show(" skonczone OK ");
            }
            else
            {
                MessageBox.Show("Brak pliku na ftp ");
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
            string plikDoZapisu = sciezkaplik.Trim();

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
                SalesOrganisationDaganzo = "613E";
                ShiptoNumberDaganzo = "";
                ShiptoNameDaganzo = "";
                ShiptoAddressDaganzo = "";
                WarehouseDaganzo = "Daganzo";
                WarehouseNameDaganzo = "Daganzo";
                SoldtoNumberDaganzo = "";
                SoldtoNameDaganzo = "";
                FileIDDaganzo = "";
                PalletIDDaganzo = plikDoZapisu;

                sqlDaganzo = $" insert into SN_JG (SerialNumber, Material, DateScanned , SalesOrganisation , ShiptoNumber , ShiptoName , ShiptoAddress , Warehouse , WarehouseName , SAPDeliveryNumber , SoldtoNumber,  SoldtoName , FileID, PalletID ) " +
                      $"values (  \'  { SerialNumberDaganzo } \', \'  {  MaterialDaganzo  }  \'  ,  \'{DateScannedDaganzoString}\'    , \'  {  SalesOrganisationDaganzo  }  \'  , \'  {  ShiptoNumberDaganzo  }  \' ," +
                      $" \' { ShiptoNameDaganzo  }  \',  \' {  ShiptoAddressDaganzo  }  \' ,  \'  {  WarehouseDaganzo  }  \' ,  \'  {  WarehouseNameDaganzo   }  \' ,  \'  {  SAPDeliveryNumberDaganzo  }  \'," +
                      $" \'  {  ShiptoNumberDaganzo  }  \',  \'  {  SoldtoNameDaganzo  }  \',  \'  {  FileIDDaganzo  }  \' ,  \'  {  PalletIDDaganzo  }  \') ";

                adapterTarget.InsertCommand = new SqlCommand(sqlDaganzo, cnnTarget);
                adapterTarget.InsertCommand.ExecuteNonQuery();

                Console.WriteLine(i);

            }

            xlWorkbook.Close(0);
            xlApp.Quit();


        }

        private void button6_Click(object sender, EventArgs e)
        {
            string plik = spradzPlikNaFTP(@"/SN_Tool/Daganzo").Trim();
            if (plik.Length > 0)
            {
                string plikKataloglokalnyplik = @"c:\DC\Daganzo\" + plik;
                string plikKatalogNaftp = @"ftp://172.26.59.13/SN_Tool/Daganzo/" + plik;
                string plikKatalogNaftpBackup = @"ftp://172.26.59.13/SN_Backup/Daganzo/" + plik;

                ftpDownload(plikKatalogNaftp, plikKataloglokalnyplik);
                Wait(3000);
                wpiszDoLoga("Skopiowano na lokalny ", "Daganzo", plik, "TV");
                zamknijProcessExcela();
                wczytajDaganzoDoBazy(plikKataloglokalnyplik);
                Wait(3000);
                wpiszDoLoga("Wczytano do bazy  ", "Daganzo", plik, "TV");
                zamknijProcessExcela();
                ftpUpload(plikKataloglokalnyplik, plikKatalogNaftpBackup);
                wpiszDoLoga("Skopiowano do backup ftp ", "Daganzo", plik, "TV");
                DeleteFTPFile(plikKatalogNaftp);
                wpiszDoLoga("Skopiowano ftp ", "Daganzo", plik, "TV");
                File.Delete(plikKataloglokalnyplik);
                wpiszDoLoga("Skasowano lokalny ", "Venlo", plik, "TV");
                MessageBox.Show(" skonczone OK ");
            }
            else
            {
                MessageBox.Show("Brak pliku na ftp ");
            }
        }

        private void button8_Click(object sender, EventArgs e)
        {
            string plik = spradzPlikNaFTP(@"/SN_Tool/Batta").Trim();

            //  string plik = "ShippedSNsWK28.xlsx";
            if (plik.Length > 0)
            {
                string plikKataloglokalnyplik = @"c:\DC\Batta\" + plik;
                string plikKatalogNaftp = @"ftp://172.26.59.13/SN_Tool/Batta/" + plik;
                string plikKatalogNaftpBackup = @"ftp://172.26.59.13/SN_Backup/Batta/" + plik;
                ftpDownload(plikKatalogNaftp, plikKataloglokalnyplik);
                Wait(3000);
                wpiszDoLoga("Skopiowano na lokalny ", "Batta", plik, "TV");
                zamknijProcessExcela();
                wczytajBattaDoBazy(plikKataloglokalnyplik);
                Wait(3000);
                zamknijProcessExcela();
                wpiszDoLoga("Wczytano do bazy ", "Batta", plik, "TV");
                ftpUpload(plikKataloglokalnyplik, plikKatalogNaftpBackup);
                wpiszDoLoga("Skopiowano do backup ftp ", "Batta", plik, "TV");
                DeleteFTPFile(plikKatalogNaftp);
                wpiszDoLoga("Skasowano ftp ", "Batta", plik, "TV");
                File.Delete(plikKataloglokalnyplik);
                wpiszDoLoga("Skasowano lokalny ", "Venlo", plik, "TV");
                MessageBox.Show(" skonczone OK ");
            }
            else
            {
                MessageBox.Show("Brak pliku na ftp ");
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
            string plikDoZapisu = sciezkaplik.Trim();

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
                WarehouseBatta = "Batta";
                WarehouseNameBatta = "Batta";
                SoldtoNumberBatta = "";
                SoldtoNameBatta = "";
                FileIDBatta = "";
                PalletIDBatta = plikDoZapisu;

                sqlBatta = $" insert into SN_JG (SerialNumber, Material, DateScanned , SalesOrganisation , ShiptoNumber , ShiptoName , ShiptoAddress , Warehouse , WarehouseName , SAPDeliveryNumber , SoldtoNumber,  SoldtoName , FileID, PalletID ) " +
                      $"values (  \'  { SerialNumberBatta } \', \'  {  MaterialBatta  }  \'  ,  \'{DateScannedBattaString}\'    , \'  {  SalesOrganisationBatta  }  \'  , \'  {  ShiptoNumberBatta  }  \' ," +
                      $" \' { ShiptoNameBatta  }  \',  \' {  ShiptoAddressBatta  }  \' ,  \'  {  WarehouseBatta  }  \' ,  \'  {  WarehouseNameBatta   }  \' ,  \'  {  SAPDeliveryNumberBatta  }  \'," +
                      $" \'  {  ShiptoNumberBatta  }  \',  \'  {  SoldtoNameBatta  }  \',  \'  {  FileIDBatta  }  \' ,  \'  {  PalletIDBatta  }  \') ";

                adapterTarget.InsertCommand = new SqlCommand(sqlBatta, cnnTarget);
                adapterTarget.InsertCommand.ExecuteNonQuery();

                Console.WriteLine(i);

            }
            xlWorkbook.Close(0);
            xlApp.Quit();
        }



        private void wpiszDoLoga(string trescOperacji, string sPlant, string sFileName, string sProductGroup)
        {
            //  DateTime steraz = DateTime.Now;

            string steraz = DateTime.Now.ToString("M/dd/yyyy HH:mm:ss", CultureInfo.InvariantCulture);

            SqlConnection cnnTarget;
            cnnTarget = new SqlConnection(connetionStringTarget);
            SqlDataAdapter adapterTarget = new SqlDataAdapter();
            cnnTarget.Open();
            string sqlImportLog = $"INSERT INTO ImportLog (DataLog, Operation, Plant , FileName , ProductGroup) VALUES( \'{steraz}\' , \'{trescOperacji}\' ,\'{sPlant}\'  ,\'{sFileName}\' ,\'{sProductGroup}\'     ) ";


            adapterTarget.InsertCommand = new SqlCommand(sqlImportLog, cnnTarget);
            adapterTarget.InsertCommand.ExecuteNonQuery();
            cnnTarget.Close();
        }

        private void button10_Click(object sender, EventArgs e)
        {

        }


        private void wczytajSAPDNDoBazy(string sciezkaplik)
        {
            string ShipToCustomerIDSAPdata;
            string ShipToNameSAPdata;
            string ShipToCountrySAPData;
            string SoldToCustomerIDSAPData;
            string SoldToNameSAPData;
            string SoldToCountrySAPData;
            string DeliveryPostDateTimeSAPData;
            string DNSAPData;
            int licznik = 0;
            string sFile = "";

            string sqlSAPDNData;

            int w1;
            int w2;
            int w3;
            int w4;
            int w5;
            int w6;
            int w7;
            int w8;
            int w9;
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
                    w1 = line.Trim().IndexOf("|", 1);
                    w2 = line.Trim().IndexOf("|", w1 + 1);
                    w3 = line.Trim().IndexOf("|", w2 + 1);
                    w4 = line.Trim().IndexOf("|", w3 + 1);
                    w5 = line.Trim().IndexOf("|", w4 + 1);
                    w6 = line.Trim().IndexOf("|", w5 + 1);
                    w7 = line.Trim().IndexOf("|", w6 + 1);
                    w8 = line.Trim().IndexOf("|", w7 + 1);
                    dlugoscLini = line.Trim().Length;

                    ShipToCustomerIDSAPdata = line.Trim().Substring(w2 + 1, w3 - w2 - 1).Trim();
                    ShipToNameSAPdata = line.Trim().Substring(w3 + 1, w4 - w3 - 1).Trim().Replace("'", " "); ;
                    ShipToCountrySAPData = line.Trim().Substring(w4 + 1, w5 - w4 - 1).Trim();
                    SoldToCustomerIDSAPData = line.Trim().Substring(w5 + 1, w6 - w5 - 1).Trim();
                    SoldToNameSAPData = line.Trim().Substring(w6 + 1, w7 - w6 - 1).Trim().Replace("'", " ");
                    SoldToCountrySAPData = line.Trim().Substring(w7 + 1, w8 - w7 - 1).Trim();
                    DeliveryPostDateTimeSAPData = line.Trim().Substring(w8 + 1, dlugoscLini - w8 - 1).Trim();

                    DNSAPData = line.Trim().Substring(w1 + 1, w2 - w1 - 1).Trim();

                    sFile = sciezkaplik.Trim();

                    sqlSAPDNData = $" insert into SAPDNData (ShipToCustomerID, ShipToName, ShipToCountry ,SoldToCustomerID  , SoldToName , SoldToCountry,DeliveryPostDateTime ,DN , FileName ) " +
                     $"values (  \'  { ShipToCustomerIDSAPdata } \', \'  {  ShipToNameSAPdata  }  \'  ,  \' {  ShipToCountrySAPData  } \'    , \'  {  SoldToCustomerIDSAPData  }  \'  , \'  {  SoldToNameSAPData  }  \' ," +
                     $" \' { SoldToCountrySAPData  }  \',  \' {  DeliveryPostDateTimeSAPData  }  \' ,  \'  {  DNSAPData  }  \',  \'  {  sFile  }  \') ";


                    adapterTarget.InsertCommand = new SqlCommand(sqlSAPDNData, cnnTarget);
                    adapterTarget.InsertCommand.ExecuteNonQuery();

                }
            }
        }
        private void button13_Click(object sender, EventArgs e)
        {
            string plik = spradzPlikNaFTP(@"/SN_Tool/SAPDNData").Trim();

            if (plik.Length > 0)
            {
                string plikKataloglokalnyplik = @"c:\DC\SAPDNData\" + plik;
                string plikKatalogNaftp = @"ftp://172.26.59.13/SN_Tool/SAPDNData/" + plik;
                string plikKatalogNaftpBackup = @"ftp://172.26.59.13/SN_Backup/SAPDNData/" + plik;

                ftpDownload(plikKatalogNaftp, plikKataloglokalnyplik);
                Wait(3000);
                zamknijProcessExcela();
                wczytajSAPDNDoBazy(plikKataloglokalnyplik);
                Wait(3000);
                zamknijProcessExcela();
                ftpUpload(plikKataloglokalnyplik, plikKatalogNaftpBackup);
                DeleteFTPFile(plikKatalogNaftp);
                File.Delete(plikKataloglokalnyplik);
                MessageBox.Show(" skonczone OK ");
            }
            else
            {
                MessageBox.Show("Brak pliku na ftp ");
            }
        }

        private void button14_Click(object sender, EventArgs e)
        {

            string path = @"c:\DC\SNoutbound.xlsx";

            try
            {
                IWorkbook workbook = null;
                FileStream fs = new FileStream(path, FileMode.Open, FileAccess.Read);
                if (path.IndexOf(".xlsx") > 0)
                {
                    workbook = new XSSFWorkbook(fs);
                }
                else if (path.IndexOf(".xls") > 0)
                {
                    workbook = new HSSFWorkbook(fs);
                }

                // ISheet sheet = workbook.CreateSheet();

                ISheet sheet = workbook.GetSheetAt(0);

                if (sheet != null)
                {
                    int rowCount = sheet.LastRowNum;

                    for (int i = 1; i <= rowCount; i++)
                    {

                        IRow curRow = sheet.GetRow(i);
                        // var cellValue = curRow.GetCell(0).StringCellValue.Trim();
                        var cellValue = curRow.GetCell(0).NumericCellValue.ToString();
                    }
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message);
            }

        }

        private void button15_Click(object sender, EventArgs e)
        {

        }


        private void WczytajMNTDoBazy(string pliklokalny)
        {
            Excel.Application xlApp = new Excel.Application();
            Excel.Workbook xlWorkbook = xlApp.Workbooks.Open(pliklokalny);
            Excel._Worksheet xlWorksheet = xlWorkbook.Sheets[1];
            Excel.Range xlRange = xlWorksheet.UsedRange;

            int rowCount = xlRange.Rows.Count;
            int colCount = xlRange.Columns.Count;

            string sGoodsIssueDate;
            string sPlant;
            string sDN;
            string sShipToName;
            string PRsShipToName;
            string sPartNumber;
            string sSerialNumber;
            string sqlMNT;
            string sFileName = pliklokalny.Trim();
            string nazwaFileName = "";

            if (sFileName.Trim().Length > 49)
            { nazwaFileName = sFileName.Substring(0, 49); }
            else { nazwaFileName = sFileName; }

            SqlConnection cnnTarget;

            cnnTarget = new SqlConnection(connetionStringTarget);
            SqlDataAdapter adapterTarget = new SqlDataAdapter();
            cnnTarget.Open();

            for (int i = 2; i <= rowCount; i++)
            {
                sGoodsIssueDate = xlRange.Cells[i, 1].Value2.ToString();
                sPlant = xlRange.Cells[i, 2].Value2.ToString().Substring(0, 4);
                sDN = xlRange.Cells[i, 2].Value2.ToString().Substring(5, 10);
                PRsShipToName = xlRange.Cells[i, 3].Value2.ToString().Replace("/", " ").Replace("'", " ");

                if (PRsShipToName.Length > 46)
                {
                    sShipToName = PRsShipToName.Substring(1, 45);
                }
                else
                {
                    sShipToName = PRsShipToName;
                }

                sPartNumber = xlRange.Cells[i, 4].Value2.ToString();
                sSerialNumber = xlRange.Cells[i, 5].Value2.ToString();


                sqlMNT = $" insert into SN_MNT_JG (GoodsIssueDate,  Plant ,DN  , ShipToName , PartNumber,SerialNumber, FileName ) " +
                      $"values (  \'  { sGoodsIssueDate } \', \'  {  sPlant  }  \'  ,  \'{sDN}\'    , \'  {  sShipToName  }  \'  , \'  {  sPartNumber.Trim()  }  \' ," +
                      $" \' { sSerialNumber.Trim()  }  \' , \' { nazwaFileName  }  \') ";

                adapterTarget.InsertCommand = new SqlCommand(sqlMNT, cnnTarget);
                adapterTarget.InsertCommand.ExecuteNonQuery();
                Console.WriteLine(i);

            }
        }

        private void button16_Click(object sender, EventArgs e)
        {
            string plik = spradzPlikNaFTP(@"/SN_Tool/MNT").Trim();


            if (plik.Length > 0)
            {
                string plikKataloglokalnyplik = @"c:\DC\MNT\" + plik;
                string plikKatalogNaftp = @"ftp://172.26.59.13/SN_Tool/MNT/" + plik;
                string plikKatalogNaftpBackup = @"ftp://172.26.59.13/SN_Backup/MNT/" + plik;

                ftpDownload(plikKatalogNaftp, plikKataloglokalnyplik);
                Wait(3000);
                zamknijProcessExcela();
                wpiszDoLoga("Skopiowano na lokalny ", "GZ", plik, "MNT");
                WczytajMNTDoBazy(plikKataloglokalnyplik);
                Wait(3000);
                zamknijProcessExcela();
                wpiszDoLoga("Wczytano do bazy ", "GZ", plik, "MNT");
                Wait(3000);
                zamknijProcessExcela();
                ftpUpload(plikKataloglokalnyplik, plikKatalogNaftpBackup);
                wpiszDoLoga("Skopiowano do backup ", "CZ", plik, "MNT");
                DeleteFTPFile(plikKatalogNaftp);
                wpiszDoLoga("Skasowano ftp ", "CZ", plik, "MNT");
                File.Delete(plikKataloglokalnyplik);
                wpiszDoLoga("Skasowano  lokalny ", "CZ", plik, "MNT");
                MessageBox.Show(" skonczone OK ");
            }
            else
            {
                MessageBox.Show("Brak pliku na ftp ");
            }

        }

        private void button7_Click(object sender, EventArgs e)
        {

        }

        private void button7_Click_1(object sender, EventArgs e)
        {
            kasujPliki(@"C:\DC\Batta");
            kasujPliki(@"C:\DC\Daganzo");
            kasujPliki(@"C:\DC\GZ");
            kasujPliki(@"C:\DC\MNT");
            kasujPliki(@"C:\DC\SAPDNData");
            kasujPliki(@"C:\DC\Venlo");

            MessageBox.Show("Pliki skasowanie !!!");

        }

        private void kasujPliki(string sciezka)
        {
            string[] plikiDoSkasowania = Directory.GetFiles(sciezka, "*.*");
            foreach (string fi in plikiDoSkasowania)
            {
                File.Delete(fi);
            }
        }

        private void kasowanieTabelTmp()
        {
            SqlConnection cnnTarget;
            string kuraMNT;
            string kuraTV;
            string kuraDN;
            cnnTarget = new SqlConnection(connetionStringTarget);
            SqlDataAdapter adapterTarget = new SqlDataAdapter();
            cnnTarget.Open();
            kuraMNT = $"delete from  SN_MNT_JG";
            kuraTV = $"delete from  SN_JG";
            kuraDN = $"delete from  SAPDNData";


            adapterTarget.InsertCommand = new SqlCommand(kuraMNT, cnnTarget);
            adapterTarget.InsertCommand.ExecuteNonQuery();

            adapterTarget.InsertCommand = new SqlCommand(kuraTV, cnnTarget);
            adapterTarget.InsertCommand.ExecuteNonQuery();

            adapterTarget.InsertCommand = new SqlCommand(kuraDN, cnnTarget);
            adapterTarget.InsertCommand.ExecuteNonQuery();

            cnnTarget.Close();
        }

        private void button17_Click(object sender, EventArgs e)
        {
            SqlConnection cnnTarget;
            string kuraMNT;
            string kuraTV;
            string kuraDN;
            cnnTarget = new SqlConnection(connetionStringTarget);
            SqlDataAdapter adapterTarget = new SqlDataAdapter();
            cnnTarget.Open();
            kuraMNT = $"delete from  SN_MNT_JG";
            kuraTV = $"delete from  SN_JG";
            kuraDN = $"delete from  SAPDNData";


            adapterTarget.InsertCommand = new SqlCommand(kuraMNT, cnnTarget);
            adapterTarget.InsertCommand.ExecuteNonQuery();

            adapterTarget.InsertCommand = new SqlCommand(kuraTV, cnnTarget);
            adapterTarget.InsertCommand.ExecuteNonQuery();

            adapterTarget.InsertCommand = new SqlCommand(kuraDN, cnnTarget);
            adapterTarget.InsertCommand.ExecuteNonQuery();

            cnnTarget.Close();

            MessageBox.Show("OK Skasowanie ");

        }

        private void dodanieTVzTabeliTmp()
        {
            SqlConnection cnnTarget;
            string kuraTVdopisanie;
            cnnTarget = new SqlConnection(connetionStringTarget);
            SqlDataAdapter adapterTarget = new SqlDataAdapter();
            cnnTarget.Open();
            kuraTVdopisanie = $"  insert into SN_TV (SerialNumber,Material ,DateScanned ,SalesOrganisation,ShiptoNumber,ShiptoName,ShiptoAddress " +
                               " ,Warehouse ,WarehouseName ,SAPDeliveryNumber ,SoldtoNumber ,SoldtoName ,FileID ,PalletID) " +
                              " SELECT dbo._SN_JG_PoleSNDN.SerialNumber, dbo._SN_JG_PoleSNDN.Material, dbo._SN_JG_PoleSNDN.DateScanned, dbo._SN_JG_PoleSNDN.SalesOrganisation, dbo._SN_JG_PoleSNDN.ShiptoNumber, dbo._SN_JG_PoleSNDN.ShiptoName,  " +
                              " dbo._SN_JG_PoleSNDN.ShiptoAddress, dbo._SN_JG_PoleSNDN.Warehouse, dbo._SN_JG_PoleSNDN.WarehouseName, dbo._SN_JG_PoleSNDN.SAPDeliveryNumber, dbo._SN_JG_PoleSNDN.SoldtoNumber,  " +
                              " dbo._SN_JG_PoleSNDN.SoldtoName, dbo._SN_JG_PoleSNDN.FileID, dbo._SN_JG_PoleSNDN.PalletID " +
                              "FROM     dbo._SN_JG_PoleSNDN LEFT OUTER JOIN dbo._SN_TV_PoleSNDN ON dbo._SN_JG_PoleSNDN.PLJG = dbo._SN_TV_PoleSNDN.PLJG WHERE(dbo._SN_TV_PoleSNDN.PLJG IS NULL) ";


            adapterTarget.InsertCommand = new SqlCommand(kuraTVdopisanie, cnnTarget);
            adapterTarget.InsertCommand.ExecuteNonQuery();

            cnnTarget.Close();
        }
        private void button18_Click(object sender, EventArgs e)
        {
            SqlConnection cnnTarget;
            string kuraTVdopisanie;
            cnnTarget = new SqlConnection(connetionStringTarget);
            SqlDataAdapter adapterTarget = new SqlDataAdapter();
            cnnTarget.Open();
            kuraTVdopisanie = $"  insert into SN_TV (SerialNumber,Material ,DateScanned ,SalesOrganisation,ShiptoNumber,ShiptoName,ShiptoAddress " +
                               " ,Warehouse ,WarehouseName ,SAPDeliveryNumber ,SoldtoNumber ,SoldtoName ,FileID ,PalletID) " +
                               " SELECT SN_JG.SerialNumber ,Material ,DateScanned ,SalesOrganisation ,ShiptoNumber ,ShiptoName ,ShiptoAddress " +
                               " ,Warehouse ,WarehouseName ,SAPDeliveryNumber ,SoldtoNumber ,SoldtoName ,FileID ,PalletID " +
                               "  from SN_JG inner join SNTV_NotInFinalTable as rs on SN_JG.SerialNumber = rs.SerialNumber";


            adapterTarget.InsertCommand = new SqlCommand(kuraTVdopisanie, cnnTarget);
            adapterTarget.InsertCommand.ExecuteNonQuery();

            cnnTarget.Close();

            MessageBox.Show("Dodane roznicowo do TV");

        }

        private void monitoryZTabeliTymczasowej()
        {
            SqlConnection cnnTarget;
            string kuraTVdopisanie;
            cnnTarget = new SqlConnection(connetionStringTarget);
            SqlDataAdapter adapterTarget = new SqlDataAdapter();
            cnnTarget.Open();
            kuraTVdopisanie = $" Insert into SN_MNT (GoodsIssueDate, Plant ,DN ,ShipToName ,PartNumber , SerialNumber , FileName ) " +
            "SELECT dbo._SN_MNT_JG_PoleSNDN.GoodsIssueDate, dbo._SN_MNT_JG_PoleSNDN.Plant, dbo._SN_MNT_JG_PoleSNDN.DN, dbo._SN_MNT_JG_PoleSNDN.ShipToName, dbo._SN_MNT_JG_PoleSNDN.PartNumber, " +
             "dbo._SN_MNT_JG_PoleSNDN.SerialNumber, dbo._SN_MNT_JG_PoleSNDN.FileName FROM     dbo._SN_MNT_JG_PoleSNDN LEFT OUTER JOIN " +
              "dbo._SN_MNT_PoleSNDN ON dbo._SN_MNT_JG_PoleSNDN.Expr1 = dbo._SN_MNT_PoleSNDN.PLMNT WHERE(dbo._SN_MNT_PoleSNDN.PLMNT IS NULL)";

            adapterTarget.InsertCommand = new SqlCommand(kuraTVdopisanie, cnnTarget);
            adapterTarget.InsertCommand.ExecuteNonQuery();

            cnnTarget.Close();
        }

        private void button19_Click(object sender, EventArgs e)
        {
            SqlConnection cnnTarget;
            string kuraTVdopisanie;
            cnnTarget = new SqlConnection(connetionStringTarget);
            SqlDataAdapter adapterTarget = new SqlDataAdapter();
            cnnTarget.Open();
            kuraTVdopisanie = $" Insert into SN_MNT (GoodsIssueDate, Plant ,DN ,ShipToName ,PartNumber , SerialNumber ) SELECT GoodsIssueDate " +
                ", Plant, DN, ShipToName, PartNumber, SN_MNT_JG.SerialNumber " +
                " FROM SN_MNT_JG inner join SNMNT_NotInFinalTable on SN_MNT_JG.SerialNumber = SNMNT_NotInFinalTable.SerialNumber ";

            adapterTarget.InsertCommand = new SqlCommand(kuraTVdopisanie, cnnTarget);
            adapterTarget.InsertCommand.ExecuteNonQuery();

            cnnTarget.Close();

            MessageBox.Show("Dodane roznicowo do TV");
        }

        private void button20_Click(object sender, EventArgs e)
        {
            SqlConnection cnnTarget;
            string kuraTVdopisanie;
            cnnTarget = new SqlConnection(connetionStringTarget);
            SqlDataAdapter adapterTarget = new SqlDataAdapter();
            cnnTarget.Open();
            kuraTVdopisanie = $"insert into SAPDNData_TV ( ShipToCustomerID, ShipToName ,ShipToCountry ,SoldToCustomerID ,SoldToName ,SoldToCountry " +
                               ",DeliveryPostDateTime, SAPDNData.DN) SELECT ShipToCustomerID, ShipToName, ShipToCountry, SoldToCustomerID, SoldToName,  SoldToCountry " +
                                ", DeliveryPostDateTime, SAPDNData.DN FROM SN_Tool.dbo.SAPDNData inner join DNTV_NotInFinalTable as rdn " +
                                 "on SAPDNData.DN = rdn.DN ";

            adapterTarget.InsertCommand = new SqlCommand(kuraTVdopisanie, cnnTarget);
            adapterTarget.InsertCommand.ExecuteNonQuery();

            cnnTarget.Close();

            MessageBox.Show("Dodane roznicowo do DN ek ");
        }

        private void zamknijProcessExcela()
        {
            foreach (var process in Process.GetProcessesByName("EXCEL"))
            {
                process.Kill();
            }
        }


        public void Wait(int time)
        {
            Thread thread = new Thread(delegate ()
            {
                System.Threading.Thread.Sleep(time);
            });
            thread.Start();
            while (thread.IsAlive)
                Application.DoEvents();
        }

        private void updateSNTV_Dnkami()
        {
            SqlConnection cnnTarget;

            cnnTarget = new SqlConnection(connetionStringTarget);
            SqlDataAdapter adapterTarget = new SqlDataAdapter();
            cnnTarget.Open();

            string kuraDNUpdate;

            kuraDNUpdate = "update SN_TV set SN_TV.ShiptoName = SAPDNData_TV.ShipToName, " +
                    "SN_TV.ShiptoAddress = SAPDNData_TV.SoldToCountry, " +
                    " SN_TV.SoldtoName = SAPDNData_TV.ShipToName " +
                    " from SN_TV inner join SAPDNData_TV on ltrim(rtrim(SAPDNData_TV.DN)) = ltrim(rtrim(SN_TV.SAPDeliveryNumber)) ";

            adapterTarget.InsertCommand = new SqlCommand(kuraDNUpdate, cnnTarget);
            adapterTarget.InsertCommand.ExecuteNonQuery();

            cnnTarget.Close();
        }

        private void button23_Click(object sender, EventArgs e)
        {
            SqlConnection cnnTarget;

            cnnTarget = new SqlConnection(connetionStringTarget);
            SqlDataAdapter adapterTarget = new SqlDataAdapter();
            cnnTarget.Open();

            string kuraDNUpdate;

            kuraDNUpdate = "update SN_TV set SN_TV.ShiptoName = SAPDNData_TV.ShipToName, " +
                    "SN_TV.ShiptoAddress = SAPDNData_TV.SoldToCountry, " +
                    " SN_TV.SoldtoName = SAPDNData_TV.ShipToName " +
                    " from SN_TV inner join SAPDNData_TV on ltrim(rtrim(SAPDNData_TV.DN)) = ltrim(rtrim(SN_TV.SAPDeliveryNumber)) ";

            adapterTarget.InsertCommand = new SqlCommand(kuraDNUpdate, cnnTarget);
            adapterTarget.InsertCommand.ExecuteNonQuery();

            cnnTarget.Close();

            MessageBox.Show(" Update SN_TV DN-ek skonczone  ");


        }



        private void button2_Click_2(object sender, EventArgs e)
        {
            wczytajWinScpPlikizFtp();

        }

        private void wczytajWinScpPlikizFtp()
        {
            string sFilename;

            var katalogi = new List<string>();

            var fulllangs = new List<string>();

            katalogi.Add(@"/SN_Tool/Batta");
            katalogi.Add(@"/SN_Tool/DaGanzo");
            katalogi.Add(@"/SN_Tool/GZ");
            katalogi.Add(@"/SN_Tool/MNT");
            katalogi.Add(@"/SN_Tool/Venlo");
            katalogi.Add(@"/SN_Tool/SAPDNData");

            SessionOptions FtpsessionOptions = new SessionOptions
            {

                Protocol = Protocol.Ftp,
                HostName = "172.26.59.13",
                UserName = "SN_Tool",
                Password = "46qD8N4w",
                FtpSecure = FtpSecure.Explicit,
                TlsHostCertificateFingerprint = "1e:91:90:86:47:16:96:7d:12:c4:ac:3f:0f:04:98:c2:3c:78:a5:0c",


            };


            SessionOptions sFtpsessionOptions = new SessionOptions
            {

                Protocol = Protocol.Sftp,
                HostName = "128.127.93.50",
                UserName = "SN_Tool",
                Password = "46qD8N4w",
                SshHostKeyFingerprint = "ssh-rsa 2048 GYol3dF9i8Br6BvsU469/Lx3qdUA18gNwIuaN/DORdE=",

            };

            using (WinSCP.Session session = new WinSCP.Session())
            {
                // Connect
                if (RodzajFTP == "sFtp")
                {
                    session.Open(sFtpsessionOptions);
                }


                if (RodzajFTP == "Ftp")
                {
                    session.Open(FtpsessionOptions);
                }
                foreach (string kat in katalogi)
                {

                    IEnumerable<RemoteFileInfo> fileInfos =
                        session.EnumerateRemoteFiles(
                            kat, null,
                            EnumerationOptions.EnumerateDirectories |
                                EnumerationOptions.AllDirectories);


                    foreach (RemoteFileInfo fileInfo in fileInfos)
                    {

                        sFilename = fileInfo.FullName.ToString();
                        fulllangs.Add(sFilename.ToString());
                    }

                    this.listboxFiles.Items.Clear();
                    foreach (string fl in fulllangs)
                    {
                        this.listboxFiles.Items.Add(fl);
                    }
                }
                // if (this.listboxFiles.Items.Count == 0) MessageBox.Show(" No files on ftp !!!");

            }
        }

        private void wczytajPlikizFtp()
        {
            List<string> listboxFiles = new List<string>();


            var fulllangs = new List<string>();


            var langs = new List<string>();

            var katalogi = new List<string>();


            katalogi.Add(@"/SN_Tool/Batta");
            katalogi.Add(@"/SN_Tool/DaGanzo");
            katalogi.Add(@"/SN_Tool/GZ");
            katalogi.Add(@"/SN_Tool/MNT");
            katalogi.Add(@"/SN_Tool/Venlo");
            katalogi.Add(@"/SN_Tool/SAPDNData");

            foreach (string kat in katalogi)
            {
                //string skatalog = @"/SN_Tool/test";

                string skatalog = kat.ToString();
                FtpWebRequest request = (FtpWebRequest)WebRequest.Create(shost + skatalog);
                request.Method = WebRequestMethods.Ftp.ListDirectory;

                request.Credentials = new NetworkCredential(susername, spassword);
                FtpWebResponse response = (FtpWebResponse)request.GetResponse();
                Stream responseStream = response.GetResponseStream();
                StreamReader reader = new StreamReader(responseStream);
                string names = reader.ReadToEnd();

                langs = names.Split(new string[] { "\r\n" }, StringSplitOptions.RemoveEmptyEntries).ToList();
                fulllangs.AddRange(langs);
                langs = null;
                reader.Close();
                response.Close();
            }
            this.listboxFiles.Items.Clear();
            foreach (string fl in fulllangs)
            {
                this.listboxFiles.Items.Add(fl);
            }
          ;
        }



        List<string> listaPlikowFtpKatalog()
        {
            var langs = new List<string>();
            FtpWebRequest request = (FtpWebRequest)WebRequest.Create(shost + @"/SN_Tool/GZ");
            request.Method = WebRequestMethods.Ftp.ListDirectory;

            request.Credentials = new NetworkCredential(susername, spassword);
            FtpWebResponse response = (FtpWebResponse)request.GetResponse();
            Stream responseStream = response.GetResponseStream();
            StreamReader reader = new StreamReader(responseStream);
            string names = reader.ReadToEnd();
            reader.Close();
            response.Close();
            return names.Split(new string[] { "\r\n" }, StringSplitOptions.RemoveEmptyEntries).ToList();

        }


        private void DeleteWinScpFTPFile(string plikdoSkasowania)
        {

            SessionOptions FtpsessionOptions = new SessionOptions
            {

                Protocol = Protocol.Ftp,
                HostName = "172.26.59.13",
                UserName = "SN_Tool",
                Password = "46qD8N4w",

                FtpSecure = FtpSecure.Explicit,
                TlsHostCertificateFingerprint = "1e:91:90:86:47:16:96:7d:12:c4:ac:3f:0f:04:98:c2:3c:78:a5:0c",

            };

            SessionOptions sFtpsessionOptions = new SessionOptions
            {

                Protocol = Protocol.Sftp,
                HostName = "128.127.93.50",
                UserName = "SN_Tool",
                Password = "46qD8N4w",
                SshHostKeyFingerprint = "ssh-rsa 2048 GYol3dF9i8Br6BvsU469/Lx3qdUA18gNwIuaN/DORdE=",

            };
            using (WinSCP.Session session = new WinSCP.Session())
            {
                // Connect
                if (RodzajFTP == "sFtp")
                {
                    session.Open(sFtpsessionOptions);
                }


                if (RodzajFTP == "Ftp")
                {
                    session.Open(FtpsessionOptions);
                }
                // Upload files
                TransferOptions transferOptions = new TransferOptions();
                transferOptions.TransferMode = TransferMode.Binary;

                session.RemoveFile(plikdoSkasowania);
                // Throw on any error

            }
        }
        private void button5_Click_1(object sender, EventArgs e)
        {


            //////////////////////////////////
            ///


            ////////////////



            string sFilename;

            string skatalog = @"/SN_Tool/MNT";

            SessionOptions FtpsessionOptions = new SessionOptions
            {

                Protocol = Protocol.Ftp,
                UserName = "SN_Tool",
                Password = "46qD8N4w",

                HostName = "172.26.59.13",
                FtpSecure = FtpSecure.Explicit,
                TlsHostCertificateFingerprint = "1e:91:90:86:47:16:96:7d:12:c4:ac:3f:0f:04:98:c2:3c:78:a5:0c",

            };


            SessionOptions sFtpsessionOptions = new SessionOptions
            {

                Protocol = Protocol.Sftp,
                HostName = "128.127.93.50",
                UserName = "SN_Tool",
                Password = "46qD8N4w",
                SshHostKeyFingerprint = "ssh-rsa 2048 GYol3dF9i8Br6BvsU469/Lx3qdUA18gNwIuaN/DORdE=",

            };

            using (WinSCP.Session session = new WinSCP.Session())
            {
                // Connect
                if (RodzajFTP == "sFtp")
                {
                    session.Open(sFtpsessionOptions);
                }


                if (RodzajFTP == "Ftp")
                {
                    session.Open(FtpsessionOptions);
                }

                IEnumerable<RemoteFileInfo> fileInfos =
                    session.EnumerateRemoteFiles(
                        skatalog, null,
                        EnumerationOptions.EnumerateDirectories |
                            EnumerationOptions.AllDirectories);
                foreach (RemoteFileInfo fileInfo in fileInfos)
                {

                    sFilename = fileInfo.FullName.ToString();
                    MessageBox.Show(sFilename.ToString());
                }
            }
        }

        private void CevaWinSCP(string plikSciezka)
        {
            string katalogLokalny = @"c:\DC\Gz\";
            string plik = plikSciezka.Substring(plikSciezka.LastIndexOf(@"/") + 1, plikSciezka.Length - plikSciezka.LastIndexOf(@"/") - 1);
            string plikKataloglokalnyplik = katalogLokalny.Trim() + plik.Trim();
            string plikKatalogNaftpBackup = @"/SN_Backup/GZ/";

            kasowanieTabelTmp();
            Wait(3000);
            kasujPliki(@"C:\DC\Batta");
            kasujPliki(@"C:\DC\Daganzo");
            kasujPliki(@"C:\DC\GZ");
            kasujPliki(@"C:\DC\MNT");
            kasujPliki(@"C:\DC\SAPDNData");
            kasujPliki(@"C:\DC\Venlo");
            Wait(3000);

            if (plikSciezka.Length > 0)
            {
                ftpWinScpDownload(plikSciezka, katalogLokalny);
                wpiszDoLoga("Skopiowano z ftp ", "CEVA GZ", plik, "TV");
                Wait(3000);
                zamknijProcessExcela();
                wczytajCeveDoBazy(plikKataloglokalnyplik);
                wpiszDoLoga("Wczytano do bazy ", "CEVA GZ", plik, "TV");
                Wait(3000);
                zamknijProcessExcela();
                ftpWinScpUpload(plikKataloglokalnyplik, plikKatalogNaftpBackup);
                wpiszDoLoga("Skopiowani backup ftp ", "CEVA GZ", plik, "TV");
                DeleteWinScpFTPFile(plikSciezka);
                wpiszDoLoga("Skasowano ftp ", "CEVA GZ", plik, "TV");
                Wait(3000);
                dodanieTVzTabeliTmp();
                wpiszDoLoga("Dodano z tmp do SN_TV ", "CEVA GZ", plik, "TV");
                Wait(3000);
                updateSNTV_Dnkami();
                wpiszDoLoga("Update DN-Kami ", "CEVA GZ", plik, "TV");
            }
        }

        private void button10_Click_1(object sender, EventArgs e)
        {


            string sFilename;

            string skatalog = @"/SN_Tool/GZ";

            SessionOptions FtpsessionOptions = new SessionOptions
            {

                Protocol = Protocol.Ftp,
                HostName = "172.26.59.13",
                UserName = "SN_Tool",
                Password = "46qD8N4w",
                FtpSecure = FtpSecure.Explicit,
                TlsHostCertificateFingerprint = "1e:91:90:86:47:16:96:7d:12:c4:ac:3f:0f:04:98:c2:3c:78:a5:0c",

            };


            SessionOptions sFtpsessionOptions = new SessionOptions
            {

                Protocol = Protocol.Sftp,
                HostName = "128.127.93.50",
                UserName = "SN_Tool",
                Password = "46qD8N4w",
                SshHostKeyFingerprint = "ssh-rsa 2048 GYol3dF9i8Br6BvsU469/Lx3qdUA18gNwIuaN/DORdE=",

            };

            using (WinSCP.Session session = new WinSCP.Session())
            {
                // Connect
                if (RodzajFTP == "sFtp")
                {
                    session.Open(sFtpsessionOptions);
                }


                if (RodzajFTP == "Ftp")
                {
                    session.Open(FtpsessionOptions);
                }

                IEnumerable<RemoteFileInfo> fileInfos =
                    session.EnumerateRemoteFiles(
                        skatalog, null,
                        EnumerationOptions.EnumerateDirectories |
                            EnumerationOptions.AllDirectories);
                foreach (RemoteFileInfo fileInfo in fileInfos)
                {

                    sFilename = fileInfo.FullName.ToString();
                    CevaWinSCP(sFilename);
                }
            }
            MessageBox.Show(" skonczone OK Ceva ");


        }

        private void VenloWinSCP(string plikSciezka)
        {
            string katalogLokalny = @"c:\DC\Venlo\";
            string plik = plikSciezka.Substring(plikSciezka.LastIndexOf(@"/") + 1, plikSciezka.Length - plikSciezka.LastIndexOf(@"/") - 1);
            string plikKataloglokalnyplik = katalogLokalny.Trim() + plik.Trim();
            string plikKatalogNaftpBackup = @"/SN_Backup/Venlo/";

            kasowanieTabelTmp();
            Wait(3000);
            kasujPliki(@"C:\DC\Batta");
            kasujPliki(@"C:\DC\Daganzo");
            kasujPliki(@"C:\DC\GZ");
            kasujPliki(@"C:\DC\MNT");
            kasujPliki(@"C:\DC\SAPDNData");
            kasujPliki(@"C:\DC\Venlo");
            Wait(3000);

            if (plikSciezka.Length > 0)
            {
                ftpWinScpDownload(plikSciezka, katalogLokalny);
                wpiszDoLoga("Skopiowano z ftp ", "Venlo", plik, "TV");
                Wait(3000);
                zamknijProcessExcela();
                wczytajVenloDoBazy(plikKataloglokalnyplik);
                wpiszDoLoga("Wczytano do bazy ", "Venlo", plik, "TV");
                Wait(3000);
                zamknijProcessExcela();
                ftpWinScpUpload(plikKataloglokalnyplik, plikKatalogNaftpBackup);
                wpiszDoLoga("Skopiowani backup ftp ", "Venlo", plik, "TV");
                DeleteWinScpFTPFile(plikSciezka);
                wpiszDoLoga("Skasowano ftp ", "Venlo", plik, "TV");
                Wait(3000);
                dodanieTVzTabeliTmp();
                Wait(3000);
                updateSNTV_Dnkami();
                wpiszDoLoga("Update DN-Kami ", "Venlo", plik, "TV");

            }

        }

        private void button15_Click_1(object sender, EventArgs e)
        {


            string sFilename;

            string skatalog = @"/SN_Tool/Venlo";

            SessionOptions FtpsessionOptions = new SessionOptions
            {

                Protocol = Protocol.Ftp,
                HostName = "172.26.59.13",
                UserName = "SN_Tool",
                Password = "46qD8N4w",
                FtpSecure = FtpSecure.Explicit,
                TlsHostCertificateFingerprint = "1e:91:90:86:47:16:96:7d:12:c4:ac:3f:0f:04:98:c2:3c:78:a5:0c",

            };


            SessionOptions sFtpsessionOptions = new SessionOptions
            {

                Protocol = Protocol.Sftp,
                HostName = "128.127.93.50",
                UserName = "SN_Tool",
                Password = "46qD8N4w",
                SshHostKeyFingerprint = "ssh-rsa 2048 GYol3dF9i8Br6BvsU469/Lx3qdUA18gNwIuaN/DORdE=",

            };

            using (WinSCP.Session session = new WinSCP.Session())
            {
                // Connect
                if (RodzajFTP == "sFtp")
                {
                    session.Open(sFtpsessionOptions);
                }


                if (RodzajFTP == "Ftp")
                {
                    session.Open(FtpsessionOptions);
                }

                IEnumerable<RemoteFileInfo> fileInfos =
                    session.EnumerateRemoteFiles(
                        skatalog, null,
                        EnumerationOptions.EnumerateDirectories |
                            EnumerationOptions.AllDirectories);
                foreach (RemoteFileInfo fileInfo in fileInfos)
                {

                    sFilename = fileInfo.FullName.ToString();
                    VenloWinSCP(sFilename);
                }
            }
            MessageBox.Show(" skonczone OK Venlo ");


        }

        private void DaganzoWinSCP(string plikSciezka)
        {
            string katalogLokalny = @"c:\DC\Daganzo\";
            string plik = plikSciezka.Substring(plikSciezka.LastIndexOf(@"/") + 1, plikSciezka.Length - plikSciezka.LastIndexOf(@"/") - 1);
            string plikKataloglokalnyplik = katalogLokalny.Trim() + plik.Trim();
            string plikKatalogNaftpBackup = @"/SN_Backup/Daganzo/";


            kasowanieTabelTmp();
            Wait(3000);
            kasujPliki(@"C:\DC\Batta");
            kasujPliki(@"C:\DC\Daganzo");
            kasujPliki(@"C:\DC\GZ");
            kasujPliki(@"C:\DC\MNT");
            kasujPliki(@"C:\DC\SAPDNData");
            kasujPliki(@"C:\DC\Venlo");
            Wait(3000);


            if (plikSciezka.Length > 0)
            {
                ftpWinScpDownload(plikSciezka, katalogLokalny);
                wpiszDoLoga("Skopiowano z ftp ", "Daganzo", plik, "TV");
                Wait(3000);
                zamknijProcessExcela();
                wczytajDaganzoDoBazy(plikKataloglokalnyplik);
                wpiszDoLoga("Wczytano do bazy ", "Daganzo", plik, "TV");
                Wait(3000);
                zamknijProcessExcela();
                ftpWinScpUpload(plikKataloglokalnyplik, plikKatalogNaftpBackup);
                wpiszDoLoga("Skopiowani backup ftp ", "Daganzo", plik, "TV");
                DeleteWinScpFTPFile(plikSciezka);
                wpiszDoLoga("Skasowano ftp ", "Daganzo", plik, "TV");
                dodanieTVzTabeliTmp();
                wpiszDoLoga("Dodano z tmp do SN_TV ", "Daganzo", plik, "TV");
                Wait(3000);
                updateSNTV_Dnkami();
                wpiszDoLoga("Update DN-Kami ", "Daganzo", plik, "TV");

            }
        }

        private void button22_Click(object sender, EventArgs e)
        {
            string plikSciezka = sprawdzPlikWinScpFtp(@"/SN_Tool/Daganzo/");
            string sFilename;

            string skatalog = @"/SN_Tool/DaGanzo";

            SessionOptions FtpsessionOptions = new SessionOptions
            {

                Protocol = Protocol.Ftp,
                HostName = "172.26.59.13",
                UserName = "SN_Tool",
                Password = "46qD8N4w",
                FtpSecure = FtpSecure.Explicit,
                TlsHostCertificateFingerprint = "1e:91:90:86:47:16:96:7d:12:c4:ac:3f:0f:04:98:c2:3c:78:a5:0c",

            };


            SessionOptions sFtpsessionOptions = new SessionOptions
            {

                Protocol = Protocol.Sftp,
                HostName = "128.127.93.50",
                UserName = "SN_Tool",
                Password = "46qD8N4w",
                SshHostKeyFingerprint = "ssh-rsa 2048 GYol3dF9i8Br6BvsU469/Lx3qdUA18gNwIuaN/DORdE=",

            };

            using (WinSCP.Session session = new WinSCP.Session())
            {
                // Connect
                if (RodzajFTP == "sFtp")
                {
                    session.Open(sFtpsessionOptions);
                }


                if (RodzajFTP == "Ftp")
                {
                    session.Open(FtpsessionOptions);
                }

                IEnumerable<RemoteFileInfo> fileInfos =
                    session.EnumerateRemoteFiles(
                        skatalog, null,
                        EnumerationOptions.EnumerateDirectories |
                            EnumerationOptions.AllDirectories);
                foreach (RemoteFileInfo fileInfo in fileInfos)
                {

                    sFilename = fileInfo.FullName.ToString();
                    DaganzoWinSCP(sFilename);
                }
            }
            MessageBox.Show(" skonczone OK Daganzo ");


        }

        private void BattaWinSCP(string plikSciezka)
        {

            string katalogLokalny = @"c:\DC\Batta\";
            string plik = plikSciezka.Substring(plikSciezka.LastIndexOf(@"/") + 1, plikSciezka.Length - plikSciezka.LastIndexOf(@"/") - 1);
            string plikKataloglokalnyplik = katalogLokalny.Trim() + plik.Trim();
            string plikKatalogNaftpBackup = @"/SN_Backup/Batta/";


            kasowanieTabelTmp();
            Wait(3000);
            kasujPliki(@"C:\DC\Batta");
            kasujPliki(@"C:\DC\Daganzo");
            kasujPliki(@"C:\DC\GZ");
            kasujPliki(@"C:\DC\MNT");
            kasujPliki(@"C:\DC\SAPDNData");
            kasujPliki(@"C:\DC\Venlo");
            Wait(3000);


            if (plikSciezka.Length > 0)
            {
                ftpWinScpDownload(plikSciezka, katalogLokalny);
                wpiszDoLoga("Skopiowano z ftp ", "Batta", plik, "TV");
                Wait(3000);
                zamknijProcessExcela();
                wczytajBattaDoBazy(plikKataloglokalnyplik);
                wpiszDoLoga("Wczytano do bazy ", "Batta", plik, "TV");
                Wait(3000);
                zamknijProcessExcela();
                ftpWinScpUpload(plikKataloglokalnyplik, plikKatalogNaftpBackup);
                wpiszDoLoga("Skopiowani backup ftp ", "Batta", plik, "TV");
                DeleteWinScpFTPFile(plikSciezka);
                wpiszDoLoga("Skasowano ftp ", "Batta", plik, "TV");
                dodanieTVzTabeliTmp();
                wpiszDoLoga("Dodano z tmp do SN_TV ", "Batta", plik, "TV");
                Wait(3000);
                updateSNTV_Dnkami();
                wpiszDoLoga("Update DN-Kami ", "Batta", plik, "TV");


            }
        }

        private void button24_Click(object sender, EventArgs e)
        {

            string sFilename;

            string skatalog = @"/SN_Tool/Batta";

            SessionOptions FtpsessionOptions = new SessionOptions
            {

                Protocol = Protocol.Ftp,
                HostName = "172.26.59.13",
                UserName = "SN_Tool",
                Password = "46qD8N4w",
                FtpSecure = FtpSecure.Explicit,
                TlsHostCertificateFingerprint = "1e:91:90:86:47:16:96:7d:12:c4:ac:3f:0f:04:98:c2:3c:78:a5:0c",

            };


            SessionOptions sFtpsessionOptions = new SessionOptions
            {

                Protocol = Protocol.Sftp,
                HostName = "128.127.93.50",
                UserName = "SN_Tool",
                Password = "46qD8N4w",
                SshHostKeyFingerprint = "ssh-rsa 2048 GYol3dF9i8Br6BvsU469/Lx3qdUA18gNwIuaN/DORdE=",

            };

            using (WinSCP.Session session = new WinSCP.Session())
            {
                // Connect
                if (RodzajFTP == "sFtp")
                {
                    session.Open(sFtpsessionOptions);
                }


                if (RodzajFTP == "Ftp")
                {
                    session.Open(FtpsessionOptions);
                }

                IEnumerable<RemoteFileInfo> fileInfos =
                    session.EnumerateRemoteFiles(
                        skatalog, null,
                        EnumerationOptions.EnumerateDirectories |
                            EnumerationOptions.AllDirectories);
                foreach (RemoteFileInfo fileInfo in fileInfos)
                {

                    sFilename = fileInfo.FullName.ToString();
                    BattaWinSCP(sFilename);
                }
            }
            MessageBox.Show(" skonczone OK Batta ");


        }

        private void button25_Click(object sender, EventArgs e)
        {
            //     string plikSciezka = sprawdzPlikWinScpFtp(@"/SN_Tool/MNT/");

            string sFilename;

            string skatalog = @"/SN_Tool/MNT";

            SessionOptions FtpsessionOptions = new SessionOptions
            {

                Protocol = Protocol.Ftp,
                HostName = "172.26.59.13",
                UserName = "SN_Tool",
                Password = "46qD8N4w",
                FtpSecure = FtpSecure.Explicit,
                TlsHostCertificateFingerprint = "1e:91:90:86:47:16:96:7d:12:c4:ac:3f:0f:04:98:c2:3c:78:a5:0c",

            };


            SessionOptions sFtpsessionOptions = new SessionOptions
            {

                Protocol = Protocol.Sftp,
                HostName = "128.127.93.50",
                UserName = "SN_Tool",
                Password = "46qD8N4w",
                SshHostKeyFingerprint = "ssh-rsa 2048 GYol3dF9i8Br6BvsU469/Lx3qdUA18gNwIuaN/DORdE=",

            };

            using (WinSCP.Session session = new WinSCP.Session())
            {
                // Connect
                if (RodzajFTP == "sFtp")
                {
                    session.Open(sFtpsessionOptions);
                }


                if (RodzajFTP == "Ftp")
                {
                    session.Open(FtpsessionOptions);
                }

                IEnumerable<RemoteFileInfo> fileInfos =
                    session.EnumerateRemoteFiles(
                        skatalog, null,
                        EnumerationOptions.EnumerateDirectories |
                            EnumerationOptions.AllDirectories);
                foreach (RemoteFileInfo fileInfo in fileInfos)
                {

                    sFilename = fileInfo.FullName.ToString();
                    MonitoryWinSCP(sFilename);
                }
            }
            MessageBox.Show(" skonczone OK MNT ");
        }


        private void MonitoryWinSCP(string plikSciezka)
        {

            string katalogLokalny = @"c:\DC\MNT\";
            string plik = plikSciezka.Substring(plikSciezka.LastIndexOf(@"/") + 1, plikSciezka.Length - plikSciezka.LastIndexOf(@"/") - 1);
            string plikKataloglokalnyplik = katalogLokalny.Trim() + plik.Trim();
            string plikKatalogNaftpBackup = @"/SN_Backup/MNT/";

            kasowanieTabelTmp();
            Wait(3000);
            kasujPliki(@"C:\DC\Batta");
            kasujPliki(@"C:\DC\Daganzo");
            kasujPliki(@"C:\DC\GZ");
            kasujPliki(@"C:\DC\MNT");
            kasujPliki(@"C:\DC\SAPDNData");
            kasujPliki(@"C:\DC\Venlo");
            Wait(3000);


            if (plik.Length > 0)
            {

                if (plikSciezka.Length > 0)
                {
                    ftpWinScpDownload(plikSciezka, katalogLokalny);
                    wpiszDoLoga("Skopiowano z ftp ", "MNT GZ", plik, "TV");
                    Wait(3000);
                    zamknijProcessExcela();
                    WczytajMNTDoBazy(plikKataloglokalnyplik);
                    wpiszDoLoga("Wczytano do bazy ", "MNT GZ", plik, "TV");
                    Wait(3000);
                    zamknijProcessExcela();
                    ftpWinScpUpload(plikKataloglokalnyplik, plikKatalogNaftpBackup);
                    wpiszDoLoga("Skopiowani backup ftp ", "MNT GZ", plik, "TV");
                    DeleteWinScpFTPFile(plikSciezka);
                    wpiszDoLoga("Skasowano ftp ", "MNT GZ", plik, "TV");
                    Wait(3000);
                    monitoryZTabeliTymczasowej();
                    wpiszDoLoga("Dopisanie z tabeli tmp do SN_MNT ", "MNT GZ", plik, "TV");

                }
            }
        }
        private void dodanieDNSAPzTabeliTymczasowej()
        {
            SqlConnection cnnTarget;
            string kuraTVdopisanie;
            cnnTarget = new SqlConnection(connetionStringTarget);
            SqlDataAdapter adapterTarget = new SqlDataAdapter();
            cnnTarget.Open();
            kuraTVdopisanie = $"insert into SAPDNData_TV ( ShipToCustomerID, ShipToName ,ShipToCountry ,SoldToCustomerID ,SoldToName ,SoldToCountry " +
                               ",DeliveryPostDateTime, SAPDNData.DN , FileName) SELECT ShipToCustomerID, ShipToName, ShipToCountry, SoldToCustomerID, SoldToName,  SoldToCountry " +
                                ", DeliveryPostDateTime, SAPDNData.DN, FileName FROM SN_Tool.dbo.SAPDNData inner join DNTV_NotInFinalTable as rdn " +
                                 "on SAPDNData.DN = rdn.DN ";

            adapterTarget.InsertCommand = new SqlCommand(kuraTVdopisanie, cnnTarget);
            adapterTarget.InsertCommand.ExecuteNonQuery();

            cnnTarget.Close();
        }


        private void SapDNWinSCP(string plikSciezka)
        {
            string katalogLokalny = @"c:\DC\SAPDNData\";
            string plik = plikSciezka.Substring(plikSciezka.LastIndexOf(@"/") + 1, plikSciezka.Length - plikSciezka.LastIndexOf(@"/") - 1);
            string plikKataloglokalnyplik = katalogLokalny.Trim() + plik.Trim();
            string plikKatalogNaftpBackup = @"/SN_Backup/SAPDNData/";


            if (plikSciezka.Length > 0)
            {

                kasujPliki(@"C:\DC\SAPDNData");
                Wait(3000);
                kasowanieTabelTmp();
                Wait(3000);

                ftpWinScpDownload(plikSciezka, katalogLokalny);
                wpiszDoLoga("Skopiowano z ftp ", "SAPDNData", plik, "TV");
                Wait(3000);
                zamknijProcessExcela();
                wczytajSAPDNDoBazy(plikKataloglokalnyplik);
                wpiszDoLoga("Wczytano do bazy ", "SAPDNData", plik, "TV");
                Wait(3000);
                zamknijProcessExcela();
                ftpWinScpUpload(plikKataloglokalnyplik, plikKatalogNaftpBackup);
                wpiszDoLoga("Skopiowani backup ftp ", "SAPDNData", plik, "TV");
                DeleteWinScpFTPFile(plikSciezka);
                wpiszDoLoga("Skasowano ftp ", "SAPDNData", plik, "TV");
                dodanieDNSAPzTabeliTymczasowej();
                wpiszDoLoga("Dodanie do tabeli SAPDNData_TV z tmp ", "SAPDNData", plik, "TV");


            }
        }

        private void button26_Click(object sender, EventArgs e)
        {
            string sFilename;

            string skatalog = @"/SN_Tool/SAPDNData";

            SessionOptions FtpsessionOptions = new SessionOptions
            {

                Protocol = Protocol.Ftp,
                HostName = "172.26.59.13",
                UserName = "SN_Tool",
                Password = "46qD8N4w",
                FtpSecure = FtpSecure.Explicit,
                TlsHostCertificateFingerprint = "1e:91:90:86:47:16:96:7d:12:c4:ac:3f:0f:04:98:c2:3c:78:a5:0c",

            };


            SessionOptions sFtpsessionOptions = new SessionOptions
            {

                Protocol = Protocol.Sftp,
                HostName = "128.127.93.50",
                UserName = "SN_Tool",
                Password = "46qD8N4w",
                SshHostKeyFingerprint = "ssh-rsa 2048 GYol3dF9i8Br6BvsU469/Lx3qdUA18gNwIuaN/DORdE=",

            };

            using (WinSCP.Session session = new WinSCP.Session())
            {
                // Connect
                if (RodzajFTP == "sFtp")
                {
                    session.Open(sFtpsessionOptions);
                }


                if (RodzajFTP == "Ftp")
                {
                    session.Open(FtpsessionOptions);
                }

                IEnumerable<RemoteFileInfo> fileInfos =
                    session.EnumerateRemoteFiles(
                        skatalog, null,
                        EnumerationOptions.EnumerateDirectories |
                            EnumerationOptions.AllDirectories);
                foreach (RemoteFileInfo fileInfo in fileInfos)
                {
                    sFilename = fileInfo.FullName.ToString();
                    SapDNWinSCP(sFilename);
                }
            }
            MessageBox.Show(" skonczone OK DN - ki SAP ");


        }

        private void button3_Click(object sender, EventArgs e)
        {

            string sciezkaplik = @"c:\temp\FI_SAP_20200820_070013.xlsx";
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
            string plikDoZapisu = @"c:\temp\FI_SAP_20200820_070013.xlsx";

            #endregion

            // czesc wspolna dla wszystkich procedur 


            for (int i = 2; i <= rowCount; i++)
            {
                string dtString = xlRange.Cells[i, 7].Value2.ToString();

                Console.WriteLine(i);

            }

            xlWorkbook.Close(0);
            xlApp.Quit();

        }

        private void mainprocedureforbackgroudjob()
        {
            leci = true;
            wczytajWinScpPlikizFtp();


            string sFilename;

            string skatalog = @"/SN_Tool/SAPDNData";

            SessionOptions FtpsessionOptions = new SessionOptions
            {

                Protocol = Protocol.Ftp,
                HostName = "172.26.59.13",
                UserName = "SN_Tool",
                Password = "46qD8N4w",
                FtpSecure = FtpSecure.Explicit,
                TlsHostCertificateFingerprint = "1e:91:90:86:47:16:96:7d:12:c4:ac:3f:0f:04:98:c2:3c:78:a5:0c",

            };


            SessionOptions sFtpsessionOptions = new SessionOptions
            {

                Protocol = Protocol.Sftp,
                HostName = "128.127.93.50",
                UserName = "SN_Tool",
                Password = "46qD8N4w",
                SshHostKeyFingerprint = "ssh-rsa 2048 GYol3dF9i8Br6BvsU469/Lx3qdUA18gNwIuaN/DORdE=",

            };

            using (WinSCP.Session session = new WinSCP.Session())
            {
                // Connect
                if (RodzajFTP == "sFtp")
                {
                    session.Open(sFtpsessionOptions);
                }


                if (RodzajFTP == "Ftp")
                {
                    session.Open(FtpsessionOptions);
                }

                // SAP data 

                IEnumerable<RemoteFileInfo> fileInfos =
                    session.EnumerateRemoteFiles(
                        skatalog, null,
                        EnumerationOptions.EnumerateDirectories |
                            EnumerationOptions.AllDirectories);
                foreach (RemoteFileInfo fileInfo in fileInfos)
                {
                    sFilename = fileInfo.FullName.ToString();
                    SapDNWinSCP(sFilename);
                }

                wczytajWinScpPlikizFtp();
                // Gorzow Ceve

                //string skatalog = @"/SN_Tool/GZ";

                IEnumerable<RemoteFileInfo> fileInfosGZ =
                  session.EnumerateRemoteFiles(
                      @"/SN_Tool/GZ", null,
                      EnumerationOptions.EnumerateDirectories |
                          EnumerationOptions.AllDirectories);
                foreach (RemoteFileInfo fileInfo in fileInfosGZ)
                {
                    sFilename = fileInfo.FullName.ToString();
                    SapDNWinSCP(sFilename);
                }

                wczytajWinScpPlikizFtp();
                // Venlo

                //  string skatalog = @"/SN_Tool/Venlo";

                IEnumerable<RemoteFileInfo> fileInfosVenlo =
                  session.EnumerateRemoteFiles(
                      @"/SN_Tool/Venlo", null,
                      EnumerationOptions.EnumerateDirectories |
                          EnumerationOptions.AllDirectories);
                foreach (RemoteFileInfo fileInfo in fileInfosVenlo)
                {
                    sFilename = fileInfo.FullName.ToString();
                    SapDNWinSCP(sFilename);
                }

                // Daganzo 
                wczytajWinScpPlikizFtp();

                //  string skatalog = @"/SN_Tool/DaGanzo";
                IEnumerable<RemoteFileInfo> fileInfosDaganzo =
                session.EnumerateRemoteFiles(
                    @"/SN_Tool/DaGanzo", null,
                    EnumerationOptions.EnumerateDirectories |
                        EnumerationOptions.AllDirectories);
                foreach (RemoteFileInfo fileInfo in fileInfosDaganzo)
                {
                    sFilename = fileInfo.FullName.ToString();
                    SapDNWinSCP(sFilename);
                }

                // Bata
                wczytajWinScpPlikizFtp();

                //   string skatalog = @"/SN_Tool/Batta";


                IEnumerable<RemoteFileInfo> fileInfosBata =
               session.EnumerateRemoteFiles(
                    @"/SN_Tool/Batta", null,
                   EnumerationOptions.EnumerateDirectories |
                       EnumerationOptions.AllDirectories);
                foreach (RemoteFileInfo fileInfo in fileInfosBata)
                {
                    sFilename = fileInfo.FullName.ToString();
                    SapDNWinSCP(sFilename);
                }

                // Monitory 
                wczytajWinScpPlikizFtp();
                IEnumerable<RemoteFileInfo> fileInfosMonitory =
             session.EnumerateRemoteFiles(
                  @"/SN_Tool/Batta", null,
                 EnumerationOptions.EnumerateDirectories |
                     EnumerationOptions.AllDirectories);
                foreach (RemoteFileInfo fileInfo in fileInfosMonitory)
                {
                    sFilename = fileInfo.FullName.ToString();
                    SapDNWinSCP(sFilename);
                }

            }


            leci = false;


        }


        private  void TimerEventProcessor(object myObject, EventArgs myEventArgs)
        {

            m_streamWriter.WriteLine("{0} {1}", "Timer going", DateTime.Now.ToLongDateString() + ": " + DateTime.Now.ToLongTimeString());
            m_streamWriter.Flush();


       //     if (Convert.ToInt32(DateTime.Now.Minute.ToString().Trim()) == 50)
         if (lastHour < DateTime.Now.Hour)
            {
                lastHour = DateTime.Now.Hour;
                m_streamWriter.WriteLine("{0} {1}", "full hour ", DateTime.Now.ToLongDateString() + ": " + DateTime.Now.ToLongTimeString());
                m_streamWriter.Flush();
                if (leci == false)
                {        
                    m_streamWriter.WriteLine("{0} {1}", "Files processing", DateTime.Now.ToLongDateString() + ": " + DateTime.Now.ToLongTimeString());
                    m_streamWriter.Flush();
                    mainprocedureforbackgroudjob();
                    m_streamWriter.WriteLine("{0} {1}", "End Files processing", DateTime.Now.ToLongDateString() + ": " + DateTime.Now.ToLongTimeString());
                }
            }
        }

        private void button3_Click_1(object sender, EventArgs e)
        {
            mainprocedureforbackgroudjob();
        }

        private void button9_Click(object sender, EventArgs e)
        {
            //  System.Windows.Forms.Timer t = new System.Windows.Forms.Timer();


            Timer.Tick += TimerEventProcessor;
            Timer.Interval = 5000;
            Timer.Start();

        }

        private void button27_Click(object sender, EventArgs e)
        {
            Timer.Stop();
        }
    }
}




