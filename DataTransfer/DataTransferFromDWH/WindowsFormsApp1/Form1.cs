using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Data.SqlClient;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

using System.Threading;


namespace WindowsFormsApp1
{
    public partial class Form1 : Form
    {
        public Form1()
        {
            InitializeComponent();
        }
        //public string connetionStringTarget = @"Data Source= 172.17.80.141;Initial Catalog=SN_Tool;User Id=sn_tool;Password=F7E@Ln!j_viFoSW*R6bi";
        public string connetionStringSource = "Server=172.26.61.125;Initial Catalog=DWH;User ID=CCM_DCMon_sql_usr;Password=LolkaBolka+!1234";

        public string connetionStringTarget = @"Data Source= 172.17.80.141;Initial Catalog=SN_Tool;User Id=sn_tool;Password=F7E@Ln!j_viFoSW*R6bi";
       // public string connetionStringSource = "Server=JGTPVDQ6V5S4;Initial Catalog=DWH;Trusted_Connection=True";


        private void button1_Click(object sender, EventArgs e)
        {


            // zmienne pomocnicze 
            string strDzien;
            string strMc;
            int licznik = 0;
            string zmSoldtoName;
            string zmShiptoName;
            string zmShiptoAddress;

            string sqlSource1;
            SqlCommand Commandsource1;
            SqlDataReader DataReadersource1;
            SqlConnection cnnSource1;


            string sqlSource;
            SqlCommand Commandsource;
            SqlDataReader DataReadersource;
            SqlConnection cnnSource;

            cnnSource = new SqlConnection(connetionStringSource);
            cnnSource.Open();
            sqlSource = $"select * from dbo.SN where year(DateScanned) = { this.tbRok.Text }  and MONTH(datescanned) = {this.tbMc.Text}";

            Commandsource = new SqlCommand(sqlSource, cnnSource);
            DataReadersource = Commandsource.ExecuteReader();

            //---------- obliczenie ilosci rekordow

            cnnSource1 = new SqlConnection(connetionStringSource);
            cnnSource1.Open();
            sqlSource1 = $"select * from dbo.SN where year(DateScanned) = { this.tbRok.Text }  and MONTH(datescanned) = {this.tbMc.Text}";

            Commandsource1 = new SqlCommand(sqlSource1, cnnSource1);
            DataReadersource1 = Commandsource1.ExecuteReader();

            DataTable dt = new DataTable();
            dt.Load(DataReadersource1);
            int numRows = dt.Rows.Count;

          

            textBox1.Text = numRows.ToString();

            //----------------------

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
                sqlTarget = $" insert into SN_TV (SerialNumber, Material, DateScanned , SalesOrganisation , ShiptoNumber , ShiptoName , ShiptoAddress , Warehouse , WarehouseName , SAPDeliveryNumber , SoldtoNumber,  SoldtoName , FileID, PalletID ) " +
                    $"values (  \'  {  DataReadersource.GetValue(0)  }  \', \'  {  DataReadersource.GetValue(1)  }  \'  ,  \' {  zmData  } \'    , \'  {  DataReadersource.GetValue(3)  }  \'  , \'  {  DataReadersource.GetValue(4)  }  \' ," +
                    $" \' { zmShiptoName  }  \',  \' {  zmShiptoAddress  }  \' ,  \'  {  DataReadersource.GetValue(7)  }  \' ,  \'  {  DataReadersource.GetValue(8)  }  \' ,  \'  {  DataReadersource.GetValue(9)  }  \'," +
                    $" \'  {  DataReadersource.GetValue(10)  }  \',  \'  {  zmSoldtoName  }  \',  \'  {  DataReadersource.GetValue(12)  }  \' ,  \'  {  DataReadersource.GetValue(13)  }  \') ";

                adapterTarget.InsertCommand = new SqlCommand(sqlTarget, cnnTarget);
                adapterTarget.InsertCommand.ExecuteNonQuery();
                Console.WriteLine("licznik " + licznik.ToString());
                this.tbCounter.Text = licznik.ToString();
            }

        }

        private void Form1_Load(object sender, EventArgs e)
        {

        }
    }
}
