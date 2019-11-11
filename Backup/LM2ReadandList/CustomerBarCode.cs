using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Text;
using System.Windows.Forms;
using System.Data.SqlClient;

namespace LM2ReadandList
{
    public partial class CustomerBarCode : Form
    {
        public string ProductName = "", ListDate = "", Boxs="",Location="";

        //資料庫宣告
        string myConnectionString;
        SqlConnection myConnection;
        string selectCmd;
        SqlConnection conn;
        SqlCommand cmd;
        SqlDataReader reader;

        public CustomerBarCode()
        {
            InitializeComponent();

            //資料庫路徑與位子
            myConnectionString = "Server=192.168.0.15;database=amsys;uid=sa;pwd=ams.sql;";   
        }

        private void BarCodeTextBox_KeyPress(object sender, KeyPressEventArgs e)
        {
            if (e.KeyChar == (char)13 && BarCodeTextBox.Text.Trim().ToString()!="")
            {
                if (BarCodeTextBox.Text.Trim().ToString().Length != 12)
                {
                    MessageBox.Show("此客戶之Bar Code長度有問題，請重新輸入", "警告-W002");
                    BarCodeTextBox.Text = "";
                    return;
                }
                //if (BarCodeTextBox.Text.Trim().ToString().Substring(0, 6) != "110666")
                //{
                //    MessageBox.Show("此客戶之Bar Code有問題，請重新輸入", "警告-W003");
                //    BarCodeTextBox.Text = "";
                //    return;
                //}
                //檢查是否已存在
                //若已存在則告知在哪一箱哪個位置
                //判斷是否已經有相同的序號入嘜頭
                myConnection = new SqlConnection(myConnectionString);
                selectCmd = "SELECT  * FROM [ShippingBody] where [CustomerBarCode]='" + BarCodeTextBox.Text.Trim().ToString() + "'";
                conn = new SqlConnection(myConnectionString);
                conn.Open();
                cmd = new SqlCommand(selectCmd, conn);
                reader = cmd.ExecuteReader();
                if (reader.Read())
                {
                    MessageBox.Show("此客戶之Bar Code已存在！\n\n在產品名稱為" + reader.GetString(1) + "\n出貨日期為" + reader.GetString(0) + "\n第" + reader.GetString(4) + "箱，第" + reader.GetString(5) + "位置\n氣瓶序號為" + reader.GetString(3), "警告-W001");
                    reader.Close();
                    conn.Close();
                    BarCodeTextBox.Text = "";
                    return;

                }
                reader.Close();
                conn.Close();

                myConnection = new SqlConnection(myConnectionString);
                selectCmd = "UPDATE[ShippingBody] SET [CustomerBarCode]='" + BarCodeTextBox.Text.Trim().ToString() + "' where [ListDate]='" + this.ListDate + "' and [ProductName]='" + this.ProductName + "' and [WhereBox]='" + this.Boxs + "' and [WhereSeat]='"+this.Location+"'";
                conn = new SqlConnection(myConnectionString);
                conn.Open();
                cmd = new SqlCommand(selectCmd, conn);
                reader = cmd.ExecuteReader();
                reader.Close();
                conn.Close();

                this.Close();

            }
        }
        private void Load_Excel()
        {
        }

        private void CustomerBarCode_Load(object sender, EventArgs e)
        {

        }
    }
}