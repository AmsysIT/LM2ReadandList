using com.google.zxing; // for BarcodeFormat
using com.google.zxing.common; // for ByteMatrix
using com.google.zxing.qrcode; // for QRCode Engine
using DataMatrix.net;
using System;
using System.Collections.Generic;
using System.Data;
using System.Data.SqlClient;
using System.Drawing;
using System.Drawing.Imaging; // for ImageFormat 
using System.Drawing.Printing;
using System.IO;
using System.Linq;
using System.Reflection;
using System.Runtime.InteropServices;
using System.Threading;
using System.Transactions;
using System.Windows.Forms;
using Excel = Microsoft.Office.Interop.Excel;

namespace LM2ReadandList
{
    public partial class Main : Form
    {
        //抓取控制印表機用
        [DllImport("user32.dll", EntryPoint = "SendMessageA")]
        private static extern int SendMessage(long hwnd, long wMsg, long wParam, string lParam);

        [DllImport("kernel32.dll")]
        static extern bool WriteProfileString(string lpAppName, string lpKeyName, string lpString);

        private const long HWND_BROADCAST = 0xffffL;
        private const long WM_WININICHANGE = 0x1a;

        string PalletNoString = "-";
        string Ebb = "";
        bool Message = false;
        public int time = 420;

        //資料庫宣告
        string myConnectionString, myConnectionString21, myConnectionString30, myConnectionString21_AMS_check;
        string AMS21_ConnectionString = "Server = 192.168.0.21; DataBase = AMS; uid = sa; pwd = dsc;";
        string selectCmd, selectCmd1;
        SqlConnection conn, conn1;
        SqlCommand cmd, cmd1;
        SqlDataReader reader, reader1;
        SqlDataAdapter sqlAdapter;
        DataTable DT, SDT;


        //用來記錄是否為PASS的字串
        string Pass = "N";
        string str = "";

        //用來記錄是否由瓶身開始讀取
        string BeGin = "N";

        //用來記錄上一個讀入的號碼
        string TempStr1 = "";
        string TempStr2 = "";


        //用來記錄瓶身瓶底位子 0=瓶身 1=瓶底 
        string Direction = "";


        string[] BoxsArray = new string[800];
        int[] BoxsCountArray = new int[800];

        //記錄一箱幾隻
        string bAboxof = "";

        int Getcount = 0;
        bool IsChangePrinter = false;
        public string ID = null;
        public string User = null;
        public string worktype;
        public string ProcessNo = null;

        public Main()
        {
            //資料庫路徑與位子
            myConnectionString = "Server=192.168.0.15;database=amsys;uid=sa;pwd=ams.sql;";
            myConnectionString30 = "Server=192.168.0.30;database=AMS2;uid=sa;pwd=Ams.sql;";
            myConnectionString21 = "Server=192.168.0.21;database=HRMDB;uid=sa;pwd=dsc;";
            myConnectionString21_AMS_check = "Server=192.168.0.21;database=AMS_Check;uid=sa;pwd=dsc;";
            InitializeComponent();
        }

        private void Form1_Load(object sender, EventArgs e)
        {
            IsChangePrinter = false;

            User_LB.Items.Clear();
            ProductName_CB.Items.Clear();

            using (conn = new SqlConnection(myConnectionString))
            {
                conn.Open();

                //載入員工
                selectCmd = "SELECT vchTestersNo,vchTestersName FROM [LaserMarkTesters]  ORDER BY vchTestersNo";
                cmd = new SqlCommand(selectCmd, conn);
                using (reader = cmd.ExecuteReader())
                {
                    while (reader.Read())
                    {
                        User_LB.Items.Add(reader.GetString(reader.GetOrdinal("vchTestersNo")) + " " + reader.GetString(reader.GetOrdinal("vchTestersName")));
                    }
                }

                //載入產品名稱
                selectCmd = "SELECT DISTINCT [ProductName] FROM [ShippingHead]  order by [ProductName] ";
                cmd = new SqlCommand(selectCmd, conn);
                using (reader = cmd.ExecuteReader())
                {
                    while (reader.Read())
                    {
                        ProductName_CB.Items.Add(reader.GetString(reader.GetOrdinal("ProductName")));
                    }
                }
            }

            //20200420 
            DT = new DataTable();
            selectCmd = "SELECT [vchManufacturingNo],[vchMarkingType],[CylinderNo],[vchHydrostaticTestDate],[ClientName],HydroLabelPass FROM [MSNBody] " +
                "  where Package = '0' ";
            sqlAdapter = new SqlDataAdapter(selectCmd, myConnectionString);
            sqlAdapter.Fill(DT);

            SDT = new DataTable();
            selectCmd = "SELECT vchBoxs FROM ShippingHead where [DemandNo] = '2201-20200409001' and ( [ProductNo] = '4C8208226188138030' or [ProductNo] = '4C7208226188100030' ) ";
            sqlAdapter = new SqlDataAdapter(selectCmd, myConnectionString);
            sqlAdapter.Fill(SDT);
        }

        private void LoadListDate()
        {
            ListDate_LB.SelectedIndex = -1;
            ListDate_LB.Items.Clear();

            //載入[ShippingHead]的ListDate
            //加入vchPrint之條件 20190212
            using (conn = new SqlConnection(myConnectionString))
            {
                conn.Open();

                selectCmd = "SELECT  DISTINCT [ListDate] FROM [ShippingHead]  where [ProductName]='" + ProductName_CB.SelectedItem.ToString() + "' and vchPrint='" + ColorListBox.SelectedItem.ToString() + "' order by [ListDate] desc";
                cmd = new SqlCommand(selectCmd, conn);
                using (reader = cmd.ExecuteReader())
                {
                    while (reader.Read())
                    {
                        ListDate_LB.Items.Add(reader.GetString(reader.GetOrdinal("ListDate")));
                    }
                }
            }
        }

        public void LoadPictrue()
        {
            try
            {
                //記錄目前裝到第幾個位子
                string SeatNo = "";

                //記錄一箱幾隻
                bAboxof = "";

                //判斷此嘜頭幾隻一箱
                using (conn = new SqlConnection(myConnectionString))
                {
                    conn.Open();

                    selectCmd = "SELECT [vchAboxof] FROM [ShippingHead] where [ListDate]='" + ListDate_LB.SelectedItem + "' and [ProductName]='" + ProductName_CB.SelectedItem + "'and [vchBoxs]='" + WhereBox_LB.SelectedItem + "'";
                    cmd = new SqlCommand(selectCmd, conn);
                    using (reader = cmd.ExecuteReader())
                    {
                        if (reader.Read())
                        {
                            bAboxof = reader.GetString(reader.GetOrdinal("vchAboxof"));
                        }
                    }
                }

                if (bAboxof == "20" || bAboxof == "40")
                {
                    //載入[ShippingHead]的ListDate
                    using (conn = new SqlConnection(myConnectionString))
                    {
                        conn.Open();

                        selectCmd = "SELECT [WhereSeat] FROM [ShippingBody] where [ListDate]='" + ListDate_LB.SelectedItem + "' and [ProductName]='" + ProductName_CB.SelectedItem + "'and [WhereBox]='" + WhereBox_LB.SelectedItem + "'  order by Convert(INT,[WhereSeat]) DESC ";
                        cmd = new SqlCommand(selectCmd, conn);
                        using (reader = cmd.ExecuteReader())
                        {
                            if (reader.Read())
                            {
                                SeatNo = reader.GetString(reader.GetOrdinal("WhereSeat"));

                                if (reader.IsDBNull(reader.GetOrdinal("WhereSeat")) == false && (Convert.ToInt32(reader.GetString(reader.GetOrdinal("WhereSeat"))) >= 1 && Convert.ToInt32(reader.GetString(reader.GetOrdinal("WhereSeat"))) <= 20))
                                {
                                    pictureBox1.Image = Image.FromFile(Application.StartupPath + @".\" + reader.GetString(reader.GetOrdinal("WhereSeat")) + ".jpg");
                                }
                            }
                            else
                            {
                                pictureBox1.Image = Image.FromFile(Application.StartupPath + @".\0.jpg");
                            }
                        }
                    }
                }

                if (bAboxof == "40")
                {
                    switch (SeatNo)
                    {
                        case "21":
                            pictureBox1.Image = Image.FromFile(Application.StartupPath + @".\21.jpg");
                            break;

                        case "22":
                            pictureBox1.Image = Image.FromFile(Application.StartupPath + @".\22.jpg");
                            break;

                        case "23":
                            pictureBox1.Image = Image.FromFile(Application.StartupPath + @".\23.jpg");
                            break;

                        case "24":
                            pictureBox1.Image = Image.FromFile(Application.StartupPath + @".\24.jpg");
                            break;

                        case "25":
                            pictureBox1.Image = Image.FromFile(Application.StartupPath + @".\25.jpg");
                            break;

                        case "26":
                            pictureBox1.Image = Image.FromFile(Application.StartupPath + @".\26.jpg");
                            break;

                        case "27":
                            pictureBox1.Image = Image.FromFile(Application.StartupPath + @".\27.jpg");
                            break;

                        case "28":
                            pictureBox1.Image = Image.FromFile(Application.StartupPath + @".\28.jpg");
                            break;

                        case "29":
                            pictureBox1.Image = Image.FromFile(Application.StartupPath + @".\29.jpg");
                            break;

                        case "30":
                            pictureBox1.Image = Image.FromFile(Application.StartupPath + @".\30.jpg");
                            break;

                        case "31":
                            pictureBox1.Image = Image.FromFile(Application.StartupPath + @".\31.jpg");
                            break;

                        case "32":
                            pictureBox1.Image = Image.FromFile(Application.StartupPath + @".\32.jpg");
                            break;

                        case "33":
                            pictureBox1.Image = Image.FromFile(Application.StartupPath + @".\33.jpg");
                            break;

                        case "34":
                            pictureBox1.Image = Image.FromFile(Application.StartupPath + @".\34.jpg");
                            break;

                        case "35":
                            pictureBox1.Image = Image.FromFile(Application.StartupPath + @".\35.jpg");
                            break;

                        case "36":
                            pictureBox1.Image = Image.FromFile(Application.StartupPath + @".\36.jpg");
                            break;

                        case "37":
                            pictureBox1.Image = Image.FromFile(Application.StartupPath + @".\37.jpg");
                            break;

                        case "38":
                            pictureBox1.Image = Image.FromFile(Application.StartupPath + @".\38.jpg");
                            break;

                        case "39":
                            pictureBox1.Image = Image.FromFile(Application.StartupPath + @".\39.jpg");
                            break;

                        case "40":
                            pictureBox1.Image = Image.FromFile(Application.StartupPath + @".\40.jpg");
                            break;
                    }
                }
                else if (bAboxof == "15")
                {
                    //載入[ShippingHead]的ListDate
                    using (conn = new SqlConnection(myConnectionString))
                    {
                        conn.Open();

                        selectCmd = "SELECT WhereSeat FROM [ShippingBody] where [ListDate]='" + ListDate_LB.SelectedItem + "' and [ProductName]='" + ProductName_CB.SelectedItem + "'and [WhereBox]='" + WhereBox_LB.SelectedItem + "'  order by Convert(INT,[WhereSeat]) DESC ";
                        cmd = new SqlCommand(selectCmd, conn);
                        using (reader = cmd.ExecuteReader())
                        {
                            if (reader.Read())
                            {
                                switch (reader.GetString(reader.GetOrdinal("WhereSeat")))
                                {
                                    case "1":
                                        pictureBox1.Image = Image.FromFile(Application.StartupPath + @".\15\15-1.jpg");
                                        break;

                                    case "2":
                                        pictureBox1.Image = Image.FromFile(Application.StartupPath + @".\15\15-2.jpg");
                                        break;

                                    case "3":
                                        pictureBox1.Image = Image.FromFile(Application.StartupPath + @".\15\15-3.jpg");
                                        break;

                                    case "4":
                                        pictureBox1.Image = Image.FromFile(Application.StartupPath + @".\15\15-4.jpg");
                                        break;

                                    case "5":
                                        pictureBox1.Image = Image.FromFile(Application.StartupPath + @".\15\15-5.jpg");
                                        break;

                                    case "6":
                                        pictureBox1.Image = Image.FromFile(Application.StartupPath + @".\15\15-6.jpg");
                                        break;

                                    case "7":
                                        pictureBox1.Image = Image.FromFile(Application.StartupPath + @".\15\15-7.jpg");
                                        break;

                                    case "8":
                                        pictureBox1.Image = Image.FromFile(Application.StartupPath + @".\15\15-8.jpg");
                                        break;

                                    case "9":
                                        pictureBox1.Image = Image.FromFile(Application.StartupPath + @".\15\15-9.jpg");
                                        break;

                                    case "10":
                                        pictureBox1.Image = Image.FromFile(Application.StartupPath + @".\15\15-10.jpg");
                                        break;

                                    case "11":
                                        pictureBox1.Image = Image.FromFile(Application.StartupPath + @".\15\15-11.jpg");
                                        break;

                                    case "12":
                                        pictureBox1.Image = Image.FromFile(Application.StartupPath + @".\15\15-12.jpg");
                                        break;

                                    case "13":
                                        pictureBox1.Image = Image.FromFile(Application.StartupPath + @".\15\15-13.jpg");
                                        break;

                                    case "14":
                                        pictureBox1.Image = Image.FromFile(Application.StartupPath + @".\15\15-14.jpg");
                                        break;

                                    case "15":
                                        pictureBox1.Image = Image.FromFile(Application.StartupPath + @".\15\15-15.jpg");
                                        break;
                                }
                            }
                            else
                            {
                                pictureBox1.Image = Image.FromFile(Application.StartupPath + @".\15\15-0.jpg");
                            }
                        }
                    }
                }
                else if (bAboxof == "25")
                {
                    //載入[ShippingHead]的ListDate
                    using (conn = new SqlConnection(myConnectionString))
                    {
                        conn.Open();

                        selectCmd = "SELECT WhereSeat FROM [ShippingBody] where [ListDate]='" + ListDate_LB.SelectedItem + "' and [ProductName]='" + ProductName_CB.SelectedItem + "'and [WhereBox]='" + WhereBox_LB.SelectedItem + "'  order by Convert(INT,[WhereSeat]) DESC ";
                        cmd = new SqlCommand(selectCmd, conn);
                        using (reader = cmd.ExecuteReader())
                        {
                            if (reader.Read())
                            {
                                switch (reader.GetString(reader.GetOrdinal("WhereSeat")))
                                {
                                    case "1":
                                        pictureBox1.Image = Image.FromFile(Application.StartupPath + @".\25\25-1.jpg");
                                        break;

                                    case "2":
                                        pictureBox1.Image = Image.FromFile(Application.StartupPath + @".\25\25-2.jpg");
                                        break;

                                    case "3":
                                        pictureBox1.Image = Image.FromFile(Application.StartupPath + @".\25\25-3.jpg");
                                        break;

                                    case "4":
                                        pictureBox1.Image = Image.FromFile(Application.StartupPath + @".\25\25-4.jpg");
                                        break;

                                    case "5":
                                        pictureBox1.Image = Image.FromFile(Application.StartupPath + @".\25\25-5.jpg");
                                        break;

                                    case "6":
                                        pictureBox1.Image = Image.FromFile(Application.StartupPath + @".\25\25-6.jpg");
                                        break;

                                    case "7":
                                        pictureBox1.Image = Image.FromFile(Application.StartupPath + @".\25\25-7.jpg");
                                        break;

                                    case "8":
                                        pictureBox1.Image = Image.FromFile(Application.StartupPath + @".\25\25-8.jpg");
                                        break;

                                    case "9":
                                        pictureBox1.Image = Image.FromFile(Application.StartupPath + @".\25\25-9.jpg");
                                        break;

                                    case "10":
                                        pictureBox1.Image = Image.FromFile(Application.StartupPath + @".\25\25-10.jpg");
                                        break;

                                    case "11":
                                        pictureBox1.Image = Image.FromFile(Application.StartupPath + @".\25\25-11.jpg");
                                        break;

                                    case "12":
                                        pictureBox1.Image = Image.FromFile(Application.StartupPath + @".\25\25-12.jpg");
                                        break;

                                    case "13":
                                        pictureBox1.Image = Image.FromFile(Application.StartupPath + @".\25\25-13.jpg");
                                        break;

                                    case "14":
                                        pictureBox1.Image = Image.FromFile(Application.StartupPath + @".\25\25-14.jpg");
                                        break;

                                    case "15":
                                        pictureBox1.Image = Image.FromFile(Application.StartupPath + @".\25\25-15.jpg");
                                        break;

                                    case "16":
                                        pictureBox1.Image = Image.FromFile(Application.StartupPath + @".\25\25-16.jpg");
                                        break;

                                    case "17":
                                        pictureBox1.Image = Image.FromFile(Application.StartupPath + @".\25\25-17.jpg");
                                        break;

                                    case "18":
                                        pictureBox1.Image = Image.FromFile(Application.StartupPath + @".\25\25-18.jpg");
                                        break;

                                    case "19":
                                        pictureBox1.Image = Image.FromFile(Application.StartupPath + @".\25\25-19.jpg");
                                        break;

                                    case "20":
                                        pictureBox1.Image = Image.FromFile(Application.StartupPath + @".\25\25-20.jpg");
                                        break;

                                    case "21":
                                        pictureBox1.Image = Image.FromFile(Application.StartupPath + @".\25\25-21.jpg");
                                        break;

                                    case "22":
                                        pictureBox1.Image = Image.FromFile(Application.StartupPath + @".\25\25-23.jpg");
                                        break;

                                    case "23":
                                        pictureBox1.Image = Image.FromFile(Application.StartupPath + @".\25\25-23.jpg");
                                        break;

                                    case "24":
                                        pictureBox1.Image = Image.FromFile(Application.StartupPath + @".\25\25-24.jpg");
                                        break;

                                    case "25":
                                        pictureBox1.Image = Image.FromFile(Application.StartupPath + @".\25\25-25.jpg");
                                        break;
                                }
                            }
                            else
                            {
                                pictureBox1.Image = Image.FromFile(Application.StartupPath + @".\25\25-0.jpg");
                            }
                        }
                    }
                }
                else if (bAboxof == "6")
                {
                    //載入[ShippingHead]的ListDate
                    using (conn = new SqlConnection(myConnectionString))
                    {
                        conn.Open();

                        selectCmd = "SELECT WhereSeat FROM [ShippingBody] where [ListDate]='" + ListDate_LB.SelectedItem + "' and [ProductName]='" + ProductName_CB.SelectedItem + "'and [WhereBox]='" + WhereBox_LB.SelectedItem + "'  order by Convert(INT,[WhereSeat]) DESC ";
                        cmd = new SqlCommand(selectCmd, conn);
                        using (reader = cmd.ExecuteReader())
                        {
                            if (reader.Read())
                            {
                                switch (reader.GetString(reader.GetOrdinal("WhereSeat")))
                                {
                                    case "1":
                                        pictureBox1.Image = Image.FromFile(Application.StartupPath + @".\6\1.jpg");
                                        break;

                                    case "2":
                                        pictureBox1.Image = Image.FromFile(Application.StartupPath + @".\6\2.jpg");
                                        break;

                                    case "3":
                                        pictureBox1.Image = Image.FromFile(Application.StartupPath + @".\6\3.jpg");
                                        break;

                                    case "4":
                                        pictureBox1.Image = Image.FromFile(Application.StartupPath + @".\6\4.jpg");
                                        break;

                                    case "5":
                                        pictureBox1.Image = Image.FromFile(Application.StartupPath + @".\6\5.jpg");
                                        break;

                                    case "6":
                                        pictureBox1.Image = Image.FromFile(Application.StartupPath + @".\6\6.jpg");
                                        break;
                                }
                            }
                            else
                            {
                                pictureBox1.Image = Image.FromFile(Application.StartupPath + @".\6\0.jpg");
                            }
                        }
                    }
                }
                else if (bAboxof == "8")
                {
                    //載入[ShippingHead]的ListDate
                    using (conn = new SqlConnection(myConnectionString))
                    {
                        conn.Open();

                        selectCmd = "SELECT WhereSeat FROM [ShippingBody] where [ListDate]='" + ListDate_LB.SelectedItem + "' and [ProductName]='" + ProductName_CB.SelectedItem + "'and [WhereBox]='" + WhereBox_LB.SelectedItem + "'  order by Convert(INT,[WhereSeat]) DESC ";
                        cmd = new SqlCommand(selectCmd, conn);
                        using (reader = cmd.ExecuteReader())
                        {
                            if (reader.Read())
                            {
                                switch (reader.GetString(reader.GetOrdinal("WhereSeat")))
                                {
                                    case "1":
                                        pictureBox1.Image = Image.FromFile(Application.StartupPath + @".\8\1.jpg");
                                        break;

                                    case "2":
                                        pictureBox1.Image = Image.FromFile(Application.StartupPath + @".\8\2.jpg");
                                        break;

                                    case "3":
                                        pictureBox1.Image = Image.FromFile(Application.StartupPath + @".\8\3.jpg");
                                        break;

                                    case "4":
                                        pictureBox1.Image = Image.FromFile(Application.StartupPath + @".\8\4.jpg");
                                        break;

                                    case "5":
                                        pictureBox1.Image = Image.FromFile(Application.StartupPath + @".\8\5.jpg");
                                        break;

                                    case "6":
                                        pictureBox1.Image = Image.FromFile(Application.StartupPath + @".\8\6.jpg");
                                        break;

                                    case "7":
                                        pictureBox1.Image = Image.FromFile(Application.StartupPath + @".\8\7.jpg");
                                        break;

                                    case "8":
                                        pictureBox1.Image = Image.FromFile(Application.StartupPath + @".\8\8.jpg");
                                        break;
                                }
                            }
                            else
                            {
                                pictureBox1.Image = Image.FromFile(Application.StartupPath + @".\8\0.jpg");
                            }
                        }
                    }
                }
                else if (bAboxof == "12")
                {
                    //載入[ShippingHead]的ListDate
                    using (conn = new SqlConnection(myConnectionString))
                    {
                        conn.Open();

                        selectCmd = "SELECT WhereSeat FROM [ShippingBody] where [ListDate]='" + ListDate_LB.SelectedItem + "' and [ProductName]='" + ProductName_CB.SelectedItem + "'and [WhereBox]='" + WhereBox_LB.SelectedItem + "'  order by Convert(INT,[WhereSeat]) DESC ";
                        cmd = new SqlCommand(selectCmd, conn);
                        using (reader = cmd.ExecuteReader())
                        {
                            if (reader.Read())
                            {
                                switch (reader.GetString(reader.GetOrdinal("WhereSeat")))
                                {
                                    case "1":
                                        pictureBox1.Image = Image.FromFile(Application.StartupPath + @".\12\12-1.jpg");
                                        break;

                                    case "2":
                                        pictureBox1.Image = Image.FromFile(Application.StartupPath + @".\12\12-2.jpg");
                                        break;

                                    case "3":
                                        pictureBox1.Image = Image.FromFile(Application.StartupPath + @".\12\12-3.jpg");
                                        break;

                                    case "4":
                                        pictureBox1.Image = Image.FromFile(Application.StartupPath + @".\12\12-4.jpg");
                                        break;

                                    case "5":
                                        pictureBox1.Image = Image.FromFile(Application.StartupPath + @".\12\12-5.jpg");
                                        break;

                                    case "6":
                                        pictureBox1.Image = Image.FromFile(Application.StartupPath + @".\12\12-6.jpg");
                                        break;

                                    case "7":
                                        pictureBox1.Image = Image.FromFile(Application.StartupPath + @".\12\12-7.jpg");
                                        break;

                                    case "8":
                                        pictureBox1.Image = Image.FromFile(Application.StartupPath + @".\12\12-8.jpg");
                                        break;

                                    case "9":
                                        pictureBox1.Image = Image.FromFile(Application.StartupPath + @".\12\12-9.jpg");
                                        break;

                                    case "10":
                                        pictureBox1.Image = Image.FromFile(Application.StartupPath + @".\12\12-10.jpg");
                                        break;

                                    case "11":
                                        pictureBox1.Image = Image.FromFile(Application.StartupPath + @".\12\12-11.jpg");
                                        break;

                                    case "12":
                                        pictureBox1.Image = Image.FromFile(Application.StartupPath + @".\12\12-12.jpg");
                                        break;
                                }
                            }
                            else
                            {
                                pictureBox1.Image = Image.FromFile(Application.StartupPath + @".\12\12-0.jpg");
                            }
                        }
                    }
                }
                else if (bAboxof == "36")
                {
                    //載入[ShippingHead]的ListDate
                    using (conn = new SqlConnection(myConnectionString))
                    {
                        conn.Open();

                        selectCmd = "SELECT WhereSeat FROM [ShippingBody] where [ListDate]='" + ListDate_LB.SelectedItem + "' and [ProductName]='" + ProductName_CB.SelectedItem + "'and [WhereBox]='" + WhereBox_LB.SelectedItem + "'  order by Convert(INT,[WhereSeat]) DESC ";
                        cmd = new SqlCommand(selectCmd, conn);
                        using (reader = cmd.ExecuteReader())
                        {
                            if (reader.Read())
                            {
                                if (reader.IsDBNull(reader.GetOrdinal("WhereSeat")) == false && (Convert.ToInt32(reader.GetString(reader.GetOrdinal("WhereSeat"))) >= 1 && Convert.ToInt32(reader.GetString(reader.GetOrdinal("WhereSeat"))) <= 117))
                                {
                                    pictureBox1.Image = Image.FromFile(Application.StartupPath + @".\36\36-" + reader.GetString(reader.GetOrdinal("WhereSeat")) + ".jpg");
                                }
                            }
                            else
                            {
                                pictureBox1.Image = Image.FromFile(Application.StartupPath + @".\36\36-0.jpg");
                            }
                        }
                    }
                }
                else if (bAboxof == "117")
                {
                    //載入[ShippingHead]的ListDate
                    using (conn = new SqlConnection(myConnectionString))
                    {
                        conn.Open();

                        selectCmd = "SELECT WhereSeat FROM [ShippingBody] where [ListDate]='" + ListDate_LB.SelectedItem + "' and [ProductName]='" + ProductName_CB.SelectedItem + "'and [WhereBox]='" + WhereBox_LB.SelectedItem + "'  order by Convert(INT,[WhereSeat]) DESC ";
                        cmd = new SqlCommand(selectCmd, conn);
                        using (reader = cmd.ExecuteReader())
                        {
                            if (reader.Read())
                            {
                                if (reader.IsDBNull(reader.GetOrdinal("WhereSeat")) == false && (Convert.ToInt32(reader.GetString(reader.GetOrdinal("WhereSeat"))) >= 1 && Convert.ToInt32(reader.GetString(reader.GetOrdinal("WhereSeat"))) <= 117))
                                {
                                    pictureBox1.Image = Image.FromFile(Application.StartupPath + @".\117\117-" + reader.GetString(reader.GetOrdinal("WhereSeat")) + ".jpg");
                                }
                            }
                            else
                            {
                                pictureBox1.Image = Image.FromFile(Application.StartupPath + @".\117\117-0.jpg");
                            }
                        }
                    }
                }
                else if (bAboxof == "30")
                {
                    //載入[ShippingHead]的ListDate
                    using (conn = new SqlConnection(myConnectionString))
                    {
                        conn.Open();

                        selectCmd = "SELECT WhereSeat FROM [ShippingBody] where [ListDate]='" + ListDate_LB.SelectedItem + "' and [ProductName]='" + ProductName_CB.SelectedItem + "'and [WhereBox]='" + WhereBox_LB.SelectedItem + "'  order by Convert(INT,[WhereSeat]) DESC ";
                        cmd = new SqlCommand(selectCmd, conn);
                        using (reader = cmd.ExecuteReader())
                        {
                            if (reader.Read())
                            {
                                if (reader.IsDBNull(reader.GetOrdinal("WhereSeat")) == false && (Convert.ToInt32(reader.GetString(reader.GetOrdinal("WhereSeat"))) >= 1 && Convert.ToInt32(reader.GetString(reader.GetOrdinal("WhereSeat"))) <= 30))
                                {
                                    pictureBox1.Image = Image.FromFile(Application.StartupPath + @".\30\30-" + reader.GetString(reader.GetOrdinal("WhereSeat")) + ".jpg");
                                }
                            }
                            else
                            {
                                pictureBox1.Image = Image.FromFile(Application.StartupPath + @".\30\30-0.jpg");
                            }
                        }
                    }
                }
                else if (bAboxof == "111")
                {
                    //載入[ShippingHead]的ListDate
                    using (conn = new SqlConnection(myConnectionString))
                    {
                        conn.Open();

                        selectCmd = "SELECT WhereSeat FROM [ShippingBody] where [ListDate]='" + ListDate_LB.SelectedItem + "' and [ProductName]='" + ProductName_CB.SelectedItem + "'and [WhereBox]='" + WhereBox_LB.SelectedItem + "'  order by Convert(INT,[WhereSeat]) DESC ";
                        cmd = new SqlCommand(selectCmd, conn);
                        using (reader = cmd.ExecuteReader())
                        {
                            if (reader.Read())
                            {
                                if (reader.IsDBNull(reader.GetOrdinal("WhereSeat")) == false && (Convert.ToInt32(reader.GetString(reader.GetOrdinal("WhereSeat"))) >= 1 && Convert.ToInt32(reader.GetString(reader.GetOrdinal("WhereSeat"))) <= 111))
                                {
                                    pictureBox1.Image = Image.FromFile(Application.StartupPath + @".\111\111-" + reader.GetString(reader.GetOrdinal("WhereSeat")) + ".jpg");
                                }
                            }
                            else
                            {
                                pictureBox1.Image = Image.FromFile(Application.StartupPath + @".\111\111-0.jpg");
                            }
                        }
                    }
                }
                else if (bAboxof == "4" || bAboxof == "3")
                {
                    //載入[ShippingHead]的ListDate
                    using (conn = new SqlConnection(myConnectionString))
                    {
                        conn.Open();

                        selectCmd = "SELECT WhereSeat FROM [ShippingBody] where [ListDate]='" + ListDate_LB.SelectedItem + "' and [ProductName]='" + ProductName_CB.SelectedItem + "'and [WhereBox]='" + WhereBox_LB.SelectedItem + "'  order by Convert(INT,[WhereSeat]) DESC ";
                        cmd = new SqlCommand(selectCmd, conn);
                        using (reader = cmd.ExecuteReader())
                        {
                            if (reader.Read())
                            {
                                switch (reader.GetString(reader.GetOrdinal("WhereSeat")))
                                {
                                    case "1":
                                        pictureBox1.Image = Image.FromFile(Application.StartupPath + @".\4\4-1.jpg");
                                        break;

                                    case "2":
                                        pictureBox1.Image = Image.FromFile(Application.StartupPath + @".\4\4-2.jpg");
                                        break;

                                    case "3":
                                        pictureBox1.Image = Image.FromFile(Application.StartupPath + @".\4\4-3.jpg");
                                        break;

                                    case "4":
                                        pictureBox1.Image = Image.FromFile(Application.StartupPath + @".\4\4-4.jpg");
                                        break;
                                }
                            }
                            else
                            {
                                pictureBox1.Image = Image.FromFile(Application.StartupPath + @".\4\4-0.jpg");
                            }
                        }
                    }
                }
                else if (bAboxof == "2")
                {
                    //載入[ShippingHead]的ListDate
                    using (conn = new SqlConnection(myConnectionString))
                    {
                        conn.Open();

                        selectCmd = "SELECT WhereSeat FROM [ShippingBody] where [ListDate]='" + ListDate_LB.SelectedItem + "' and [ProductName]='" + ProductName_CB.SelectedItem + "'and [WhereBox]='" + WhereBox_LB.SelectedItem + "'  order by Convert(INT,[WhereSeat]) DESC ";
                        cmd = new SqlCommand(selectCmd, conn);
                        using (reader = cmd.ExecuteReader())
                        {
                            if (reader.Read())
                            {
                                switch (reader.GetString(reader.GetOrdinal("WhereSeat")))
                                {
                                    case "1":
                                        pictureBox1.Image = Image.FromFile(Application.StartupPath + @".\2\2-1.jpg");
                                        break;
                                    case "2":
                                        pictureBox1.Image = Image.FromFile(Application.StartupPath + @".\2\2-2.jpg");
                                        break;
                                }
                            }
                            else
                            {
                                pictureBox1.Image = Image.FromFile(Application.StartupPath + @".\2\2-0.jpg");
                            }
                        }
                    }
                }
            }
            catch
            {
                Thread.Sleep(500);
                LoadPictrue();
            }
        }

        private void StepTimer_Tick(object sender, EventArgs e)
        {
            if (User_LB.Text == "")
            {
                StepLabel1.BackColor = Color.Red;
            }
            else
            {
                StepLabel1.BackColor = Color.MediumTurquoise;
            }

            if (ProductName_CB.Text == "")
            {
                StepLabel2.BackColor = Color.Red;
            }
            else
            {
                StepLabel2.BackColor = Color.MediumTurquoise;
            }

            if (ColorListBox.SelectedIndex == -1)
            {
                StepLabel3.BackColor = Color.Red;
            }
            else
            {
                StepLabel3.BackColor = Color.MediumTurquoise;
            }

            if (ListDate_LB.SelectedIndex == -1)
            {
                StepLabel4.BackColor = Color.Red;
            }
            else
            {
                StepLabel4.BackColor = Color.MediumTurquoise;
            }

            if (WhereBox_LB.SelectedIndex == -1)
            {
                StepLabel5.BackColor = Color.Red;
            }
            else
            {
                StepLabel5.BackColor = Color.MediumTurquoise;
            }

            if (ProductName_CB.Text == "")
            {
                Product_L.Text = "產品名稱：";
            }

            if (WhereBox_LB.SelectedIndex == -1)
            {
                NowBoxsLabel.Text = "目前箱號：";
                ABoxofLabel.Text = "一箱幾隻：";
                PrintLabel.Text = "塗裝漆別";
                AssemblyLabel.Text = "氣瓶配件";
                StorageLabel.Text = "嘜頭狀態：";
                CustomerPO_L.Text = "PO：";
                PalletNoLabel.Text = "棧板號：";

                pictureBox1.Image = null;
            }

            if (ReadyGroupBox.Enabled == false)
            {
                LuckButton.ForeColor = Color.Red;
            }
            else
            {
                LuckButton.ForeColor = Color.Black;
            }
        }

        public string Aboxof()
        {
            string temp = "";

            //載入[ShippingHead]的一箱幾隻
            using (conn = new SqlConnection(myConnectionString))
            {
                conn.Open();

                selectCmd = "SELECT vchAboxof FROM [ShippingHead] where [ListDate]='" + ListDate_LB.SelectedItem + "' and [ProductName]='" + ProductName_CB.SelectedItem + "'and [vchBoxs]='" + WhereBox_LB.SelectedItem + "' ";
                cmd = new SqlCommand(selectCmd, conn);
                using (reader = cmd.ExecuteReader())
                {
                    while (reader.Read())
                    {
                        temp = reader.GetString(reader.GetOrdinal("vchAboxof"));
                    }
                }
            }

            return temp;
        }

        public string APalletof()
        {
            string temp = "";

            //載入[ShippingHead]的棧板編號
            using (conn = new SqlConnection(myConnectionString))
            {
                conn.Open();

                selectCmd = "SELECT isnull(PalletNo,'') PalletNo FROM [ShippingHead] where [ListDate]='" + ListDate_LB.SelectedItem + "' and [ProductName]='" + ProductName_CB.SelectedItem + "'and [vchBoxs]='" + WhereBox_LB.SelectedItem + "' ";
                cmd = new SqlCommand(selectCmd, conn);
                using (reader = cmd.ExecuteReader())
                {
                    if (reader.Read())
                    {
                        temp = reader.GetString(reader.GetOrdinal("PalletNo"));
                    }
                }
            }

            return temp;
        }

        private void LuckButton_Click(object sender, EventArgs e)
        {
            if (User_LB.Text == "")
            {
                MessageBox.Show("尚未選擇測試人員", "警告");
                return;
            }
            else if (ListDate_LB.SelectedIndex == -1)
            {
                MessageBox.Show("尚未選擇嘜頭日期", "警告");
                return;
            }
            else if (ProductName_CB.Text == "")
            {
                MessageBox.Show("尚未選擇嘜頭名稱", "警告");
                return;
            }
            else if (WhereBox_LB.SelectedIndex == -1)
            {
                MessageBox.Show("尚未選擇嘜頭箱號", "警告");
                return;
            }

            if (LinkLMCheckBox.Checked == true)
            {
                DirectionJudgmentTimer.Enabled = true;
            }
            else
            {
                DirectionJudgmentTimer.Enabled = false;
            }

            ReadyGroupBox.Enabled = !ReadyGroupBox.Enabled;
            KeyInGroupBox.Enabled = !KeyInGroupBox.Enabled;
            RefreshhButton.Enabled = !RefreshhButton.Enabled;
            NoLMGroupBox.Enabled = !NoLMGroupBox.Enabled;

            BottleTextBox.Text = "";
            BottomTextBox.Text = "";
        }

        private void DirectionJudgmentTimer_Tick(object sender, EventArgs e)
        {
            string where = "";

            using (conn = new SqlConnection(myConnectionString))
            {
                conn.Open();

                selectCmd = "SELECT [vchWhere] FROM [LaserMarkDirection] ";
                cmd = new SqlCommand(selectCmd, conn);
                using (reader = cmd.ExecuteReader())
                {
                    if (reader.Read())
                    {
                        where = reader.GetString(0);
                    }
                }

                if (where == "0")
                {
                    //提示此序號已經載入罵頭
                    TipTextLabel.Visible = false;

                    BottleTextBox.Focus();
                    BottleLabel.Visible = true;
                    BottomLabel.Visible = false;
                    Direction = "0";
                    BeGin = "Y";
                }
                else if (where == "1")
                {
                    BottomTextBox.Focus();
                    BottleLabel.Visible = false;
                    BottomLabel.Visible = true;
                    Direction = "1";

                    if (BeGin == "N")
                    {
                        BottomTextBox.Text = "";
                    }
                }

                selectCmd = "SELECT [vchWhere] FROM [LaserMarkDirection] ";
                cmd = new SqlCommand(selectCmd, conn);
                using (reader = cmd.ExecuteReader())
                {
                    if (reader.Read())
                    {
                        if (reader.GetString(0) == "2")
                        {
                            Pass = "Y";
                        }
                        else
                        {
                            Pass = "N";
                        }
                    }
                }
            }
        }

        public void LoadSQLDate()
        {
            DataTable DT = new DataTable();
            //載入已放入的氣瓶內容
            dataGridView1.AutoGenerateColumns = false;

            selectCmd = "SELECT [WhereBox] 嘜頭箱號,[WhereSeat] 嘜頭位置,[CylinderNumbers] 氣瓶序號,[CustomerBarCode] 客戶BARCODE,[CylinderWeight] 氣瓶重量 FROM [ShippingBody] " +
                "Where [ListDate] = @ListDate and [ProductName]= @ProductName and [WhereBox] = @WhereBox  order by Convert(INT,[WhereSeat]) asc ";
            sqlAdapter = new SqlDataAdapter(selectCmd, myConnectionString);
            sqlAdapter.SelectCommand.Parameters.AddWithValue("@ListDate", ListDate_LB.SelectedItem);
            sqlAdapter.SelectCommand.Parameters.AddWithValue("@ProductName", ProductName_CB.SelectedItem);
            sqlAdapter.SelectCommand.Parameters.AddWithValue("@WhereBox", WhereBox_LB.SelectedItem);
            sqlAdapter.Fill(DT);

            dataGridView1.DataSource = DT;

            if (dataGridView1.Rows.Count > 0)
            {
                dataGridView1.CurrentCell = dataGridView1.Rows[(dataGridView1.Rows.Count - 1)].Cells[0];
            }
        }

        private void MarkBarCode(string BoxNo)
        {
            Code128 MyCode = new Code128();

            //條碼高度
            MyCode.Height = 50;

            //可見號碼
            MyCode.ValueFont = new Font("細明體", 12, FontStyle.Regular);

            //產生條碼
            Image img = MyCode.GetCodeImage(BoxNo, Code128.Encode.Code128A);

            pictureBox1.Image = img;

            //如果資料匣不在自動新增
            if (!Directory.Exists(@"C:\Code"))
            {
                Directory.CreateDirectory(@"C:\Code");
            }

            string saveQRcode = @"C:\Code\";

            pictureBox1.Image.Save(saveQRcode + BoxNo + ".png");
        }

        private void button2_Click(object sender, EventArgs e)
        {
            if (WhereBox_LB.SelectedIndex == -1)
            {
                MessageBox.Show("請選擇箱號.Select the box number.");
                return;
            }
            //20200921 取消PO限制
            /*
            if (CustomerPO_L.Text.Contains("查無PO"))
            {
                MessageBox.Show("查無PO資料，請聯繫生管");
                return;
            }
            */
            MakeQRCode();
            MarkBarCode(WhereBox_LB.SelectedItem.ToString());

            OutputExcel();
            GC.Collect();
        }

        //EXCEL輸出
        private void OutputExcel()
        {
            //判斷一箱幾隻
            string Aboxof = "", PackingMarks = "", Client = "", DemandNo = string.Empty;
            PalletNoString = "-";

            //判斷一箱幾隻
            using (conn = new SqlConnection(myConnectionString))
            {
                conn.Open();

                selectCmd = "SELECT vchAboxof ,ISNULL(PackingMarks,'') PackingMarks, ISNULL([PalletNo],'-') PalletNo, [Client]" +
                    ", vchAboxof, [DemandNo] FROM [ShippingHead] where [ListDate] = @ListDate AND [ProductName]= @ProductName" +
                    " AND [vchBoxs]= @vchBoxs";
                cmd = new SqlCommand(selectCmd, conn);
                cmd.Parameters.AddWithValue("@ListDate", ListDate_LB.SelectedItem);
                cmd.Parameters.AddWithValue("@ProductName", ProductName_CB.Text);
                cmd.Parameters.AddWithValue("@vchBoxs", WhereBox_LB.SelectedItem);
                using (reader = cmd.ExecuteReader())
                {
                    if (reader.Read())
                    {
                        Aboxof = reader.GetString(reader.GetOrdinal("vchAboxof"));
                        PackingMarks = reader.GetValue(reader.GetOrdinal("PackingMarks")).ToString();
                        PalletNoString = reader.GetValue(reader.GetOrdinal("PalletNo")).ToString();
                        Client = reader.GetValue(reader.GetOrdinal("Client")).ToString();
                        DemandNo = reader.GetValue(reader.GetOrdinal("DemandNo")).ToString();
                    }
                }
            }

            MarkBarCode(PalletNoString);

            //判別是哪個裝箱嘜頭資訊
            if (PackingMarks.ToUpper().Trim().StartsWith("SGA") == true)
            {
                if (Aboxof == "1")
                {
                    //嘜頭表單客製化-SGA  GLADIATAIR
                    Customer_SGA_Form(Aboxof, PackingMarks);
                }
                else
                {//LOGO可能會因客戶有所不同
                    AMS_Form(Aboxof, PackingMarks);
                    //AMS_Form(Aboxof, "SGA-GLADIATAIR");
                }
            }
            else if ((Client.ToUpper().Trim().Contains("ESTRATEGO") == true || Client.ToUpper().Trim().Contains("SC ROBALL") == true) && PackingMarks.ToUpper().Trim().StartsWith("ESTRATEGO") == true)
            {
                //SC ROBALL
                //Estratego
                //僅有LOGO
                Customer_Estratego_Form(Aboxof, PackingMarks);
            }
            else
            {
                //其它-嘜頭表單AMS標準
                //LOGO可能會因客戶有所不同
                AMS_Form(Aboxof, PackingMarks);
            }
            try
            {
                //用來自動跳下一箱
                String BoxsListBoxIndex = WhereBox_LB.SelectedIndex.ToString();
                WhereBox_LB.SelectedIndex = (Convert.ToInt32(BoxsListBoxIndex) + 1);
            }
            catch
            {

            }

            //按完列印FOCUS移到別的地方
            HistoryListBox.Focus();

            //如果不與雷刻程式連線時
            if (LinkLMCheckBox.Checked == false)
            {
                BottleTextBox.Focus();
            }
        }

        private void AMS_Form(string Aboxof, string PackingMarks)
        {
            //公司定義的嘜頭表格
            Excel.Application oXL = new Excel.Application();
            Excel.Workbook oWB;
            Excel.Worksheet oSheet;

            string srcFileName = "";

            if (Aboxof == "20")
            {
                srcFileName = Application.StartupPath + @".\NewListOut.xlsx";//EXCEL檔案路徑
            }
            else if (Aboxof == "40")
            {
                srcFileName = Application.StartupPath + @".\NewListOut40.xlsx";//EXCEL檔案路徑
            }
            else if (Aboxof == "36")
            {
                srcFileName = Application.StartupPath + @".\NewListOut36.xlsx";//EXCEL檔案路徑
            }
            else if (Aboxof == "15")
            {
                srcFileName = Application.StartupPath + @".\NewListOut15.xlsx";//EXCEL檔案路徑
            }
            else if (Aboxof == "16")
            {
                srcFileName = Application.StartupPath + @".\NewListOut16.xlsx";//EXCEL檔案路徑
            }
            else if (Aboxof == "6")
            {
                srcFileName = Application.StartupPath + @".\NewListOut6.xlsx";//EXCEL檔案路徑
            }
            else if (Aboxof == "8")
            {
                srcFileName = Application.StartupPath + @".\NewListOut8.xlsx";//EXCEL檔案路徑
            }
            else if (Aboxof == "10")
            {
                srcFileName = Application.StartupPath + @".\NewListOut10.xlsx";//EXCEL檔案路徑
            }
            else if (Aboxof == "12")
            {
                srcFileName = Application.StartupPath + @".\NewListOut12.xlsx";//EXCEL檔案路徑
            }
            else if (Aboxof == "25")
            {
                srcFileName = Application.StartupPath + @".\NewListOut25.xlsx";//EXCEL檔案路徑
            }
            else if (Aboxof == "30")
            {
                srcFileName = Application.StartupPath + @".\NewListOut30.xlsx";//EXCEL檔案路徑
            }
            else if (Aboxof == "1")
            {
                srcFileName = Application.StartupPath + @".\NewListOut1.xlsx";//EXCEL檔案路徑
            }
            else if (Aboxof == "117")
            {
                srcFileName = Application.StartupPath + @".\NewListOut117.xlsx";//EXCEL檔案路徑
            }
            else if (Aboxof == "4" || Aboxof == "3")
            {
                srcFileName = Application.StartupPath + @".\NewListOut4.xlsx";//EXCEL檔案路徑
            }
            else if (Aboxof == "2")
            {
                srcFileName = Application.StartupPath + @".\NewListOut2.xlsx";//EXCEL檔案路徑
            }
            else if (Aboxof == "111")
            {
                srcFileName = Application.StartupPath + @".\NewListOut111.xlsx";//EXCEL檔案路徑
            }
            else if (Aboxof == "5")
            {
                srcFileName = Application.StartupPath + @".\NewListOut5.xlsx";//EXCEL檔案路徑
            }

            try
            {
                //產生一個Workbook物件，並加入Application//改成.open以及在()中輸入開啟位子
                oWB = oXL.Workbooks.Open(srcFileName, Type.Missing, Type.Missing, Type.Missing, Type.Missing,
                                        Type.Missing, Type.Missing, Type.Missing, Type.Missing,
                                        Type.Missing, Type.Missing, Type.Missing, Type.Missing,
                                        Type.Missing, Type.Missing);
            }
            catch
            {
                MessageBox.Show(@"找不到EXCEL檔案！", "Warning");
                return;
            }

            GetThisBoxMaxCount();

            //設定工作表
            oSheet = (Excel.Worksheet)oWB.ActiveSheet;

            //插入1維條碼
            //預設位子在X:343,Y:396
            //預設1維條碼圖片大小200*30            
            int oneY = 427;
            string oneadd = @"C:\Code\";

            using (conn = new SqlConnection(myConnectionString))
            {
                conn.Open();

                selectCmd = "SELECT  * FROM [ShippingHead] where [ListDate]='" + ListDate_LB.SelectedItem + "' and [ProductName]='" + ProductName_CB.Text + "' and [vchBoxs]='" + WhereBox_LB.SelectedItem + "'";
                cmd = new SqlCommand(selectCmd, conn);
                reader = cmd.ExecuteReader();
                while (reader.Read())
                {
                    Excel.Worksheet xSheet = (Excel.Worksheet)oWB.Sheets[1];
                    if (Aboxof == "8")
                    {
                        oSheet.Shapes.AddPicture(oneadd + WhereBox_LB.SelectedItem + ".png", Microsoft.Office.Core.MsoTriState.msoFalse,
                                            Microsoft.Office.Core.MsoTriState.msoTrue, 252, oneY, 200, 30);
                        oSheet.Shapes.AddPicture(oneadd + PalletNoString + ".png", Microsoft.Office.Core.MsoTriState.msoFalse,
                        Microsoft.Office.Core.MsoTriState.msoTrue, 704, oneY, 200, 30);
                    }
                    else if (Aboxof == "6")
                    {
                        oSheet.Shapes.AddPicture(oneadd + WhereBox_LB.SelectedItem + ".png", Microsoft.Office.Core.MsoTriState.msoFalse,
                                            Microsoft.Office.Core.MsoTriState.msoTrue, 252, oneY, 200, 30);
                        oSheet.Shapes.AddPicture(oneadd + PalletNoString + ".png", Microsoft.Office.Core.MsoTriState.msoFalse,
                        Microsoft.Office.Core.MsoTriState.msoTrue, 704, oneY, 200, 30);
                    }
                    else if (Aboxof == "16")
                    {
                        oSheet.Shapes.AddPicture(oneadd + WhereBox_LB.SelectedItem + ".png", Microsoft.Office.Core.MsoTriState.msoFalse,
                                            Microsoft.Office.Core.MsoTriState.msoTrue, 252, oneY, 200, 30);
                        oSheet.Shapes.AddPicture(oneadd + PalletNoString + ".png", Microsoft.Office.Core.MsoTriState.msoFalse,
                        Microsoft.Office.Core.MsoTriState.msoTrue, 704, oneY, 200, 30);
                    }
                    else if (Aboxof == "10")
                    {
                        oSheet.Shapes.AddPicture(oneadd + WhereBox_LB.SelectedItem + ".png", Microsoft.Office.Core.MsoTriState.msoFalse,
                                            Microsoft.Office.Core.MsoTriState.msoTrue, 222, oneY, 200, 30);
                        oSheet.Shapes.AddPicture(oneadd + PalletNoString + ".png", Microsoft.Office.Core.MsoTriState.msoFalse,
                        Microsoft.Office.Core.MsoTriState.msoTrue, 704, oneY, 200, 30);
                    }
                    else if (Aboxof == "20")
                    {
                        oSheet.Shapes.AddPicture(oneadd + WhereBox_LB.SelectedItem + ".png", Microsoft.Office.Core.MsoTriState.msoFalse,
                        Microsoft.Office.Core.MsoTriState.msoTrue, 265, oneY, 200, 30);
                        //PalletNoString
                        oSheet.Shapes.AddPicture(oneadd + PalletNoString + ".png", Microsoft.Office.Core.MsoTriState.msoFalse,
                        Microsoft.Office.Core.MsoTriState.msoTrue, 754, oneY, 200, 30);
                    }
                    else if (Aboxof == "40")
                    {
                        oSheet.Shapes.AddPicture(oneadd + WhereBox_LB.SelectedItem + ".png", Microsoft.Office.Core.MsoTriState.msoFalse,
                        Microsoft.Office.Core.MsoTriState.msoTrue, 220, oneY, 200, 30);
                        oSheet.Shapes.AddPicture(oneadd + PalletNoString + ".png", Microsoft.Office.Core.MsoTriState.msoFalse,
                        Microsoft.Office.Core.MsoTriState.msoTrue, 704, oneY, 200, 30);
                    }
                    else if (Aboxof == "36")
                    {
                        oSheet.Shapes.AddPicture(oneadd + WhereBox_LB.SelectedItem + ".png", Microsoft.Office.Core.MsoTriState.msoFalse,
                        Microsoft.Office.Core.MsoTriState.msoTrue, 240, oneY, 200, 30);
                        oSheet.Shapes.AddPicture(oneadd + PalletNoString + ".png", Microsoft.Office.Core.MsoTriState.msoFalse,
                        Microsoft.Office.Core.MsoTriState.msoTrue, 774, oneY, 200, 30);
                    }
                    else if (Aboxof == "25")
                    {
                        oSheet.Shapes.AddPicture(oneadd + WhereBox_LB.SelectedItem + ".png", Microsoft.Office.Core.MsoTriState.msoFalse,
                        Microsoft.Office.Core.MsoTriState.msoTrue, 268, oneY, 200, 30);
                        oSheet.Shapes.AddPicture(oneadd + PalletNoString + ".png", Microsoft.Office.Core.MsoTriState.msoFalse,
                        Microsoft.Office.Core.MsoTriState.msoTrue, 754, oneY, 200, 30);
                    }
                    else if (Aboxof == "30")
                    {
                        oSheet.Shapes.AddPicture(oneadd + WhereBox_LB.SelectedItem + ".png", Microsoft.Office.Core.MsoTriState.msoFalse,
                        Microsoft.Office.Core.MsoTriState.msoTrue, 268, 431, 200, 30);
                        oSheet.Shapes.AddPicture(oneadd + PalletNoString + ".png", Microsoft.Office.Core.MsoTriState.msoFalse,
                        Microsoft.Office.Core.MsoTriState.msoTrue, 754, 430, 200, 30);
                    }
                    else if (Aboxof == "15")
                    {
                        oSheet.Shapes.AddPicture(oneadd + WhereBox_LB.SelectedItem + ".png", Microsoft.Office.Core.MsoTriState.msoFalse,
                        Microsoft.Office.Core.MsoTriState.msoTrue, 260, oneY, 200, 30);
                        oSheet.Shapes.AddPicture(oneadd + PalletNoString + ".png", Microsoft.Office.Core.MsoTriState.msoFalse,
                        Microsoft.Office.Core.MsoTriState.msoTrue, 754, oneY, 200, 30);
                    }
                    else if (Aboxof == "12")
                    {
                        oSheet.Shapes.AddPicture(oneadd + WhereBox_LB.SelectedItem + ".png", Microsoft.Office.Core.MsoTriState.msoFalse,
                        Microsoft.Office.Core.MsoTriState.msoTrue, 260, oneY, 200, 30);
                        oSheet.Shapes.AddPicture(oneadd + PalletNoString + ".png", Microsoft.Office.Core.MsoTriState.msoFalse,
                        Microsoft.Office.Core.MsoTriState.msoTrue, 704, oneY, 200, 30);
                    }
                    else if (Aboxof == "117")
                    {
                        oSheet.Shapes.AddPicture(oneadd + WhereBox_LB.SelectedItem + ".png", Microsoft.Office.Core.MsoTriState.msoFalse,
                        Microsoft.Office.Core.MsoTriState.msoTrue, 200, 587, 200, 30);
                        oSheet.Shapes.AddPicture(oneadd + PalletNoString + ".png", Microsoft.Office.Core.MsoTriState.msoFalse,
                        Microsoft.Office.Core.MsoTriState.msoTrue, 620, 587, 130, 30);
                    }
                    else if (Aboxof == "111")
                    {
                        oSheet.Shapes.AddPicture(oneadd + WhereBox_LB.SelectedItem + ".png", Microsoft.Office.Core.MsoTriState.msoFalse,
                        Microsoft.Office.Core.MsoTriState.msoTrue, 200, 587, 200, 30);
                        oSheet.Shapes.AddPicture(oneadd + PalletNoString + ".png", Microsoft.Office.Core.MsoTriState.msoFalse,
                        Microsoft.Office.Core.MsoTriState.msoTrue, 620, 587, 130, 30);
                    }
                    else if (Aboxof == "4" || Aboxof == "3")
                    {
                        oSheet.Shapes.AddPicture(oneadd + WhereBox_LB.SelectedItem + ".png", Microsoft.Office.Core.MsoTriState.msoFalse,
                                            Microsoft.Office.Core.MsoTriState.msoTrue, 256, oneY, 200, 30);
                        oSheet.Shapes.AddPicture(oneadd + PalletNoString + ".png", Microsoft.Office.Core.MsoTriState.msoFalse,
                       Microsoft.Office.Core.MsoTriState.msoTrue, 704, oneY, 200, 30);
                    }
                    else if (Aboxof == "2")
                    {
                        oSheet.Shapes.AddPicture(oneadd + WhereBox_LB.SelectedItem + ".png", Microsoft.Office.Core.MsoTriState.msoFalse,
                                            Microsoft.Office.Core.MsoTriState.msoTrue, 256, oneY, 200, 30);
                        oSheet.Shapes.AddPicture(oneadd + PalletNoString + ".png", Microsoft.Office.Core.MsoTriState.msoFalse,
                      Microsoft.Office.Core.MsoTriState.msoTrue, 704, oneY, 200, 30);
                    }
                    else if (Aboxof == "1")
                    {
                        oSheet.Shapes.AddPicture(oneadd + WhereBox_LB.SelectedItem + ".png", Microsoft.Office.Core.MsoTriState.msoFalse,
                                            Microsoft.Office.Core.MsoTriState.msoTrue, 256, oneY, 200, 30);
                    }
                    else if (Aboxof == "5")
                    {
                        oSheet.Shapes.AddPicture(oneadd + WhereBox_LB.SelectedItem + ".png", Microsoft.Office.Core.MsoTriState.msoFalse,
                                            Microsoft.Office.Core.MsoTriState.msoTrue, 256, oneY, 200, 30);
                    }
                }
            }

            string Client = "";

            if (Aboxof == "20")
            {
                string HowMuch = "";
                int Cumulative = 0;
                int Total = 0;

                //載入嘜頭資料
                using (conn = new SqlConnection(myConnectionString))
                {
                    conn.Open();

                    selectCmd = "SELECT vchAboxof FROM [ShippingHead] where [ListDate]='" + ListDate_LB.SelectedItem + "' and [ProductName]='" + ProductName_CB.SelectedItem + "'";
                    cmd = new SqlCommand(selectCmd, conn);
                    using (reader = cmd.ExecuteReader())
                    {
                        while (reader.Read())
                        {
                            HowMuch = reader.GetString(reader.GetOrdinal("vchAboxof"));
                            Cumulative++;
                        }
                    }

                    Total = Convert.ToInt32(HowMuch) * Cumulative;

                    //載入嘜頭資料

                    selectCmd = "SELECT  Client, (vchPrint + ' ' + ProductName + ' ' + Marking) PartDescription, isnull(CustomerProductNo,'') CustomerProductNo, vchBoxs, isnull(PalletNo,'') PalletNo " +
                        "FROM [ShippingHead] where [ListDate]='" + ListDate_LB.SelectedItem + "' and [ProductName]='" + ProductName_CB.SelectedItem + "'and [vchBoxs]='" + WhereBox_LB.SelectedItem + "'";
                    cmd = new SqlCommand(selectCmd, conn);
                    using (reader = cmd.ExecuteReader())
                    {
                        if (reader.Read())
                        {
                            Client = reader.GetString(reader.GetOrdinal("Client")).Trim();

                            //載入客戶產品名稱
                            oSheet.Cells[1, 7] = reader.GetString(reader.GetOrdinal("PartDescription"));

                            //載入客戶產品型號
                            oSheet.Cells[2, 7] = reader.GetString(reader.GetOrdinal("CustomerProductNo"));

                            //載入一箱幾隻
                            oSheet.Cells[4, 7] = Getcount;

                            //載入箱號
                            oSheet.Cells[10, 2] = reader.GetString(reader.GetOrdinal("vchBoxs"));

                            //載入客戶名稱
                            oSheet.Cells[3, 7] = reader.GetString(reader.GetOrdinal("Client"));

                            //載入箱號
                            oSheet.Cells[10, 10] = reader.GetString(reader.GetOrdinal("PalletNo"));

                            //20200410 加入PO
                            oSheet.Cells[5, 11] = CustomerPO_L.Text;

                            //該客戶要其自己的logo
                            if (reader.GetString(reader.GetOrdinal("Client")).Trim().CompareTo("Wicked Sportz") == 0)
                            {
                                oSheet.Shapes.AddPicture(Application.StartupPath + @".\LOGO-ENAIRGY_Wicked Sportz.jpg", Microsoft.Office.Core.MsoTriState.msoFalse,
                                                Microsoft.Office.Core.MsoTriState.msoTrue, 2, 17, 212, 125);
                            }
                            if (reader.GetString(reader.GetOrdinal("Client")).Trim().CompareTo("達成數位") == 0)
                            {
                                oSheet.Shapes.AddPicture(Application.StartupPath + @".\LOGO_DCT.jpg", Microsoft.Office.Core.MsoTriState.msoFalse,
                                                Microsoft.Office.Core.MsoTriState.msoTrue, 2, 17, 212, 125);
                            }
                            else if (Client.ToUpper().StartsWith("EMB"))
                            {
                                oSheet.Shapes.AddPicture(Application.StartupPath + @".\LOGO_EMB.png", Microsoft.Office.Core.MsoTriState.msoFalse,
                                                    Microsoft.Office.Core.MsoTriState.msoTrue, 2, 17, 212, 125);
                            }
                            else if (Client.ToUpper().StartsWith("PAINTBALL SPORTS"))
                            {
                                oSheet.Shapes.AddPicture(Application.StartupPath + @".\LOGO_Paintball Sports GmbH.png", Microsoft.Office.Core.MsoTriState.msoFalse,
                                                    Microsoft.Office.Core.MsoTriState.msoTrue, 2, 17, 212, 125);
                            }
                        }
                    }
                }

                int serialnooneX = 7, serialnooneY = 205;
                string serialnooneadd = @"C:\SerialNoCode\";
                string FirstCNO = "";

                //載入嘜頭氣瓶序號位子
                using (conn = new SqlConnection(myConnectionString))
                {
                    conn.Open();

                    selectCmd = "SELECT WhereSeat, CylinderNumbers FROM [ShippingBody] where [ListDate]='" + ListDate_LB.SelectedItem + "' and [ProductName]='" + ProductName_CB.SelectedItem + "'and [WhereBox]='" + WhereBox_LB.SelectedItem + "'";
                    cmd = new SqlCommand(selectCmd, conn);
                    using (reader = cmd.ExecuteReader())
                    {
                        while (reader.Read())
                        {
                            serialnooneX = 3; serialnooneY = 203;

                            switch (reader.GetString(reader.GetOrdinal("WhereSeat")))
                            {
                                case "1":
                                    oSheet.Cells[6, 1] = reader.GetString(reader.GetOrdinal("CylinderNumbers"));
                                    FirstCNO = reader.GetString(reader.GetOrdinal("CylinderNumbers"));
                                    break;

                                case "2":
                                    oSheet.Cells[6, 3] = reader.GetString(reader.GetOrdinal("CylinderNumbers"));
                                    break;

                                case "3":
                                    oSheet.Cells[6, 5] = reader.GetString(reader.GetOrdinal("CylinderNumbers"));
                                    break;

                                case "4":
                                    oSheet.Cells[6, 7] = reader.GetString(reader.GetOrdinal("CylinderNumbers"));
                                    break;

                                case "5":
                                    oSheet.Cells[6, 9] = reader.GetString(reader.GetOrdinal("CylinderNumbers"));
                                    break;

                                case "6":
                                    oSheet.Cells[7, 1] = reader.GetString(reader.GetOrdinal("CylinderNumbers"));
                                    break;

                                case "7":
                                    oSheet.Cells[7, 3] = reader.GetString(reader.GetOrdinal("CylinderNumbers"));
                                    break;

                                case "8":
                                    oSheet.Cells[7, 5] = reader.GetString(reader.GetOrdinal("CylinderNumbers"));
                                    break;

                                case "9":
                                    oSheet.Cells[7, 7] = reader.GetString(reader.GetOrdinal("CylinderNumbers"));
                                    break;

                                case "10":
                                    oSheet.Cells[7, 9] = reader.GetString(reader.GetOrdinal("CylinderNumbers"));
                                    break;

                                case "11":
                                    oSheet.Cells[8, 1] = reader.GetString(reader.GetOrdinal("CylinderNumbers"));
                                    break;

                                case "12":
                                    oSheet.Cells[8, 3] = reader.GetString(reader.GetOrdinal("CylinderNumbers"));
                                    break;

                                case "13":
                                    oSheet.Cells[8, 5] = reader.GetString(reader.GetOrdinal("CylinderNumbers"));
                                    break;

                                case "14":
                                    oSheet.Cells[8, 7] = reader.GetString(reader.GetOrdinal("CylinderNumbers"));
                                    break;

                                case "15":
                                    oSheet.Cells[8, 9] = reader.GetString(reader.GetOrdinal("CylinderNumbers"));
                                    break;

                                case "16":
                                    oSheet.Cells[9, 1] = reader.GetString(reader.GetOrdinal("CylinderNumbers"));
                                    break;

                                case "17":
                                    oSheet.Cells[9, 3] = reader.GetString(reader.GetOrdinal("CylinderNumbers"));
                                    break;

                                case "18":
                                    oSheet.Cells[9, 5] = reader.GetString(reader.GetOrdinal("CylinderNumbers"));
                                    break;

                                case "19":
                                    oSheet.Cells[9, 7] = reader.GetString(reader.GetOrdinal("CylinderNumbers"));
                                    break;

                                case "20":
                                    oSheet.Cells[9, 9] = reader.GetString(reader.GetOrdinal("CylinderNumbers"));
                                    break;
                            }

                            serialnooneX = serialnooneX + ((Convert.ToInt32(reader.GetString(reader.GetOrdinal("WhereSeat"))) + 4) % 5) * 145;
                            serialnooneY = serialnooneY + ((Convert.ToInt32(reader.GetString(reader.GetOrdinal("WhereSeat"))) - 1) / 5) * 56;
                            oSheet.Shapes.AddPicture(serialnooneadd + reader.GetString(reader.GetOrdinal("CylinderNumbers")) + ".png", Microsoft.Office.Core.MsoTriState.msoFalse,
                            Microsoft.Office.Core.MsoTriState.msoTrue, serialnooneX, serialnooneY, 44, 44);// 130, 22);
                        }
                    }

                    if ((Client.Contains("Scientific Gas Australia Pty Ltd") || Client.Contains("Airtanks")) && PackingMarks.Trim().CompareTo("SGA-GLADIATAIR") == 0)
                    {
                        string ProductNO = "";

                        //該客戶要其自己的logo  PartNo   Part Description
                        selectCmd = "SELECT  Product_NO FROM MSNBody,Manufacturing where [CylinderNo]='" + FirstCNO + "' and vchManufacturingNo=  Manufacturing_NO";
                        cmd = new SqlCommand(selectCmd, conn);
                        using (reader = cmd.ExecuteReader())
                        {
                            if (reader.Read())
                            {
                                ProductNO = reader.GetValue(0).ToString();
                            }
                        }

                        selectCmd = "SELECT  ProductCode, ProductDescription FROM CustomerPackingMark " +
                            "where ProductNo='" + ProductNO + "' and LogoType='" + (PackingMarks.Trim().Contains("-") == true ? PackingMarks.Trim().Split('-')[1].Trim().ToUpper() : PackingMarks.Trim()) + "'";
                        cmd = new SqlCommand(selectCmd, conn);
                        using (reader = cmd.ExecuteReader())
                        {
                            if (reader.Read())
                            {
                                //載入客戶產品名稱
                                oSheet.Cells[1, 7] = reader.GetString(reader.GetOrdinal("ProductDescription"));

                                //載入客戶產品型號
                                oSheet.Cells[2, 7] = reader.GetString(reader.GetOrdinal("ProductCode"));
                            }
                        }

                        //LOGO
                        oSheet.Shapes.AddPicture(Application.StartupPath + @".\LOGO_SGA_GLADIATAIR.png", Microsoft.Office.Core.MsoTriState.msoFalse,
                                            Microsoft.Office.Core.MsoTriState.msoTrue, 2, 17, 212, 125);
                    }
                    else if (Client.ToUpper().StartsWith("EMB"))
                    {
                        oSheet.Shapes.AddPicture(Application.StartupPath + @".\LOGO_EMB.png", Microsoft.Office.Core.MsoTriState.msoFalse,
                                            Microsoft.Office.Core.MsoTriState.msoTrue, 2, 17, 212, 125);
                    }
                    else if ((Client.ToUpper().StartsWith("HATSAN") == true) && PackingMarks.Trim().CompareTo("HATSAN") == 0)
                    {
                        //LOGO
                        oSheet.Shapes.AddPicture(Application.StartupPath + @".\LOGO_HATSAN.png", Microsoft.Office.Core.MsoTriState.msoFalse,
                                            Microsoft.Office.Core.MsoTriState.msoTrue, 2, 17, 212, 125);
                    }
                    else if (Client.ToUpper().Contains("ADRENALICIA S.L."))
                    {
                        oSheet.Shapes.AddPicture(Application.StartupPath + @".\LOGO_RogerSports.png", Microsoft.Office.Core.MsoTriState.msoFalse,
                    Microsoft.Office.Core.MsoTriState.msoTrue, 2, 17, 212, 125);
                    }
                    else if (Client.ToUpper().StartsWith("AIR TEC") == true)
                    {
                        //20190314 AIR TEC 1.55L 增加Country of Origin : Taiwan 字樣
                        //增加Country of Origin : Taiwan 字樣
                        oSheet.Cells[4, 11] = "COO：";
                        oSheet.Cells[4, 13] = "Taiwan";

                        //加框
                        Excel.Range excelRange = oSheet.get_Range(oSheet.Cells[4, 11], oSheet.Cells[4, 13]);
                        excelRange.Borders.LineStyle = Excel.XlLineStyle.xlContinuous;
                        excelRange.Borders.get_Item(Excel.XlBordersIndex.xlInsideVertical).Weight = Excel.XlBorderWeight.xlMedium;
                        excelRange.Borders.get_Item(Excel.XlBordersIndex.xlEdgeLeft).Weight = Excel.XlBorderWeight.xlMedium;
                        excelRange.Borders.get_Item(Excel.XlBordersIndex.xlEdgeRight).Weight = Excel.XlBorderWeight.xlMedium;
                    }

                    else if (Client.ToUpper().StartsWith("PAINTBALL SPORTS"))
                    {
                        oSheet.Shapes.AddPicture(Application.StartupPath + @".\LOGO_Paintball Sports GmbH.png", Microsoft.Office.Core.MsoTriState.msoFalse,
                                            Microsoft.Office.Core.MsoTriState.msoTrue, 2, 17, 212, 125);
                    }

                    //if (StorageStatus == "N")//20190212
                    {
                        //預設位子在X:680,Y:155
                        //預設QRCODE圖片大小250*250

                        //插入圖片
                        int picX = 730, picY = 185;
                        string picadd = @"C:\QRCode\";

                        selectCmd = "SELECT ListDate, ProductName, vchBoxs FROM [ShippingHead] where [ListDate]='" + ListDate_LB.SelectedItem + "' and [ProductName]='" + ProductName_CB.Text + "' and [vchBoxs]='" + WhereBox_LB.SelectedItem + "'";
                        cmd = new SqlCommand(selectCmd, conn);
                        using (reader = cmd.ExecuteReader())
                        {
                            while (reader.Read())
                            {
                                Excel.Worksheet xSheet = (Excel.Worksheet)oWB.Sheets[1];

                                oSheet.Shapes.AddPicture(picadd + reader.GetString(reader.GetOrdinal("ListDate")) + reader.GetString(reader.GetOrdinal("ProductName")) + reader.GetString(reader.GetOrdinal("vchBoxs")) + ".png", Microsoft.Office.Core.MsoTriState.msoFalse, Microsoft.Office.Core.MsoTriState.msoTrue, picX, picY, 250, 250);

                                if (picX == 885)
                                {
                                    picY += 70;
                                    picX = 125;
                                }
                                else
                                {
                                    picX += 190;
                                }
                            }
                        }
                    }
                }
            }
            else if (Aboxof == "36")
            {
                //載入嘜頭資料
                using (conn = new SqlConnection(myConnectionString))
                {
                    conn.Open();

                    selectCmd = "SELECT  [Client], (vchPrint + ' ' + ProductName + ' ' + Marking) PartDescription, isnull([CustomerProductName],''),isnull([CustomerProductNo],''),[vchBoxs],isnull(PalletNo,'') FROM [ShippingHead] where [ListDate]='" + ListDate_LB.SelectedItem + "' and [ProductName]='" + ProductName_CB.SelectedItem + "'and [vchBoxs]='" + WhereBox_LB.SelectedItem + "'";
                    cmd = new SqlCommand(selectCmd, conn);
                    using (reader = cmd.ExecuteReader())
                    {
                        while (reader.Read())
                        {
                            Client = reader.GetString(0).Trim();
                            //載入客戶產品名稱
                            oSheet.Cells[1, 8] = reader.GetString(2);

                            //載入客戶產品型號
                            oSheet.Cells[2, 8] = reader.GetString(3);

                            //載入一箱幾隻
                            oSheet.Cells[4, 8] = Getcount;

                            //載入箱號
                            oSheet.Cells[12, 2] = reader.GetString(4);

                            //20200410 加入PO
                            oSheet.Cells[5, 13] = CustomerPO_L.Text;

                            //載入客戶名稱
                            oSheet.Cells[3, 8] = reader.GetString(0);

                            //載入棧板號
                            oSheet.Cells[12, 11] = reader.GetString(5);

                            if (reader.GetString(0).Trim().CompareTo("Wicked Sportz") == 0)
                            {
                                //該客戶要其自己的logo
                                oSheet.Shapes.AddPicture(Application.StartupPath + @".\LOGO-ENAIRGY_Wicked Sportz.jpg", Microsoft.Office.Core.MsoTriState.msoFalse,
                                                Microsoft.Office.Core.MsoTriState.msoTrue, 16, 17, 212, 125);
                                // Application.StartupPath + @".\LOGO-ENAIRGY_Wicked Sportz.jpg"
                            }
                            if (reader.GetString(0).Trim().CompareTo("達成數位") == 0)
                            {
                                //該客戶要其自己的logo
                                oSheet.Shapes.AddPicture(Application.StartupPath + @".\LOGO_DCT.jpg", Microsoft.Office.Core.MsoTriState.msoFalse,
                                                Microsoft.Office.Core.MsoTriState.msoTrue, 16, 17, 212, 125);
                                // Application.StartupPath + @".\LOGO-ENAIRGY_Wicked Sportz.jpg"
                            }
                            else if (reader.GetString(reader.GetOrdinal("Client")).Trim().Contains("ADRENALICIA S.L."))
                            {
                                oSheet.Shapes.AddPicture(Application.StartupPath + @".\LOGO_RogerSports.png", Microsoft.Office.Core.MsoTriState.msoFalse,
                            Microsoft.Office.Core.MsoTriState.msoTrue, 2, 17, 212, 125);
                            }
                            else if (Client.ToUpper().StartsWith("EMB"))
                            {
                                oSheet.Shapes.AddPicture(Application.StartupPath + @".\LOGO_EMB.png", Microsoft.Office.Core.MsoTriState.msoFalse,
                                                    Microsoft.Office.Core.MsoTriState.msoTrue, 2, 17, 212, 125);
                            }

                            else if (Client.ToUpper().StartsWith("PAINTBALL SPORTS"))
                            {
                                oSheet.Shapes.AddPicture(Application.StartupPath + @".\LOGO_Paintball Sports GmbH.png", Microsoft.Office.Core.MsoTriState.msoFalse,
                                                    Microsoft.Office.Core.MsoTriState.msoTrue, 2, 17, 212, 125);
                            }
                        }
                    }
                }

                string FirstCNO = "";

                //載入嘜頭氣瓶序號位子
                using (conn = new SqlConnection(myConnectionString))
                {
                    conn.Open();

                    selectCmd = "SELECT  * FROM [ShippingBody] where [ListDate]='" + ListDate_LB.SelectedItem + "' and [ProductName]='" + ProductName_CB.SelectedItem + "'and [WhereBox]='" + WhereBox_LB.SelectedItem + "'";
                    cmd = new SqlCommand(selectCmd, conn);
                    using (reader = cmd.ExecuteReader())
                    {
                        while (reader.Read())
                        {
                            if (FirstCNO == "")
                            {
                                FirstCNO = reader.GetString(3);
                            }
                            if (reader.IsDBNull(5) == false && (Convert.ToInt32(reader.GetString(5)) >= 1 && Convert.ToInt32(reader.GetString(5)) <= 36))
                            {

                                oSheet.Cells[6 + ((Convert.ToInt32(reader.GetString(5)) - 1) / 6), 1 + ((Convert.ToInt32(reader.GetString(5)) - 1) % 6) * 2] = reader.GetString(3);

                            }
                        }
                    }

                    if ((Client.Contains("Scientific Gas Australia Pty Ltd") || Client == "Airtanks Limited") && PackingMarks.Trim().CompareTo("SGA-GLADIATAIR") == 0)
                    {
                        string ProductNO = "";

                        //該客戶要其自己的logo  PartNo   Part Description
                        selectCmd = "SELECT  Product_NO FROM MSNBody,Manufacturing where [CylinderNo]='" + FirstCNO + "' and vchManufacturingNo=  Manufacturing_NO";
                        cmd = new SqlCommand(selectCmd, conn);
                        using (reader = cmd.ExecuteReader())
                        {
                            if (reader.Read())
                            {
                                ProductNO = reader.GetValue(0).ToString();
                            }
                        }

                        selectCmd = "SELECT  ProductCode, ProductDescription FROM CustomerPackingMark where ProductNo='" + ProductNO + "' and LogoType='" + (PackingMarks.Trim().Contains("-") == true ? PackingMarks.Trim().Split('-')[1].Trim().ToUpper() : PackingMarks.Trim()) + "'";
                        cmd = new SqlCommand(selectCmd, conn);
                        using (reader = cmd.ExecuteReader())
                        {
                            if (reader.Read())
                            {
                                //載入客戶產品名稱
                                oSheet.Cells[1, 8] = reader.GetString(1);

                                //載入客戶產品型號
                                oSheet.Cells[2, 8] = reader.GetString(0);
                            }
                        }

                        //LOGO
                        oSheet.Shapes.AddPicture(Application.StartupPath + @".\LOGO_SGA_GLADIATAIR.png", Microsoft.Office.Core.MsoTriState.msoFalse,
                                            Microsoft.Office.Core.MsoTriState.msoTrue, 2, 17, 212, 125);
                    }
                    else if ((Client.ToUpper().StartsWith("HATSAN") == true) && PackingMarks.Trim().CompareTo("HATSAN") == 0)
                    {
                        //LOGO
                        oSheet.Shapes.AddPicture(Application.StartupPath + @".\LOGO_HATSAN.png", Microsoft.Office.Core.MsoTriState.msoFalse,
                                            Microsoft.Office.Core.MsoTriState.msoTrue, 2, 17, 212, 125);
                    }
                    else if (Client.ToUpper().StartsWith("EMB"))
                    {
                        oSheet.Shapes.AddPicture(Application.StartupPath + @".\LOGO_EMB.png", Microsoft.Office.Core.MsoTriState.msoFalse,
                                            Microsoft.Office.Core.MsoTriState.msoTrue, 2, 17, 212, 125);
                    }
                    else if (Client.ToUpper().StartsWith("PAINTBALL SPORTS"))
                    {
                        oSheet.Shapes.AddPicture(Application.StartupPath + @".\LOGO_Paintball Sports GmbH.png", Microsoft.Office.Core.MsoTriState.msoFalse,
                                            Microsoft.Office.Core.MsoTriState.msoTrue, 2, 17, 212, 125);
                    }

                    //if (StorageStatus == "N")//20190212
                    {
                        //預設位子在X:680,Y:155
                        //預設QRCODE圖片大小250*250

                        //插入二維條碼
                        int picX = 750, picY = 179;
                        string picadd = @"C:\QRCode\";

                        selectCmd = "SELECT  * FROM [ShippingHead] where [ListDate]='" + ListDate_LB.SelectedItem + "' and [ProductName]='" + ProductName_CB.Text + "' and [vchBoxs]='" + WhereBox_LB.SelectedItem + "'";
                        cmd = new SqlCommand(selectCmd, conn);
                        using (reader = cmd.ExecuteReader())
                        {
                            while (reader.Read())
                            {
                                Excel.Worksheet xSheet = (Excel.Worksheet)oWB.Sheets[1];
                                oSheet.Shapes.AddPicture(picadd + (reader.GetString(0) + reader.GetString(1) + reader.GetString(3)) + ".png", Microsoft.Office.Core.MsoTriState.msoFalse,
                                Microsoft.Office.Core.MsoTriState.msoTrue, picX, picY, 250, 250);
                                if (picX == 885)
                                {
                                    picY += 70;
                                    picX = 125;
                                }
                                else
                                {
                                    picX += 190;
                                }
                            }
                        }
                    }
                }
            }
            else if (Aboxof == "40")
            {
                //載入嘜頭資料
                using (conn = new SqlConnection(myConnectionString))
                {
                    conn.Open();

                    selectCmd = "SELECT  [Client],(vchPrint + ' ' + ProductName + ' ' + Marking) PartDescription,isnull([CustomerProductName],''),isnull([CustomerProductNo],''),[vchBoxs],isnull(PalletNo,'') FROM [ShippingHead] where [ListDate]='" + ListDate_LB.SelectedItem + "' and [ProductName]='" + ProductName_CB.SelectedItem + "'and [vchBoxs]='" + WhereBox_LB.SelectedItem + "'";
                    cmd = new SqlCommand(selectCmd, conn);
                    using (reader = cmd.ExecuteReader())
                    {
                        while (reader.Read())
                        {
                            Client = reader.GetString(0).Trim();
                            //載入客戶產品名稱
                            oSheet.Cells[1, 8] = reader.GetString(2);

                            //載入客戶產品型號
                            oSheet.Cells[2, 8] = reader.GetString(3);

                            //載入一箱幾隻
                            oSheet.Cells[4, 8] = Getcount;

                            //載入箱號
                            oSheet.Cells[14, 2] = reader.GetString(4);

                            //20200410 加入PO
                            oSheet.Cells[5, 11] = CustomerPO_L.Text;

                            //載入客戶名稱
                            oSheet.Cells[3, 8] = reader.GetString(0);

                            //載入棧板號
                            oSheet.Cells[14, 10] = reader.GetString(5);

                            if (reader.GetString(0).Trim().CompareTo("Wicked Sportz") == 0)
                            {
                                //該客戶要其自己的logo
                                oSheet.Shapes.AddPicture(Application.StartupPath + @".\LOGO-ENAIRGY_Wicked Sportz.jpg", Microsoft.Office.Core.MsoTriState.msoFalse,
                                                Microsoft.Office.Core.MsoTriState.msoTrue, 18, 17, 212, 125);
                                // Application.StartupPath + @".\LOGO-ENAIRGY_Wicked Sportz.jpg"
                            }
                            if (reader.GetString(0).Trim().CompareTo("達成數位") == 0)
                            {
                                //該客戶要其自己的logo
                                oSheet.Shapes.AddPicture(Application.StartupPath + @".\LOGO_DCT.jpg", Microsoft.Office.Core.MsoTriState.msoFalse,
                                                Microsoft.Office.Core.MsoTriState.msoTrue, 18, 17, 212, 125);
                                // Application.StartupPath + @".\LOGO-ENAIRGY_Wicked Sportz.jpg"
                            }
                            else if (reader.GetString(reader.GetOrdinal("Client")).Trim().Contains("ADRENALICIA S.L."))
                            {
                                oSheet.Shapes.AddPicture(Application.StartupPath + @".\LOGO_RogerSports.png", Microsoft.Office.Core.MsoTriState.msoFalse,
                            Microsoft.Office.Core.MsoTriState.msoTrue, 2, 17, 212, 125);
                            }
                            else if (Client.ToUpper().StartsWith("EMB"))
                            {
                                oSheet.Shapes.AddPicture(Application.StartupPath + @".\LOGO_EMB.png", Microsoft.Office.Core.MsoTriState.msoFalse,
                                                    Microsoft.Office.Core.MsoTriState.msoTrue, 2, 17, 212, 125);
                            }
                            else if (Client.ToUpper().StartsWith("PAINTBALL SPORTS"))
                            {
                                oSheet.Shapes.AddPicture(Application.StartupPath + @".\LOGO_Paintball Sports GmbH.png", Microsoft.Office.Core.MsoTriState.msoFalse,
                                                    Microsoft.Office.Core.MsoTriState.msoTrue, 2, 17, 212, 125);
                            }
                        }
                    }
                }

                string FirstCNO = "";

                //載入嘜頭氣瓶序號位子
                using (conn = new SqlConnection(myConnectionString))
                {
                    conn.Open();

                    selectCmd = "SELECT  * FROM [ShippingBody] where [ListDate]='" + ListDate_LB.SelectedItem + "' and [ProductName]='" + ProductName_CB.SelectedItem + "'and [WhereBox]='" + WhereBox_LB.SelectedItem + "'";
                    cmd = new SqlCommand(selectCmd, conn);
                    using (reader = cmd.ExecuteReader())
                    {
                        while (reader.Read())
                        {
                            switch (reader.GetString(5))
                            {
                                case "1":
                                    oSheet.Cells[6, 1] = reader.GetString(3);
                                    FirstCNO = reader.GetString(3);
                                    break;

                                case "2":
                                    oSheet.Cells[6, 3] = reader.GetString(3);
                                    break;

                                case "3":
                                    oSheet.Cells[6, 5] = reader.GetString(3);
                                    break;

                                case "4":
                                    oSheet.Cells[6, 7] = reader.GetString(3);
                                    break;

                                case "5":
                                    oSheet.Cells[6, 9] = reader.GetString(3);
                                    break;

                                case "6":
                                    oSheet.Cells[7, 1] = reader.GetString(3);
                                    break;

                                case "7":
                                    oSheet.Cells[7, 3] = reader.GetString(3);
                                    break;

                                case "8":
                                    oSheet.Cells[7, 5] = reader.GetString(3);
                                    break;

                                case "9":
                                    oSheet.Cells[7, 7] = reader.GetString(3);
                                    break;

                                case "10":
                                    oSheet.Cells[7, 9] = reader.GetString(3);
                                    break;

                                case "11":
                                    oSheet.Cells[8, 1] = reader.GetString(3);
                                    break;

                                case "12":
                                    oSheet.Cells[8, 3] = reader.GetString(3);
                                    break;

                                case "13":
                                    oSheet.Cells[8, 5] = reader.GetString(3);
                                    break;

                                case "14":
                                    oSheet.Cells[8, 7] = reader.GetString(3);
                                    break;

                                case "15":
                                    oSheet.Cells[8, 9] = reader.GetString(3);
                                    break;

                                case "16":
                                    oSheet.Cells[9, 1] = reader.GetString(3);
                                    break;

                                case "17":
                                    oSheet.Cells[9, 3] = reader.GetString(3);
                                    break;

                                case "18":
                                    oSheet.Cells[9, 5] = reader.GetString(3);
                                    break;

                                case "19":
                                    oSheet.Cells[9, 7] = reader.GetString(3);
                                    break;

                                case "20":
                                    oSheet.Cells[9, 9] = reader.GetString(3);
                                    break;

                                case "21":
                                    oSheet.Cells[10, 1] = reader.GetString(3);
                                    break;

                                case "22":
                                    oSheet.Cells[10, 3] = reader.GetString(3);
                                    break;

                                case "23":
                                    oSheet.Cells[10, 5] = reader.GetString(3);
                                    break;

                                case "24":
                                    oSheet.Cells[10, 7] = reader.GetString(3);
                                    break;

                                case "25":
                                    oSheet.Cells[10, 9] = reader.GetString(3);
                                    break;

                                case "26":
                                    oSheet.Cells[11, 1] = reader.GetString(3);
                                    break;

                                case "27":
                                    oSheet.Cells[11, 3] = reader.GetString(3);
                                    break;

                                case "28":
                                    oSheet.Cells[11, 5] = reader.GetString(3);
                                    break;

                                case "29":
                                    oSheet.Cells[11, 7] = reader.GetString(3);
                                    break;

                                case "30":
                                    oSheet.Cells[11, 9] = reader.GetString(3);
                                    break;

                                case "31":
                                    oSheet.Cells[12, 1] = reader.GetString(3);
                                    break;

                                case "32":
                                    oSheet.Cells[12, 3] = reader.GetString(3);
                                    break;

                                case "33":
                                    oSheet.Cells[12, 5] = reader.GetString(3);
                                    break;

                                case "34":
                                    oSheet.Cells[12, 7] = reader.GetString(3);
                                    break;

                                case "35":
                                    oSheet.Cells[12, 9] = reader.GetString(3);
                                    break;

                                case "36":
                                    oSheet.Cells[13, 1] = reader.GetString(3);
                                    break;

                                case "37":
                                    oSheet.Cells[13, 3] = reader.GetString(3);
                                    break;

                                case "38":
                                    oSheet.Cells[13, 5] = reader.GetString(3);
                                    break;

                                case "39":
                                    oSheet.Cells[13, 7] = reader.GetString(3);
                                    break;

                                case "40":
                                    oSheet.Cells[13, 9] = reader.GetString(3);
                                    break;
                            }
                        }
                    }

                    if ((Client.Contains("Scientific Gas Australia Pty Ltd") || Client == "Airtanks Limited") && PackingMarks.Trim().CompareTo("SGA-GLADIATAIR") == 0)
                    {
                        string ProductNO = "";

                        //該客戶要其自己的logo  PartNo   Part Description
                        selectCmd = "SELECT  Product_NO FROM MSNBody,Manufacturing where [CylinderNo]='" + FirstCNO + "' and vchManufacturingNo=  Manufacturing_NO";
                        cmd = new SqlCommand(selectCmd, conn);
                        using (reader = cmd.ExecuteReader())
                        {
                            if (reader.Read())
                            {
                                ProductNO = reader.GetValue(0).ToString();
                            }
                        }

                        selectCmd = "SELECT  ProductCode, ProductDescription FROM CustomerPackingMark where ProductNo='" + ProductNO + "' and LogoType='" + (PackingMarks.Trim().Contains("-") == true ? PackingMarks.Trim().Split('-')[1].Trim().ToUpper() : PackingMarks.Trim()) + "'";
                        cmd = new SqlCommand(selectCmd, conn);
                        using (reader = cmd.ExecuteReader())
                        {
                            if (reader.Read())
                            {
                                //載入客戶產品名稱
                                oSheet.Cells[1, 8] = reader.GetString(1);

                                //載入客戶產品型號
                                oSheet.Cells[2, 8] = reader.GetString(0);
                            }
                        }

                        //LOGO
                        oSheet.Shapes.AddPicture(Application.StartupPath + @".\LOGO_SGA_GLADIATAIR.png", Microsoft.Office.Core.MsoTriState.msoFalse,
                                            Microsoft.Office.Core.MsoTriState.msoTrue, 15, 17, 212, 125);
                    }
                    else if (Client.ToUpper().StartsWith("EMB"))
                    {
                        oSheet.Shapes.AddPicture(Application.StartupPath + @".\LOGO_EMB.png", Microsoft.Office.Core.MsoTriState.msoFalse,
                                            Microsoft.Office.Core.MsoTriState.msoTrue, 2, 17, 212, 125);
                    }
                    else if ((Client.ToUpper().StartsWith("HATSAN") == true) && PackingMarks.Trim().CompareTo("HATSAN") == 0)
                    {
                        //LOGO
                        oSheet.Shapes.AddPicture(Application.StartupPath + @".\LOGO_HATSAN.png", Microsoft.Office.Core.MsoTriState.msoFalse,
                                            Microsoft.Office.Core.MsoTriState.msoTrue, 2, 17, 212, 125);
                    }
                    else if (Client.ToUpper().StartsWith("PAINTBALL SPORTS"))
                    {
                        oSheet.Shapes.AddPicture(Application.StartupPath + @".\LOGO_Paintball Sports GmbH.png", Microsoft.Office.Core.MsoTriState.msoFalse,
                                            Microsoft.Office.Core.MsoTriState.msoTrue, 2, 17, 212, 125);
                    }

                    //if (StorageStatus == "N")//20190212
                    {
                        //預設位子在X:680,Y:155
                        //預設QRCODE圖片大小250*250

                        //插入二維條碼
                        int picX = 680, picY = 180;
                        string picadd = @"C:\QRCode\";

                        selectCmd = "SELECT  * FROM [ShippingHead] where [ListDate]='" + ListDate_LB.SelectedItem + "' and [ProductName]='" + ProductName_CB.Text + "' and [vchBoxs]='" + WhereBox_LB.SelectedItem + "'";
                        cmd = new SqlCommand(selectCmd, conn);
                        using (reader = cmd.ExecuteReader())
                        {
                            while (reader.Read())
                            {
                                Excel.Worksheet xSheet = (Excel.Worksheet)oWB.Sheets[1];
                                oSheet.Shapes.AddPicture(picadd + (reader.GetString(0) + reader.GetString(1) + reader.GetString(3)) + ".png", Microsoft.Office.Core.MsoTriState.msoFalse,
                                Microsoft.Office.Core.MsoTriState.msoTrue, picX, picY, 250, 250);
                                if (picX == 885)
                                {
                                    picY += 70;
                                    picX = 125;
                                }
                                else
                                {
                                    picX += 190;
                                }
                            }
                        }
                    }
                }
            }
            else if (Aboxof == "15")
            {
                string HowMuch = "";
                int Cumulative = 0;
                int Total = 0;

                //載入嘜頭資料
                using (conn = new SqlConnection(myConnectionString))
                {
                    conn.Open();

                    selectCmd = "SELECT  * FROM [ShippingHead] where [ListDate]='" + ListDate_LB.SelectedItem + "' and [ProductName]='" + ProductName_CB.SelectedItem + "'";
                    cmd = new SqlCommand(selectCmd, conn);
                    using (reader = cmd.ExecuteReader())
                    {
                        while (reader.Read())
                        {
                            HowMuch = reader.GetString(4);
                            Cumulative++;
                        }
                    }

                    Total = Convert.ToInt32(HowMuch) * Cumulative;

                    //載入嘜頭資料
                    selectCmd = "SELECT  [Client],(vchPrint + ' ' + ProductName + ' ' + Marking) PartDescription,isnull([CustomerProductName],''),isnull([CustomerProductNo],''),[vchBoxs],isnull(PalletNo,'') FROM [ShippingHead] where [ListDate]='" + ListDate_LB.SelectedItem + "' and [ProductName]='" + ProductName_CB.SelectedItem + "'and [vchBoxs]='" + WhereBox_LB.SelectedItem + "'";
                    cmd = new SqlCommand(selectCmd, conn);
                    using (reader = cmd.ExecuteReader())
                    {
                        while (reader.Read())
                        {
                            Client = reader.GetString(0).Trim();
                            //載入客戶產品名稱
                            oSheet.Cells[1, 7] = reader.GetString(reader.GetOrdinal("PartDescription"));

                            //載入客戶產品型號
                            oSheet.Cells[2, 7] = reader.GetString(3);

                            //載入一箱幾隻
                            oSheet.Cells[4, 7] = Getcount;

                            //載入箱號
                            oSheet.Cells[9, 2] = reader.GetString(4);

                            //20200410 加入PO
                            oSheet.Cells[5, 11] = CustomerPO_L.Text;

                            //載入客戶名稱
                            oSheet.Cells[3, 7] = reader.GetString(0);

                            //棧板號
                            oSheet.Cells[9, 10] = reader.GetString(5);

                            if (reader.GetString(reader.GetOrdinal("Client")).Trim().CompareTo("Wicked Sportz") == 0)
                            {
                                //該客戶要其自己的logo
                                oSheet.Shapes.AddPicture(Application.StartupPath + @".\LOGO-ENAIRGY_Wicked Sportz.jpg", Microsoft.Office.Core.MsoTriState.msoFalse,
                                                Microsoft.Office.Core.MsoTriState.msoTrue, 3, 17, 212, 125);
                                // Application.StartupPath + @".\LOGO-ENAIRGY_Wicked Sportz.jpg"
                            }
                            if (reader.GetString(0).Trim().CompareTo("達成數位") == 0)
                            {
                                //該客戶要其自己的logo
                                oSheet.Shapes.AddPicture(Application.StartupPath + @".\LOGO_DCT.jpg", Microsoft.Office.Core.MsoTriState.msoFalse,
                                                Microsoft.Office.Core.MsoTriState.msoTrue, 3, 17, 212, 125);
                                // Application.StartupPath + @".\LOGO-ENAIRGY_Wicked Sportz.jpg"
                            }
                            else if (reader.GetString(reader.GetOrdinal("Client")).Trim().Contains("ADRENALICIA S.L."))
                            {
                                oSheet.Shapes.AddPicture(Application.StartupPath + @".\LOGO_RogerSports.png", Microsoft.Office.Core.MsoTriState.msoFalse,
                            Microsoft.Office.Core.MsoTriState.msoTrue, 2, 17, 212, 125);
                            }
                            else if (Client.ToUpper().StartsWith("EMB"))
                            {
                                oSheet.Shapes.AddPicture(Application.StartupPath + @".\LOGO_EMB.png", Microsoft.Office.Core.MsoTriState.msoFalse,
                                                    Microsoft.Office.Core.MsoTriState.msoTrue, 2, 17, 212, 125);
                            }
                            else if (Client.ToUpper().StartsWith("PAINTBALL SPORTS"))
                            {
                                oSheet.Shapes.AddPicture(Application.StartupPath + @".\LOGO_Paintball Sports GmbH.png", Microsoft.Office.Core.MsoTriState.msoFalse,
                                                    Microsoft.Office.Core.MsoTriState.msoTrue, 2, 17, 212, 125);
                            }
                        }
                    }
                }

                int serialnooneX = 7, serialnooneY = 209;
                string serialnooneadd = @"C:\SerialNoCode\";

                string FirstCNO = "";

                //載入嘜頭氣瓶序號位子
                using (conn = new SqlConnection(myConnectionString))
                {
                    conn.Open();

                    selectCmd = "SELECT  * FROM [ShippingBody] where [ListDate]='" + ListDate_LB.SelectedItem + "' and [ProductName]='" + ProductName_CB.SelectedItem + "'and [WhereBox]='" + WhereBox_LB.SelectedItem + "'";
                    cmd = new SqlCommand(selectCmd, conn);
                    using (reader = cmd.ExecuteReader())
                    {
                        while (reader.Read())
                        {
                            serialnooneX = 3; serialnooneY = 211;
                            switch (reader.GetString(5))
                            {
                                case "1":
                                    oSheet.Cells[6, 1] = reader.GetString(3);
                                    FirstCNO = reader.GetString(3);
                                    break;

                                case "2":
                                    oSheet.Cells[6, 3] = reader.GetString(3);
                                    break;

                                case "3":
                                    oSheet.Cells[6, 5] = reader.GetString(3);
                                    break;

                                case "4":
                                    oSheet.Cells[6, 7] = reader.GetString(3);
                                    break;

                                case "5":
                                    oSheet.Cells[6, 9] = reader.GetString(3);
                                    break;

                                case "6":
                                    oSheet.Cells[7, 1] = reader.GetString(3);
                                    break;

                                case "7":
                                    oSheet.Cells[7, 3] = reader.GetString(3);
                                    break;

                                case "8":
                                    oSheet.Cells[7, 5] = reader.GetString(3);
                                    break;

                                case "9":
                                    oSheet.Cells[7, 7] = reader.GetString(3);
                                    break;

                                case "10":
                                    oSheet.Cells[7, 9] = reader.GetString(3);
                                    break;

                                case "11":
                                    oSheet.Cells[8, 1] = reader.GetString(3);
                                    break;

                                case "12":
                                    oSheet.Cells[8, 3] = reader.GetString(3);
                                    break;

                                case "13":
                                    oSheet.Cells[8, 5] = reader.GetString(3);
                                    break;

                                case "14":
                                    oSheet.Cells[8, 7] = reader.GetString(3);
                                    break;

                                case "15":
                                    oSheet.Cells[8, 9] = reader.GetString(3);
                                    break;
                            }
                            serialnooneX = serialnooneX + ((Convert.ToInt32(reader.GetString(5)) + 4) % 5) * 145;
                            serialnooneY = serialnooneY + ((Convert.ToInt32(reader.GetString(5)) - 1) / 5) * 75;
                            oSheet.Shapes.AddPicture(serialnooneadd + reader.GetString(3) + ".png", Microsoft.Office.Core.MsoTriState.msoFalse,
                            Microsoft.Office.Core.MsoTriState.msoTrue, serialnooneX, serialnooneY, 44, 44);//, 130, 25);
                        }
                    }


                    if ((Client.Contains("Scientific Gas Australia Pty Ltd") || Client == "Airtanks Limited") && PackingMarks.Trim().CompareTo("SGA-GLADIATAIR") == 0)
                    {
                        string ProductNO = "";

                        //該客戶要其自己的logo  PartNo   Part Description
                        selectCmd = "SELECT  Product_NO FROM MSNBody,Manufacturing where [CylinderNo]='" + FirstCNO + "' and vchManufacturingNo=  Manufacturing_NO";
                        cmd = new SqlCommand(selectCmd, conn);
                        using (reader = cmd.ExecuteReader())
                        {
                            if (reader.Read())
                            {
                                ProductNO = reader.GetValue(0).ToString();
                            }
                        }

                        selectCmd = "SELECT  ProductCode, ProductDescription FROM CustomerPackingMark where ProductNo='" + ProductNO + "' and LogoType='" + (PackingMarks.Trim().Contains("-") == true ? PackingMarks.Trim().Split('-')[1].Trim().ToUpper() : PackingMarks.Trim()) + "'";
                        cmd = new SqlCommand(selectCmd, conn);
                        using (reader = cmd.ExecuteReader())
                        {
                            if (reader.Read())
                            {
                                //載入客戶產品名稱
                                oSheet.Cells[1, 7] = reader.GetString(1);

                                //載入客戶產品型號
                                oSheet.Cells[2, 7] = reader.GetString(0);
                            }
                        }

                        //LOGO
                        oSheet.Shapes.AddPicture(Application.StartupPath + @".\LOGO_SGA_GLADIATAIR.png", Microsoft.Office.Core.MsoTriState.msoFalse,
                                            Microsoft.Office.Core.MsoTriState.msoTrue, 3, 17, 212, 125);
                    }
                    else if ((Client.ToUpper().StartsWith("HATSAN") == true) && PackingMarks.Trim().CompareTo("HATSAN") == 0)
                    {
                        //LOGO
                        oSheet.Shapes.AddPicture(Application.StartupPath + @".\LOGO_HATSAN.png", Microsoft.Office.Core.MsoTriState.msoFalse,
                                            Microsoft.Office.Core.MsoTriState.msoTrue, 3, 17, 212, 125);
                    }
                    else if (Client.ToUpper().StartsWith("EMB"))
                    {
                        oSheet.Shapes.AddPicture(Application.StartupPath + @".\LOGO_EMB.png", Microsoft.Office.Core.MsoTriState.msoFalse,
                                            Microsoft.Office.Core.MsoTriState.msoTrue, 2, 17, 212, 125);
                    }
                    else if (Client.ToUpper().StartsWith("PAINTBALL SPORTS"))
                    {
                        oSheet.Shapes.AddPicture(Application.StartupPath + @".\LOGO_Paintball Sports GmbH.png", Microsoft.Office.Core.MsoTriState.msoFalse,
                                            Microsoft.Office.Core.MsoTriState.msoTrue, 2, 17, 212, 125);
                    }

                    //if (StorageStatus == "N")//20190212
                    {
                        //預設位子在X:680,Y:155
                        //預設QRCODE圖片大小250*250

                        //插入圖片
                        int picX = 732, picY = 187;
                        string picadd = @"C:\QRCode\";

                        selectCmd = "SELECT  * FROM [ShippingHead] where [ListDate]='" + ListDate_LB.SelectedItem + "' and [ProductName]='" + ProductName_CB.Text + "' and [vchBoxs]='" + WhereBox_LB.SelectedItem + "'";
                        cmd = new SqlCommand(selectCmd, conn);
                        using (reader = cmd.ExecuteReader())
                        {
                            while (reader.Read())
                            {
                                Excel.Worksheet xSheet = (Excel.Worksheet)oWB.Sheets[1];
                                oSheet.Shapes.AddPicture(picadd + (reader.GetString(0) + reader.GetString(1) + reader.GetString(3)) + ".png", Microsoft.Office.Core.MsoTriState.msoFalse,
                                Microsoft.Office.Core.MsoTriState.msoTrue, picX, picY, 250, 250);
                                if (picX == 885)
                                {
                                    picY += 70;
                                    picX = 125;
                                }
                                else
                                {
                                    picX += 190;
                                }
                            }
                        }
                    }
                }
            }
            else if (Aboxof == "12")
            {
                string HowMuch = "";
                int Cumulative = 0;
                int Total = 0;
                string DemandNo = string.Empty;

                //載入嘜頭資料
                using (conn = new SqlConnection(myConnectionString))
                {
                    conn.Open();

                    selectCmd = "SELECT  * FROM [ShippingHead] where [ListDate]='" + ListDate_LB.SelectedItem + "' and [ProductName]='" + ProductName_CB.SelectedItem + "'";
                    cmd = new SqlCommand(selectCmd, conn);
                    using (reader = cmd.ExecuteReader())
                    {
                        while (reader.Read())
                        {
                            HowMuch = reader.GetString(4);
                            DemandNo = reader.GetString(reader.GetOrdinal("DemandNo"));
                            Cumulative++;
                        }
                    }

                    Total = Convert.ToInt32(HowMuch) * Cumulative;

                    //載入嘜頭資料

                    selectCmd = "SELECT  [Client],(vchPrint + ' ' + ProductName + ' ' + Marking) PartDescription,isnull([CustomerProductName],''),isnull([CustomerProductNo],''),[vchBoxs],isnull(PalletNo,'') FROM [ShippingHead] where [ListDate]='" + ListDate_LB.SelectedItem + "' and [ProductName]='" + ProductName_CB.SelectedItem + "'and [vchBoxs]='" + WhereBox_LB.SelectedItem + "'";
                    cmd = new SqlCommand(selectCmd, conn);
                    using (reader = cmd.ExecuteReader())
                    {
                        while (reader.Read())
                        {
                            //1.20200821 AMS CC 試單 暫依此規則
                            if (DemandNo == "2201-20200820001")
                            {
                                oSheet.Cells[5, 1] = "Batch No./Serial No.";

                            }
                            Client = reader.GetString(0).Trim();
                            //載入客戶產品名稱
                            oSheet.Cells[1, 7] = reader.GetString(reader.GetOrdinal("PartDescription"));

                            //載入客戶產品型號
                            oSheet.Cells[2, 7] = reader.GetString(3);

                            //載入一箱幾隻
                            oSheet.Cells[4, 7] = Getcount;

                            //載入箱號
                            oSheet.Cells[12, 2] = reader.GetString(4);

                            //20200410 加入PO
                            oSheet.Cells[5, 9] = CustomerPO_L.Text;

                            //載入客戶名稱
                            oSheet.Cells[3, 7] = reader.GetString(0);

                            //載入箱號
                            oSheet.Cells[12, 8] = reader.GetString(5);

                            if (reader.GetString(0).Trim().CompareTo("Wicked Sportz") == 0)
                            {
                                //該客戶要其自己的logo
                                oSheet.Shapes.AddPicture(Application.StartupPath + @".\LOGO-ENAIRGY_Wicked Sportz.jpg", Microsoft.Office.Core.MsoTriState.msoFalse,
                                                Microsoft.Office.Core.MsoTriState.msoTrue, 12, 17, 212, 125);
                                // Application.StartupPath + @".\LOGO-ENAIRGY_Wicked Sportz.jpg"
                            }
                            if (reader.GetString(0).Trim().CompareTo("達成數位") == 0)
                            {
                                //該客戶要其自己的logo
                                oSheet.Shapes.AddPicture(Application.StartupPath + @".\LOGO_DCT.jpg", Microsoft.Office.Core.MsoTriState.msoFalse,
                                                Microsoft.Office.Core.MsoTriState.msoTrue, 12, 17, 212, 125);
                                // Application.StartupPath + @".\LOGO-ENAIRGY_Wicked Sportz.jpg"
                            }
                            else if (reader.GetString(reader.GetOrdinal("Client")).Trim().Contains("ADRENALICIA S.L."))
                            {
                                oSheet.Shapes.AddPicture(Application.StartupPath + @".\LOGO_RogerSports.png", Microsoft.Office.Core.MsoTriState.msoFalse,
                            Microsoft.Office.Core.MsoTriState.msoTrue, 2, 17, 212, 125);
                            }
                            else if (Client.ToUpper().StartsWith("EMB"))
                            {
                                oSheet.Shapes.AddPicture(Application.StartupPath + @".\LOGO_EMB.png", Microsoft.Office.Core.MsoTriState.msoFalse,
                                                    Microsoft.Office.Core.MsoTriState.msoTrue, 2, 17, 212, 125);
                            }
                            else if (Client.ToUpper().StartsWith("PAINTBALL SPORTS"))
                            {
                                oSheet.Shapes.AddPicture(Application.StartupPath + @".\LOGO_Paintball Sports GmbH.png", Microsoft.Office.Core.MsoTriState.msoFalse,
                                                    Microsoft.Office.Core.MsoTriState.msoTrue, 2, 17, 212, 125);
                            }
                        }
                    }
                }


                int serialnooneX = 10, serialnooneY = 212;
                string serialnooneadd = @"C:\SerialNoCode\";

                string FirstCNO = "";

                //載入嘜頭氣瓶序號位子
                using (conn = new SqlConnection(myConnectionString))
                {
                    conn.Open();

                    selectCmd = "SELECT  * FROM [ShippingBody] where [ListDate]='" + ListDate_LB.SelectedItem + "' and [ProductName]='" + ProductName_CB.SelectedItem + "'and [WhereBox]='" + WhereBox_LB.SelectedItem + "'";
                    cmd = new SqlCommand(selectCmd, conn);
                    using (reader = cmd.ExecuteReader())
                    {
                        while (reader.Read())
                        {
                            serialnooneX = 3; serialnooneY = 212;
                            switch (reader.GetString(5))
                            {
                                case "1":
                                    //1.20200821 AMS CC 試單 暫依此規則
                                    if (DemandNo == "2201-20200820001")
                                    {
                                        oSheet.Cells[6, 1] = reader.GetString(reader.GetOrdinal("LotNumber")) + "\n" + reader.GetString(3);
                                    }
                                    else
                                    {
                                        oSheet.Cells[6, 1] = reader.GetString(3);
                                    }
                                    FirstCNO = reader.GetString(3);
                                    break;

                                case "2":
                                    //1.20200821 AMS CC 試單 暫依此規則
                                    if (DemandNo == "2201-20200820001")
                                    {
                                        oSheet.Cells[6, 3] = reader.GetString(reader.GetOrdinal("LotNumber")) + "\n" + reader.GetString(3);
                                    }
                                    else
                                    {
                                        oSheet.Cells[6, 3] = reader.GetString(3);
                                    }
                                    break;

                                case "3":
                                    //1.20200821 AMS CC 試單 暫依此規則
                                    if (DemandNo == "2201-20200820001")
                                    {
                                        oSheet.Cells[6, 5] = reader.GetString(reader.GetOrdinal("LotNumber")) + "\n" + reader.GetString(3);
                                    }
                                    else
                                    {
                                        oSheet.Cells[6, 5] = reader.GetString(3);
                                    }
                                    break;

                                case "4":
                                    //1.20200821 AMS CC 試單 暫依此規則
                                    if (DemandNo == "2201-20200820001")
                                    {
                                        oSheet.Cells[6, 7] = reader.GetString(reader.GetOrdinal("LotNumber")) + "\n" + reader.GetString(3);
                                    }
                                    else
                                    {
                                        oSheet.Cells[6, 7] = reader.GetString(3);
                                    }
                                    break;

                                case "5":
                                    //1.20200821 AMS CC 試單 暫依此規則
                                    if (DemandNo == "2201-20200820001")
                                    {
                                        oSheet.Cells[8, 1] = reader.GetString(reader.GetOrdinal("LotNumber")) + "\n" + reader.GetString(3);
                                    }
                                    else
                                    {
                                        oSheet.Cells[8, 1] = reader.GetString(3);
                                    }
                                    break;

                                case "6":
                                    //1.20200821 AMS CC 試單 暫依此規則
                                    if (DemandNo == "2201-20200820001")
                                    {
                                        oSheet.Cells[8, 3] = reader.GetString(reader.GetOrdinal("LotNumber")) + "\n" + reader.GetString(3);
                                    }
                                    else
                                    {
                                        oSheet.Cells[8, 3] = reader.GetString(3);
                                    }
                                    break;

                                case "7":
                                    //1.20200821 AMS CC 試單 暫依此規則
                                    if (DemandNo == "2201-20200820001")
                                    {
                                        oSheet.Cells[8, 5] = reader.GetString(reader.GetOrdinal("LotNumber")) + "\n" + reader.GetString(3);
                                    }
                                    else
                                    {
                                        oSheet.Cells[8, 5] = reader.GetString(3);
                                    }
                                    break;

                                case "8":
                                    //1.20200821 AMS CC 試單 暫依此規則
                                    if (DemandNo == "2201-20200820001")
                                    {
                                        oSheet.Cells[8, 7] = reader.GetString(reader.GetOrdinal("LotNumber")) + "\n" + reader.GetString(3);
                                    }
                                    else
                                    {
                                        oSheet.Cells[8, 7] = reader.GetString(3);
                                    }
                                    break;

                                case "9":
                                    //1.20200821 AMS CC 試單 暫依此規則
                                    if (DemandNo == "2201-20200820001")
                                    {
                                        oSheet.Cells[10, 1] = reader.GetString(reader.GetOrdinal("LotNumber")) + "\n" + reader.GetString(3);
                                    }
                                    else
                                    {
                                        oSheet.Cells[10, 1] = reader.GetString(3);
                                    }
                                    break;

                                case "10":
                                    //1.20200821 AMS CC 試單 暫依此規則
                                    if (DemandNo == "2201-20200820001")
                                    {
                                        oSheet.Cells[10, 3] = reader.GetString(reader.GetOrdinal("LotNumber")) + "\n" + reader.GetString(3);
                                    }
                                    else
                                    {
                                        oSheet.Cells[10, 3] = reader.GetString(3);
                                    }
                                    break;

                                case "11":
                                    //1.20200821 AMS CC 試單 暫依此規則
                                    if (DemandNo == "2201-20200820001")
                                    {
                                        oSheet.Cells[10, 5] = reader.GetString(reader.GetOrdinal("LotNumber")) + "\n" + reader.GetString(3);
                                    }
                                    else
                                    {
                                        oSheet.Cells[10, 5] = reader.GetString(3);
                                    }
                                    break;

                                case "12":
                                    //1.20200821 AMS CC 試單 暫依此規則
                                    if (DemandNo == "2201-20200820001")
                                    {
                                        oSheet.Cells[10, 7] = reader.GetString(reader.GetOrdinal("LotNumber")) + "\n" + reader.GetString(3);
                                    }
                                    else
                                    {
                                        oSheet.Cells[10, 7] = reader.GetString(3);
                                    }
                                    break;
                            }
                            serialnooneX = serialnooneX + ((Convert.ToInt32(reader.GetString(5)) + 3) % 4) * 157;
                            serialnooneY = serialnooneY + ((Convert.ToInt32(reader.GetString(5)) - 1) / 4) * 75;
                            oSheet.Shapes.AddPicture(serialnooneadd + reader.GetString(3) + ".png", Microsoft.Office.Core.MsoTriState.msoFalse,
                            Microsoft.Office.Core.MsoTriState.msoTrue, serialnooneX, serialnooneY, 45, 45);//, 130, 25);
                        }
                    }

                    if ((Client.Contains("Scientific Gas Australia Pty Ltd") || Client == "Airtanks Limited") && PackingMarks.Trim().CompareTo("SGA-GLADIATAIR") == 0)
                    {
                        string ProductNO = "";

                        //該客戶要其自己的logo  PartNo   Part Description
                        selectCmd = "SELECT  Product_NO FROM MSNBody,Manufacturing where [CylinderNo]='" + FirstCNO + "' and vchManufacturingNo=  Manufacturing_NO";
                        cmd = new SqlCommand(selectCmd, conn);
                        using (reader = cmd.ExecuteReader())
                        {
                            if (reader.Read())
                            {
                                ProductNO = reader.GetValue(0).ToString();
                            }
                        }

                        selectCmd = "SELECT  ProductCode, ProductDescription FROM CustomerPackingMark where ProductNo='" + ProductNO + "' and LogoType='" + (PackingMarks.Trim().Contains("-") == true ? PackingMarks.Trim().Split('-')[1].Trim().ToUpper() : PackingMarks.Trim()) + "'";
                        cmd = new SqlCommand(selectCmd, conn);
                        using (reader = cmd.ExecuteReader())
                        {
                            if (reader.Read())
                            {
                                //載入客戶產品名稱
                                oSheet.Cells[1, 7] = reader.GetString(1);

                                //載入客戶產品型號
                                oSheet.Cells[2, 7] = reader.GetString(0);
                            }
                        }

                        //LOGO
                        oSheet.Shapes.AddPicture(Application.StartupPath + @".\LOGO_SGA_GLADIATAIR.png", Microsoft.Office.Core.MsoTriState.msoFalse,
                                            Microsoft.Office.Core.MsoTriState.msoTrue, 3, 17, 212, 125);
                    }
                    else if ((Client.ToUpper().StartsWith("HATSAN") == true) && PackingMarks.Trim().CompareTo("HATSAN") == 0)
                    {
                        //LOGO
                        oSheet.Shapes.AddPicture(Application.StartupPath + @".\LOGO_HATSAN.png", Microsoft.Office.Core.MsoTriState.msoFalse,
                                            Microsoft.Office.Core.MsoTriState.msoTrue, 3, 17, 212, 125);
                    }
                    else if (Client.ToUpper().StartsWith("EMB"))
                    {
                        oSheet.Shapes.AddPicture(Application.StartupPath + @".\LOGO_EMB.png", Microsoft.Office.Core.MsoTriState.msoFalse,
                                            Microsoft.Office.Core.MsoTriState.msoTrue, 2, 17, 212, 125);
                    }
                    else if (Client.ToUpper().StartsWith("PAINTBALL SPORTS"))
                    {
                        oSheet.Shapes.AddPicture(Application.StartupPath + @".\LOGO_Paintball Sports GmbH.png", Microsoft.Office.Core.MsoTriState.msoFalse,
                                            Microsoft.Office.Core.MsoTriState.msoTrue, 2, 17, 212, 125);
                    }

                    //if (StorageStatus == "N")//20190212
                    {
                        //預設位子在X:680,Y:155
                        //預設QRCODE圖片大小250*250

                        //插入圖片
                        int picX = 680, picY = 185;
                        string picadd = @"C:\QRCode\";

                        selectCmd = "SELECT  * FROM [ShippingHead] where [ListDate]='" + ListDate_LB.SelectedItem + "' and [ProductName]='" + ProductName_CB.Text + "' and [vchBoxs]='" + WhereBox_LB.SelectedItem + "'";
                        cmd = new SqlCommand(selectCmd, conn);
                        using (reader = cmd.ExecuteReader())
                        {
                            while (reader.Read())
                            {
                                Excel.Worksheet xSheet = (Excel.Worksheet)oWB.Sheets[1];
                                oSheet.Shapes.AddPicture(picadd + (reader.GetString(0) + reader.GetString(1) + reader.GetString(3)) + ".png", Microsoft.Office.Core.MsoTriState.msoFalse,
                                Microsoft.Office.Core.MsoTriState.msoTrue, picX, picY, 250, 250);
                                if (picX == 885)
                                {
                                    picY += 70;
                                    picX = 125;
                                }
                                else
                                {
                                    picX += 190;
                                }
                            }
                        }
                    }
                }
            }
            else if (Aboxof == "6")
            {
                string HowMuch = "";
                int Cumulative = 0;
                int Total = 0;

                //載入嘜頭資料
                using (conn = new SqlConnection(myConnectionString))
                {
                    conn.Open();

                    selectCmd = "SELECT  * FROM [ShippingHead] where [ListDate]='" + ListDate_LB.SelectedItem + "' and [ProductName]='" + ProductName_CB.SelectedItem + "'";
                    cmd = new SqlCommand(selectCmd, conn);
                    using (reader = cmd.ExecuteReader())
                    {
                        while (reader.Read())
                        {
                            HowMuch = reader.GetString(4);
                            Cumulative++;
                        }
                    }

                    Total = Convert.ToInt32(HowMuch) * Cumulative;

                    //載入嘜頭資料
                    selectCmd = "SELECT  [Client],(vchPrint + ' ' + ProductName + ' ' + Marking) PartDescription,isnull([CustomerProductName],''),isnull([CustomerProductNo],''),[vchBoxs],isnull(PalletNo,'') FROM [ShippingHead] where [ListDate]='" + ListDate_LB.SelectedItem + "' and [ProductName]='" + ProductName_CB.SelectedItem + "'and [vchBoxs]='" + WhereBox_LB.SelectedItem + "'";
                    cmd = new SqlCommand(selectCmd, conn);
                    using (reader = cmd.ExecuteReader())
                    {
                        while (reader.Read())
                        {
                            Client = reader.GetString(0).Trim();
                            //載入客戶產品名稱
                            oSheet.Cells[1, 7] = reader.GetString(reader.GetOrdinal("PartDescription"));

                            //載入客戶產品型號
                            oSheet.Cells[2, 7] = reader.GetString(3);

                            //載入一箱幾隻
                            oSheet.Cells[4, 7] = Getcount;

                            //載入箱號
                            oSheet.Cells[10, 2] = reader.GetString(4);

                            //20200410 加入PO
                            oSheet.Cells[5, 9] = CustomerPO_L.Text;

                            //載入客戶名稱
                            oSheet.Cells[3, 7] = reader.GetString(0);

                            //載入箱號
                            oSheet.Cells[10, 8] = reader.GetString(5);

                            if (reader.GetString(0).Trim().CompareTo("Wicked Sportz") == 0)
                            {
                                //該客戶要其自己的logo  //Wicked Sportz
                                oSheet.Shapes.AddPicture(Application.StartupPath + @".\LOGO-ENAIRGY_Wicked Sportz.jpg", Microsoft.Office.Core.MsoTriState.msoFalse,
                                                Microsoft.Office.Core.MsoTriState.msoTrue, 12, 17, 212, 125);
                                // Application.StartupPath + @".\LOGO-ENAIRGY_Wicked Sportz.jpg"
                            }
                            if (reader.GetString(0).Trim().CompareTo("達成數位") == 0)
                            {
                                //該客戶要其自己的logo  //Wicked Sportz
                                oSheet.Shapes.AddPicture(Application.StartupPath + @".\LOGO_DCT.jpg", Microsoft.Office.Core.MsoTriState.msoFalse,
                                                Microsoft.Office.Core.MsoTriState.msoTrue, 12, 17, 212, 125);
                                // Application.StartupPath + @".\LOGO-ENAIRGY_Wicked Sportz.jpg"
                            }
                            else if (reader.GetString(reader.GetOrdinal("Client")).Trim().Contains("ADRENALICIA S.L."))
                            {
                                oSheet.Shapes.AddPicture(Application.StartupPath + @".\LOGO_RogerSports.png", Microsoft.Office.Core.MsoTriState.msoFalse,
                            Microsoft.Office.Core.MsoTriState.msoTrue, 2, 17, 212, 125);
                            }
                            else if (Client.ToUpper().StartsWith("EMB"))
                            {
                                oSheet.Shapes.AddPicture(Application.StartupPath + @".\LOGO_EMB.png", Microsoft.Office.Core.MsoTriState.msoFalse,
                                                    Microsoft.Office.Core.MsoTriState.msoTrue, 2, 17, 212, 125);
                            }
                            else if (Client.ToUpper().StartsWith("PAINTBALL SPORTS"))
                            {
                                oSheet.Shapes.AddPicture(Application.StartupPath + @".\LOGO_Paintball Sports GmbH.png", Microsoft.Office.Core.MsoTriState.msoFalse,
                                                    Microsoft.Office.Core.MsoTriState.msoTrue, 2, 17, 212, 125);
                            }
                        }
                    }
                }

                int serialnooneX = 10, serialnooneY = 309;
                string serialnooneadd = @"C:\SerialNoCode\";
                string FirstCNO = "";

                //載入嘜頭氣瓶序號位子
                using (conn = new SqlConnection(myConnectionString))
                {
                    conn.Open();

                    selectCmd = "SELECT  * FROM [ShippingBody] where [ListDate]='" + ListDate_LB.SelectedItem + "' and [ProductName]='" + ProductName_CB.SelectedItem + "'and [WhereBox]='" + WhereBox_LB.SelectedItem + "'";
                    cmd = new SqlCommand(selectCmd, conn);
                    using (reader = cmd.ExecuteReader())
                    {
                        while (reader.Read())
                        {
                            serialnooneX = 49; serialnooneY = 215;
                            switch (reader.GetString(5))
                            {
                                case "1":
                                    oSheet.Cells[6, 1] = reader.GetString(3);
                                    FirstCNO = reader.GetString(3);
                                    break;

                                case "2":
                                    oSheet.Cells[6, 3] = reader.GetString(3);
                                    break;

                                case "3":
                                    oSheet.Cells[6, 5] = reader.GetString(3);
                                    break;

                                case "4":
                                    oSheet.Cells[8, 1] = reader.GetString(3);
                                    break;

                                case "5":
                                    oSheet.Cells[8, 3] = reader.GetString(3);
                                    break;

                                case "6":
                                    oSheet.Cells[8, 5] = reader.GetString(3);
                                    break;
                            }
                            serialnooneX = serialnooneX + ((Convert.ToInt32(reader.GetString(5)) + 3) % 3) * 215;
                            serialnooneY = serialnooneY + ((Convert.ToInt32(reader.GetString(5)) - 1) / 3) * 111;
                            oSheet.Shapes.AddPicture(serialnooneadd + reader.GetString(3) + ".png", Microsoft.Office.Core.MsoTriState.msoFalse,
                            Microsoft.Office.Core.MsoTriState.msoTrue, serialnooneX, serialnooneY, 60, 60);//, 130, 25);
                        }
                    }

                    if ((Client.Contains("Scientific Gas Australia Pty Ltd") || Client == "Airtanks Limited") && PackingMarks.Trim().CompareTo("SGA-GLADIATAIR") == 0)
                    {
                        string ProductNO = "";

                        //該客戶要其自己的logo  PartNo   Part Description
                        selectCmd = "SELECT  Product_NO FROM MSNBody,Manufacturing where [CylinderNo]='" + FirstCNO + "' and vchManufacturingNo=  Manufacturing_NO";
                        cmd = new SqlCommand(selectCmd, conn);
                        using (reader = cmd.ExecuteReader())
                        {
                            if (reader.Read())
                            {
                                ProductNO = reader.GetValue(0).ToString();
                            }
                        }

                        selectCmd = "SELECT  ProductCode, ProductDescription FROM CustomerPackingMark where ProductNo='" + ProductNO + "' and LogoType='" + (PackingMarks.Trim().Contains("-") == true ? PackingMarks.Trim().Split('-')[1].Trim().ToUpper() : PackingMarks.Trim()) + "'";
                        cmd = new SqlCommand(selectCmd, conn);
                        using (reader = cmd.ExecuteReader())
                        {
                            if (reader.Read())
                            {
                                //載入客戶產品名稱
                                oSheet.Cells[1, 7] = reader.GetString(1);

                                //載入客戶產品型號
                                oSheet.Cells[2, 7] = reader.GetString(0);
                            }
                        }

                        //LOGO
                        oSheet.Shapes.AddPicture(Application.StartupPath + @".\LOGO_SGA_GLADIATAIR.png", Microsoft.Office.Core.MsoTriState.msoFalse,
                                            Microsoft.Office.Core.MsoTriState.msoTrue, 12, 17, 212, 125);
                    }
                    else if ((Client.ToUpper().StartsWith("HATSAN") == true) && PackingMarks.Trim().CompareTo("HATSAN") == 0)
                    {
                        //LOGO
                        oSheet.Shapes.AddPicture(Application.StartupPath + @".\LOGO_HATSAN.png", Microsoft.Office.Core.MsoTriState.msoFalse,
                                            Microsoft.Office.Core.MsoTriState.msoTrue, 12, 17, 212, 125);
                    }
                    else if (Client.ToUpper().StartsWith("EMB"))
                    {
                        oSheet.Shapes.AddPicture(Application.StartupPath + @".\LOGO_EMB.png", Microsoft.Office.Core.MsoTriState.msoFalse,
                                            Microsoft.Office.Core.MsoTriState.msoTrue, 2, 17, 212, 125);
                    }

                    else if (Client.ToUpper().StartsWith("PAINTBALL SPORTS"))
                    {
                        oSheet.Shapes.AddPicture(Application.StartupPath + @".\LOGO_Paintball Sports GmbH.png", Microsoft.Office.Core.MsoTriState.msoFalse,
                                            Microsoft.Office.Core.MsoTriState.msoTrue, 2, 17, 212, 125);
                    }

                    //if (StorageStatus == "N")//20190212
                    {

                        //預設位子在X:680,Y:155
                        //預設QRCODE圖片大小250*250

                        //插入圖片
                        int picX = 680, picY = 182;
                        string picadd = @"C:\QRCode\";

                        selectCmd = "SELECT  * FROM [ShippingHead] where [ListDate]='" + ListDate_LB.SelectedItem + "' and [ProductName]='" + ProductName_CB.Text + "' and [vchBoxs]='" + WhereBox_LB.SelectedItem + "'";
                        cmd = new SqlCommand(selectCmd, conn);
                        using (reader = cmd.ExecuteReader())
                        {
                            while (reader.Read())
                            {
                                Excel.Worksheet xSheet = (Excel.Worksheet)oWB.Sheets[1];
                                oSheet.Shapes.AddPicture(picadd + (reader.GetString(0) + reader.GetString(1) + reader.GetString(3)) + ".png", Microsoft.Office.Core.MsoTriState.msoFalse,
                                Microsoft.Office.Core.MsoTriState.msoTrue, picX, picY, 250, 250);
                                if (picX == 885)
                                {
                                    picY += 70;
                                    picX = 125;
                                }
                                else
                                {
                                    picX += 190;
                                }
                            }
                        }
                    }
                }
            }
            else if (Aboxof == "8")
            {
                string HowMuch = "";
                int Cumulative = 0;
                int Total = 0;

                //載入嘜頭資料
                using (conn = new SqlConnection(myConnectionString))
                {
                    conn.Open();

                    selectCmd = "SELECT  * FROM [ShippingHead] where [ListDate]='" + ListDate_LB.SelectedItem + "' and [ProductName]='" + ProductName_CB.SelectedItem + "'";
                    cmd = new SqlCommand(selectCmd, conn);
                    using (reader = cmd.ExecuteReader())
                    {
                        while (reader.Read())
                        {
                            HowMuch = reader.GetString(4);
                            Cumulative++;
                        }
                    }

                    Total = Convert.ToInt32(HowMuch) * Cumulative;

                    //載入嘜頭資料
                    selectCmd = "SELECT  [Client],(vchPrint + ' ' + ProductName + ' ' + Marking) PartDescription,isnull([CustomerProductName],''),isnull([CustomerProductNo],''),[vchBoxs],isnull(PalletNo,'') FROM [ShippingHead] where [ListDate]='" + ListDate_LB.SelectedItem + "' and [ProductName]='" + ProductName_CB.SelectedItem + "'and [vchBoxs]='" + WhereBox_LB.SelectedItem + "'";
                    cmd = new SqlCommand(selectCmd, conn);
                    using (reader = cmd.ExecuteReader())
                    {
                        while (reader.Read())
                        {
                            Client = reader.GetString(0).Trim();
                            //載入客戶產品名稱
                            oSheet.Cells[1, 7] = reader.GetString(reader.GetOrdinal("PartDescription"));

                            //載入客戶產品型號
                            oSheet.Cells[2, 7] = reader.GetString(3);

                            //載入一箱幾隻
                            oSheet.Cells[4, 7] = Getcount;

                            //載入箱號
                            oSheet.Cells[10, 2] = reader.GetString(4);

                            //20200410 加入PO
                            oSheet.Cells[5, 9] = CustomerPO_L.Text;

                            //載入客戶名稱
                            oSheet.Cells[3, 7] = reader.GetString(0);

                            //載入箱號
                            oSheet.Cells[10, 8] = reader.GetString(5);

                            if (reader.GetString(0).Trim().CompareTo("Wicked Sportz") == 0)
                            {
                                //該客戶要其自己的logo  //Wicked Sportz
                                oSheet.Shapes.AddPicture(Application.StartupPath + @".\LOGO-ENAIRGY_Wicked Sportz.jpg", Microsoft.Office.Core.MsoTriState.msoFalse,
                                                Microsoft.Office.Core.MsoTriState.msoTrue, 12, 17, 212, 125);
                                // Application.StartupPath + @".\LOGO-ENAIRGY_Wicked Sportz.jpg"
                            }
                            if (reader.GetString(0).Trim().CompareTo("達成數位") == 0)
                            {
                                //該客戶要其自己的logo  //Wicked Sportz
                                oSheet.Shapes.AddPicture(Application.StartupPath + @".\LOGO_DCT.jpg", Microsoft.Office.Core.MsoTriState.msoFalse,
                                                Microsoft.Office.Core.MsoTriState.msoTrue, 12, 17, 212, 125);
                                // Application.StartupPath + @".\LOGO-ENAIRGY_Wicked Sportz.jpg"
                            }
                            else if (reader.GetString(reader.GetOrdinal("Client")).Trim().Contains("ADRENALICIA S.L."))
                            {
                                oSheet.Shapes.AddPicture(Application.StartupPath + @".\LOGO_RogerSports.png", Microsoft.Office.Core.MsoTriState.msoFalse,
                            Microsoft.Office.Core.MsoTriState.msoTrue, 2, 17, 212, 125);
                            }
                            else if (Client.ToUpper().StartsWith("EMB"))
                            {
                                oSheet.Shapes.AddPicture(Application.StartupPath + @".\LOGO_EMB.png", Microsoft.Office.Core.MsoTriState.msoFalse,
                                                    Microsoft.Office.Core.MsoTriState.msoTrue, 2, 17, 212, 125);
                            }
                            else if (Client.ToUpper().StartsWith("PAINTBALL SPORTS"))
                            {
                                oSheet.Shapes.AddPicture(Application.StartupPath + @".\LOGO_Paintball Sports GmbH.png", Microsoft.Office.Core.MsoTriState.msoFalse,
                                                    Microsoft.Office.Core.MsoTriState.msoTrue, 2, 17, 212, 125);
                            }
                        }
                    }
                }

                int serialnooneX = 10, serialnooneY = 239;
                string serialnooneadd = @"C:\SerialNoCode\";
                string FirstCNO = "";

                //載入嘜頭氣瓶序號位子
                using (conn = new SqlConnection(myConnectionString))
                {
                    conn.Open();

                    selectCmd = "SELECT  * FROM [ShippingBody] where [ListDate]='" + ListDate_LB.SelectedItem + "' and [ProductName]='" + ProductName_CB.SelectedItem + "'and [WhereBox]='" + WhereBox_LB.SelectedItem + "'";
                    cmd = new SqlCommand(selectCmd, conn);
                    using (reader = cmd.ExecuteReader())
                    {
                        while (reader.Read())
                        {
                            serialnooneX = 49; serialnooneY = 215;
                            switch (reader.GetString(5))
                            {
                                case "1":
                                    oSheet.Cells[6, 1] = reader.GetString(3);
                                    FirstCNO = reader.GetString(3);
                                    break;

                                case "2":
                                    oSheet.Cells[6, 3] = reader.GetString(3);
                                    break;

                                case "3":
                                    oSheet.Cells[6, 5] = reader.GetString(3);
                                    break;

                                case "4":
                                    oSheet.Cells[6, 7] = reader.GetString(3);
                                    break;

                                case "5":
                                    oSheet.Cells[8, 1] = reader.GetString(3);
                                    break;

                                case "6":
                                    oSheet.Cells[8, 3] = reader.GetString(3);
                                    break;

                                case "7":
                                    oSheet.Cells[8, 5] = reader.GetString(3);
                                    break;

                                case "8":
                                    oSheet.Cells[8, 7] = reader.GetString(3);
                                    break;
                            }
                            serialnooneX = serialnooneX + ((Convert.ToInt32(reader.GetString(5)) + 3) % 4) * 159;
                            serialnooneY = serialnooneY + ((Convert.ToInt32(reader.GetString(5)) - 1) / 4) * 111;
                            oSheet.Shapes.AddPicture(serialnooneadd + reader.GetString(3) + ".png", Microsoft.Office.Core.MsoTriState.msoFalse,
                            Microsoft.Office.Core.MsoTriState.msoTrue, serialnooneX, serialnooneY, 60, 60);//, 130, 25);
                        }
                    }

                    if ((Client.Contains("Scientific Gas Australia Pty Ltd") || Client == "Airtanks Limited") && PackingMarks.Trim().CompareTo("SGA-GLADIATAIR") == 0)
                    {
                        string ProductNO = "";

                        //該客戶要其自己的logo  PartNo   Part Description
                        selectCmd = "SELECT  Product_NO FROM MSNBody,Manufacturing where [CylinderNo]='" + FirstCNO + "' and vchManufacturingNo=  Manufacturing_NO";
                        cmd = new SqlCommand(selectCmd, conn);
                        using (reader = cmd.ExecuteReader())
                        {
                            if (reader.Read())
                            {
                                ProductNO = reader.GetValue(0).ToString();
                            }
                        }

                        selectCmd = "SELECT  ProductCode, ProductDescription FROM CustomerPackingMark where ProductNo='" + ProductNO + "' and LogoType='" + (PackingMarks.Trim().Contains("-") == true ? PackingMarks.Trim().Split('-')[1].Trim().ToUpper() : PackingMarks.Trim()) + "'";
                        cmd = new SqlCommand(selectCmd, conn);
                        using (reader = cmd.ExecuteReader())
                        {
                            if (reader.Read())
                            {
                                //載入客戶產品名稱
                                oSheet.Cells[1, 7] = reader.GetString(1);

                                //載入客戶產品型號
                                oSheet.Cells[2, 7] = reader.GetString(0);
                            }
                        }

                        //LOGO
                        oSheet.Shapes.AddPicture(Application.StartupPath + @".\LOGO_SGA_GLADIATAIR.png", Microsoft.Office.Core.MsoTriState.msoFalse,
                                            Microsoft.Office.Core.MsoTriState.msoTrue, 12, 17, 212, 125);
                    }
                    else if ((Client.ToUpper().StartsWith("HATSAN") == true) && PackingMarks.Trim().CompareTo("HATSAN") == 0)
                    {
                        //LOGO
                        oSheet.Shapes.AddPicture(Application.StartupPath + @".\LOGO_HATSAN.png", Microsoft.Office.Core.MsoTriState.msoFalse,
                                            Microsoft.Office.Core.MsoTriState.msoTrue, 12, 17, 212, 125);
                    }
                    else if (Client.ToUpper().StartsWith("EMB"))
                    {
                        oSheet.Shapes.AddPicture(Application.StartupPath + @".\LOGO_EMB.png", Microsoft.Office.Core.MsoTriState.msoFalse,
                                            Microsoft.Office.Core.MsoTriState.msoTrue, 2, 17, 212, 125);
                    }

                    else if (Client.ToUpper().StartsWith("PAINTBALL SPORTS"))
                    {
                        oSheet.Shapes.AddPicture(Application.StartupPath + @".\LOGO_Paintball Sports GmbH.png", Microsoft.Office.Core.MsoTriState.msoFalse,
                                            Microsoft.Office.Core.MsoTriState.msoTrue, 2, 17, 212, 125);
                    }

                    //if (StorageStatus == "N")//20190212
                    {

                        //預設位子在X:680,Y:155
                        //預設QRCODE圖片大小250*250

                        //插入圖片
                        int picX = 680, picY = 182;
                        string picadd = @"C:\QRCode\";

                        selectCmd = "SELECT  * FROM [ShippingHead] where [ListDate]='" + ListDate_LB.SelectedItem + "' and [ProductName]='" + ProductName_CB.Text + "' and [vchBoxs]='" + WhereBox_LB.SelectedItem + "'";
                        cmd = new SqlCommand(selectCmd, conn);
                        using (reader = cmd.ExecuteReader())
                        {
                            while (reader.Read())
                            {
                                Excel.Worksheet xSheet = (Excel.Worksheet)oWB.Sheets[1];
                                oSheet.Shapes.AddPicture(picadd + (reader.GetString(0) + reader.GetString(1) + reader.GetString(3)) + ".png", Microsoft.Office.Core.MsoTriState.msoFalse,
                                Microsoft.Office.Core.MsoTriState.msoTrue, picX, picY, 250, 250);
                                if (picX == 885)
                                {
                                    picY += 70;
                                    picX = 125;
                                }
                                else
                                {
                                    picX += 190;
                                }
                            }
                        }
                    }
                }
            }
            else if (Aboxof == "16")
            {
                string HowMuch = "";
                int Cumulative = 0;
                int Total = 0;

                //載入嘜頭資料
                using (conn = new SqlConnection(myConnectionString))
                {
                    conn.Open();

                    selectCmd = "SELECT  * FROM [ShippingHead] where [ListDate]='" + ListDate_LB.SelectedItem + "' and [ProductName]='" + ProductName_CB.SelectedItem + "'";
                    cmd = new SqlCommand(selectCmd, conn);
                    using (reader = cmd.ExecuteReader())
                    {
                        while (reader.Read())
                        {
                            HowMuch = reader.GetString(4);
                            Cumulative++;
                        }
                    }

                    Total = Convert.ToInt32(HowMuch) * Cumulative;

                    //載入嘜頭資料
                    selectCmd = "SELECT  [Client],(vchPrint + ' ' + ProductName + ' ' + Marking) PartDescription,isnull([CustomerProductName],''),isnull([CustomerProductNo],''),[vchBoxs] ,isnull(PalletNo,'')FROM [ShippingHead] where [ListDate]='" + ListDate_LB.SelectedItem + "' and [ProductName]='" + ProductName_CB.SelectedItem + "'and [vchBoxs]='" + WhereBox_LB.SelectedItem + "'";
                    cmd = new SqlCommand(selectCmd, conn);
                    using (reader = cmd.ExecuteReader())
                    {
                        while (reader.Read())
                        {
                            Client = reader.GetString(0).Trim();
                            //載入客戶產品名稱
                            oSheet.Cells[1, 7] = reader.GetString(reader.GetOrdinal("PartDescription"));

                            //載入客戶產品型號
                            oSheet.Cells[2, 7] = reader.GetString(3);

                            //載入一箱幾隻
                            oSheet.Cells[4, 7] = Getcount;

                            //載入箱號
                            oSheet.Cells[10, 2] = reader.GetString(4);

                            //20200410 加入PO
                            oSheet.Cells[5, 9] = CustomerPO_L.Text;

                            //載入客戶名稱
                            oSheet.Cells[3, 7] = reader.GetString(0);

                            //載入棧板編號
                            oSheet.Cells[10, 8] = reader.GetString(5);

                            if (reader.GetString(0).Trim().CompareTo("Wicked Sportz") == 0)
                            {
                                //該客戶要其自己的logo
                                oSheet.Shapes.AddPicture(Application.StartupPath + @".\LOGO-ENAIRGY_Wicked Sportz.jpg", Microsoft.Office.Core.MsoTriState.msoFalse,
                                                Microsoft.Office.Core.MsoTriState.msoTrue, 12, 17, 212, 125);
                                // Application.StartupPath + @".\LOGO-ENAIRGY_Wicked Sportz.jpg"
                            }
                            if (reader.GetString(0).Trim().CompareTo("達成數位") == 0)
                            {
                                //該客戶要其自己的logo
                                oSheet.Shapes.AddPicture(Application.StartupPath + @".\LOGO_DCT.jpg", Microsoft.Office.Core.MsoTriState.msoFalse,
                                                Microsoft.Office.Core.MsoTriState.msoTrue, 12, 17, 212, 125);
                                // Application.StartupPath + @".\LOGO-ENAIRGY_Wicked Sportz.jpg"
                            }
                            else if (reader.GetString(reader.GetOrdinal("Client")).Trim().Contains("ADRENALICIA S.L."))
                            {
                                oSheet.Shapes.AddPicture(Application.StartupPath + @".\LOGO_RogerSports.png", Microsoft.Office.Core.MsoTriState.msoFalse,
                            Microsoft.Office.Core.MsoTriState.msoTrue, 2, 17, 212, 125);
                            }
                            else if (Client.ToUpper().StartsWith("EMB"))
                            {
                                oSheet.Shapes.AddPicture(Application.StartupPath + @".\LOGO_EMB.png", Microsoft.Office.Core.MsoTriState.msoFalse,
                                                    Microsoft.Office.Core.MsoTriState.msoTrue, 2, 17, 212, 125);
                            }

                            else if (Client.ToUpper().StartsWith("PAINTBALL SPORTS"))
                            {
                                oSheet.Shapes.AddPicture(Application.StartupPath + @".\LOGO_Paintball Sports GmbH.png", Microsoft.Office.Core.MsoTriState.msoFalse,
                                                    Microsoft.Office.Core.MsoTriState.msoTrue, 2, 17, 212, 125);
                            }
                        }
                    }
                }

                int serialnooneX = 10, serialnooneY = 239;
                string serialnooneadd = @"C:\SerialNoCode\";

                string FirstCNO = "";

                //載入嘜頭氣瓶序號位子
                using (conn = new SqlConnection(myConnectionString))
                {
                    conn.Open();

                    selectCmd = "SELECT  * FROM [ShippingBody] where [ListDate]='" + ListDate_LB.SelectedItem + "' and [ProductName]='" + ProductName_CB.SelectedItem + "'and [WhereBox]='" + WhereBox_LB.SelectedItem + "'";
                    cmd = new SqlCommand(selectCmd, conn);
                    using (reader = cmd.ExecuteReader())
                    {
                        while (reader.Read())
                        {
                            //serialnooneX = 10; serialnooneY = 239;
                            serialnooneX = 1; serialnooneY = 207;
                            switch (reader.GetString(5))
                            {
                                case "1":
                                    oSheet.Cells[6, 1] = reader.GetString(3);
                                    FirstCNO = reader.GetString(3);
                                    break;

                                case "2":
                                    oSheet.Cells[6, 3] = reader.GetString(3);
                                    break;

                                case "3":
                                    oSheet.Cells[6, 5] = reader.GetString(3);
                                    break;

                                case "4":
                                    oSheet.Cells[6, 7] = reader.GetString(3);
                                    break;

                                case "5":
                                    oSheet.Cells[7, 1] = reader.GetString(3);
                                    break;

                                case "6":
                                    oSheet.Cells[7, 3] = reader.GetString(3);
                                    break;

                                case "7":
                                    oSheet.Cells[7, 5] = reader.GetString(3);
                                    break;

                                case "8":
                                    oSheet.Cells[7, 7] = reader.GetString(3);
                                    break;

                                case "9":
                                    oSheet.Cells[8, 1] = reader.GetString(3);
                                    break;

                                case "10":
                                    oSheet.Cells[8, 3] = reader.GetString(3);
                                    break;

                                case "11":
                                    oSheet.Cells[8, 5] = reader.GetString(3);
                                    break;

                                case "12":
                                    oSheet.Cells[8, 7] = reader.GetString(3);
                                    break;

                                case "13":
                                    oSheet.Cells[9, 1] = reader.GetString(3);
                                    break;

                                case "14":
                                    oSheet.Cells[9, 3] = reader.GetString(3);
                                    break;

                                case "15":
                                    oSheet.Cells[9, 5] = reader.GetString(3);
                                    break;

                                case "16":
                                    oSheet.Cells[9, 7] = reader.GetString(3);
                                    break;
                            }
                            serialnooneX = serialnooneX + ((Convert.ToInt32(reader.GetString(5)) + 3) % 4) * 156;
                            serialnooneY = serialnooneY + ((Convert.ToInt32(reader.GetString(5)) - 1) / 4) * 56;
                            oSheet.Shapes.AddPicture(serialnooneadd + reader.GetString(3) + ".png", Microsoft.Office.Core.MsoTriState.msoFalse,
                            Microsoft.Office.Core.MsoTriState.msoTrue, serialnooneX, serialnooneY, 37, 37);
                        }
                    }

                    if ((Client.Contains("Scientific Gas Australia Pty Ltd") || Client == "Airtanks Limited") && PackingMarks.Trim().CompareTo("SGA-GLADIATAIR") == 0)
                    {
                        string ProductNO = "";
                        //該客戶要其自己的logo  PartNo   Part Description
                        selectCmd = "SELECT  Product_NO FROM MSNBody,Manufacturing where [CylinderNo]='" + FirstCNO + "' and vchManufacturingNo=  Manufacturing_NO";
                        cmd = new SqlCommand(selectCmd, conn);
                        using (reader = cmd.ExecuteReader())
                        {
                            if (reader.Read())
                            {
                                ProductNO = reader.GetValue(0).ToString();
                            }
                        }

                        selectCmd = "SELECT  ProductCode, ProductDescription FROM CustomerPackingMark where ProductNo='" + ProductNO + "' and LogoType='" + (PackingMarks.Trim().Contains("-") == true ? PackingMarks.Trim().Split('-')[1].Trim().ToUpper() : PackingMarks.Trim()) + "'";
                        cmd = new SqlCommand(selectCmd, conn);
                        using (reader = cmd.ExecuteReader())
                        {
                            if (reader.Read())
                            {
                                //載入客戶產品名稱
                                oSheet.Cells[1, 7] = reader.GetString(1);

                                //載入客戶產品型號
                                oSheet.Cells[2, 7] = reader.GetString(0);
                            }
                        }

                        //LOGO
                        oSheet.Shapes.AddPicture(Application.StartupPath + @".\LOGO_SGA_GLADIATAIR.png", Microsoft.Office.Core.MsoTriState.msoFalse,
                                            Microsoft.Office.Core.MsoTriState.msoTrue, 3, 17, 212, 125);
                    }
                    else if (Client.ToUpper().StartsWith("EMB"))
                    {
                        oSheet.Shapes.AddPicture(Application.StartupPath + @".\LOGO_EMB.png", Microsoft.Office.Core.MsoTriState.msoFalse,
                                            Microsoft.Office.Core.MsoTriState.msoTrue, 2, 17, 212, 125);
                    }
                    else if ((Client.ToUpper().StartsWith("HATSAN") == true) && PackingMarks.Trim().CompareTo("HATSAN") == 0)
                    {
                        //LOGO
                        oSheet.Shapes.AddPicture(Application.StartupPath + @".\LOGO_HATSAN.png", Microsoft.Office.Core.MsoTriState.msoFalse,
                                            Microsoft.Office.Core.MsoTriState.msoTrue, 3, 17, 212, 125);
                    }

                    else if (Client.ToUpper().StartsWith("PAINTBALL SPORTS"))
                    {
                        oSheet.Shapes.AddPicture(Application.StartupPath + @".\LOGO_Paintball Sports GmbH.png", Microsoft.Office.Core.MsoTriState.msoFalse,
                                            Microsoft.Office.Core.MsoTriState.msoTrue, 2, 17, 212, 125);
                    }

                    //if (StorageStatus == "N")//20190212
                    {

                        //預設位子在X:680,Y:155
                        //預設QRCODE圖片大小250*250

                        //插入圖片
                        int picX = 680, picY = 185;
                        string picadd = @"C:\QRCode\";

                        selectCmd = "SELECT  * FROM [ShippingHead] where [ListDate]='" + ListDate_LB.SelectedItem + "' and [ProductName]='" + ProductName_CB.Text + "' and [vchBoxs]='" + WhereBox_LB.SelectedItem + "'";
                        cmd = new SqlCommand(selectCmd, conn);
                        using (reader = cmd.ExecuteReader())
                        {
                            while (reader.Read())
                            {
                                Excel.Worksheet xSheet = (Excel.Worksheet)oWB.Sheets[1];
                                oSheet.Shapes.AddPicture(picadd + (reader.GetString(0) + reader.GetString(1) + reader.GetString(3)) + ".png", Microsoft.Office.Core.MsoTriState.msoFalse,
                                Microsoft.Office.Core.MsoTriState.msoTrue, picX, picY, 250, 250);
                                if (picX == 885)
                                {
                                    picY += 70;
                                    picX = 125;
                                }
                                else
                                {
                                    picX += 190;
                                }
                            }
                        }
                    }
                }
            }
            else if (Aboxof == "10")
            {
                string HowMuch = "";
                int Cumulative = 0;
                int Total = 0;

                //載入嘜頭資料
                using (conn = new SqlConnection(myConnectionString))
                {
                    conn.Open();

                    selectCmd = "SELECT  * FROM [ShippingHead] where [ListDate]='" + ListDate_LB.SelectedItem + "' and [ProductName]='" + ProductName_CB.SelectedItem + "'";
                    cmd = new SqlCommand(selectCmd, conn);
                    using (reader = cmd.ExecuteReader())
                    {
                        while (reader.Read())
                        {
                            HowMuch = reader.GetString(4);
                            Cumulative++;
                        }
                    }

                    Total = Convert.ToInt32(HowMuch) * Cumulative;

                    //載入嘜頭資料

                    selectCmd = "SELECT  [Client],(vchPrint + ' ' + ProductName + ' ' + Marking) PartDescription,isnull([CustomerProductName],''),isnull([CustomerProductNo],''),[vchBoxs],isnull(PalletNo,'') FROM [ShippingHead] where [ListDate]='" + ListDate_LB.SelectedItem + "' and [ProductName]='" + ProductName_CB.SelectedItem + "'and [vchBoxs]='" + WhereBox_LB.SelectedItem + "'";
                    cmd = new SqlCommand(selectCmd, conn);
                    using (reader = cmd.ExecuteReader())
                    {
                        while (reader.Read())
                        {
                            Client = reader.GetString(0).Trim();
                            //載入客戶產品名稱
                            oSheet.Cells[1, 8] = reader.GetString(2);

                            //載入客戶產品型號
                            oSheet.Cells[2, 8] = reader.GetString(3);

                            //載入一箱幾隻
                            oSheet.Cells[4, 8] = Getcount;

                            //載入箱號
                            oSheet.Cells[10, 2] = reader.GetString(4);

                            //20200410 加入PO
                            oSheet.Cells[5, 11] = CustomerPO_L.Text;

                            //載入客戶名稱
                            oSheet.Cells[3, 8] = reader.GetString(0);

                            //載入棧板編號
                            oSheet.Cells[10, 10] = reader.GetString(5);

                            if (reader.GetString(0).Trim().CompareTo("Wicked Sportz") == 0)
                            {
                                //該客戶要其自己的logo
                                oSheet.Shapes.AddPicture(Application.StartupPath + @".\LOGO-ENAIRGY_Wicked Sportz.jpg", Microsoft.Office.Core.MsoTriState.msoFalse,
                                                Microsoft.Office.Core.MsoTriState.msoTrue, 19, 17, 212, 125);
                                // Application.StartupPath + @".\LOGO-ENAIRGY_Wicked Sportz.jpg"
                            }
                            if (reader.GetString(0).Trim().CompareTo("達成數位") == 0)
                            {
                                //該客戶要其自己的logo
                                oSheet.Shapes.AddPicture(Application.StartupPath + @".\LOGO_DCT.jpg", Microsoft.Office.Core.MsoTriState.msoFalse,
                                                Microsoft.Office.Core.MsoTriState.msoTrue, 19, 17, 212, 125);
                                // Application.StartupPath + @".\LOGO-ENAIRGY_Wicked Sportz.jpg"
                            }
                            else if (reader.GetString(reader.GetOrdinal("Client")).Trim().Contains("ADRENALICIA S.L."))
                            {
                                oSheet.Shapes.AddPicture(Application.StartupPath + @".\LOGO_RogerSports.png", Microsoft.Office.Core.MsoTriState.msoFalse,
                            Microsoft.Office.Core.MsoTriState.msoTrue, 2, 17, 212, 125);
                            }
                            else if (Client.ToUpper().StartsWith("EMB"))
                            {
                                oSheet.Shapes.AddPicture(Application.StartupPath + @".\LOGO_EMB.png", Microsoft.Office.Core.MsoTriState.msoFalse,
                                                    Microsoft.Office.Core.MsoTriState.msoTrue, 2, 17, 212, 125);
                            }

                            else if (Client.ToUpper().StartsWith("PAINTBALL SPORTS"))
                            {
                                oSheet.Shapes.AddPicture(Application.StartupPath + @".\LOGO_Paintball Sports GmbH.png", Microsoft.Office.Core.MsoTriState.msoFalse,
                                                    Microsoft.Office.Core.MsoTriState.msoTrue, 2, 17, 212, 125);
                            }
                        }
                    }
                }

                int serialnooneX = 10, serialnooneY = 239;
                string serialnooneadd = @"C:\SerialNoCode\";

                string FirstCNO = "";

                //載入嘜頭氣瓶序號位子
                using (conn = new SqlConnection(myConnectionString))
                {
                    conn.Open();

                    selectCmd = "SELECT  * FROM [ShippingBody] where [ListDate]='" + ListDate_LB.SelectedItem + "' and [ProductName]='" + ProductName_CB.SelectedItem + "'and [WhereBox]='" + WhereBox_LB.SelectedItem + "'";
                    cmd = new SqlCommand(selectCmd, conn);
                    using (reader = cmd.ExecuteReader())
                    {
                        while (reader.Read())
                        {
                            serialnooneX = 35; serialnooneY = 219;
                            switch (reader.GetString(5))
                            {
                                case "1":
                                    oSheet.Cells[6, 1] = reader.GetString(3);
                                    FirstCNO = reader.GetString(3);
                                    break;

                                case "2":
                                    oSheet.Cells[6, 3] = reader.GetString(3);
                                    break;

                                case "3":
                                    oSheet.Cells[6, 5] = reader.GetString(3);
                                    break;

                                case "4":
                                    oSheet.Cells[6, 7] = reader.GetString(3);
                                    break;

                                case "5":
                                    oSheet.Cells[6, 9] = reader.GetString(3);
                                    break;

                                case "6":
                                    oSheet.Cells[8, 1] = reader.GetString(3);
                                    break;

                                case "7":
                                    oSheet.Cells[8, 3] = reader.GetString(3);
                                    break;

                                case "8":
                                    oSheet.Cells[8, 5] = reader.GetString(3);
                                    break;

                                case "9":
                                    oSheet.Cells[8, 7] = reader.GetString(3);
                                    break;

                                case "10":
                                    oSheet.Cells[8, 9] = reader.GetString(3);
                                    break;
                            }
                            int i = Convert.ToInt32(reader.GetString(5));
                            i = i > 5 ? i - 5 : i;
                            serialnooneX = 35 + (i - 1) * 127;
                            serialnooneY = serialnooneY + ((Convert.ToInt32(reader.GetString(5)) - 1) / 5) * 111;
                            oSheet.Shapes.AddPicture(serialnooneadd + reader.GetString(3) + ".png", Microsoft.Office.Core.MsoTriState.msoFalse,
                            Microsoft.Office.Core.MsoTriState.msoTrue, serialnooneX, serialnooneY, 60, 60);//, 110, 25);
                        }
                    }

                    if ((Client.Contains("Scientific Gas Australia Pty Ltd") || Client == "Airtanks Limited") && PackingMarks.Trim().CompareTo("SGA-GLADIATAIR") == 0)
                    {
                        string ProductNO = "";

                        //該客戶要其自己的logo  PartNo   Part Description
                        selectCmd = "SELECT  Product_NO FROM MSNBody,Manufacturing where [CylinderNo]='" + FirstCNO + "' and vchManufacturingNo=  Manufacturing_NO";
                        cmd = new SqlCommand(selectCmd, conn);
                        using (reader = cmd.ExecuteReader())
                        {
                            if (reader.Read())
                            {
                                ProductNO = reader.GetValue(0).ToString();
                            }
                        }

                        selectCmd = "SELECT  ProductCode, ProductDescription FROM CustomerPackingMark where ProductNo='" + ProductNO + "' and LogoType='" + (PackingMarks.Trim().Contains("-") == true ? PackingMarks.Trim().Split('-')[1].Trim().ToUpper() : PackingMarks.Trim()) + "'";
                        cmd = new SqlCommand(selectCmd, conn);
                        using (reader = cmd.ExecuteReader())
                        {
                            if (reader.Read())
                            {
                                //載入客戶產品名稱
                                oSheet.Cells[1, 8] = reader.GetString(1);

                                //載入客戶產品型號
                                oSheet.Cells[2, 8] = reader.GetString(0);
                            }
                        }

                        //LOGO
                        oSheet.Shapes.AddPicture(Application.StartupPath + @".\LOGO_SGA_GLADIATAIR.png", Microsoft.Office.Core.MsoTriState.msoFalse,
                                            Microsoft.Office.Core.MsoTriState.msoTrue, 3, 17, 212, 125);
                    }
                    else if (Client.ToUpper().StartsWith("EMB"))
                    {
                        oSheet.Shapes.AddPicture(Application.StartupPath + @".\LOGO_EMB.png", Microsoft.Office.Core.MsoTriState.msoFalse,
                                            Microsoft.Office.Core.MsoTriState.msoTrue, 2, 17, 212, 125);
                    }
                    else if ((Client.ToUpper().StartsWith("HATSAN") == true) && PackingMarks.Trim().CompareTo("HATSAN") == 0)
                    {
                        //LOGO
                        oSheet.Shapes.AddPicture(Application.StartupPath + @".\LOGO_HATSAN.png", Microsoft.Office.Core.MsoTriState.msoFalse,
                                            Microsoft.Office.Core.MsoTriState.msoTrue, 3, 17, 212, 125);
                    }

                    else if (Client.ToUpper().StartsWith("PAINTBALL SPORTS"))
                    {
                        oSheet.Shapes.AddPicture(Application.StartupPath + @".\LOGO_Paintball Sports GmbH.png", Microsoft.Office.Core.MsoTriState.msoFalse,
                                            Microsoft.Office.Core.MsoTriState.msoTrue, 2, 17, 212, 125);
                    }

                    //if (StorageStatus == "N")//20190212
                    {
                        //預設位子在X:680,Y:155
                        //預設QRCODE圖片大小250*250

                        //插入圖片
                        int picX = 680, picY = 185;
                        string picadd = @"C:\QRCode\";

                        selectCmd = "SELECT  * FROM [ShippingHead] where [ListDate]='" + ListDate_LB.SelectedItem + "' and [ProductName]='" + ProductName_CB.Text + "' and [vchBoxs]='" + WhereBox_LB.SelectedItem + "'";
                        cmd = new SqlCommand(selectCmd, conn);
                        using (reader = cmd.ExecuteReader())
                        {
                            while (reader.Read())
                            {
                                Excel.Worksheet xSheet = (Excel.Worksheet)oWB.Sheets[1];
                                oSheet.Shapes.AddPicture(picadd + (reader.GetString(0) + reader.GetString(1) + reader.GetString(3)) + ".png", Microsoft.Office.Core.MsoTriState.msoFalse,
                                Microsoft.Office.Core.MsoTriState.msoTrue, picX, picY, 250, 250);
                                if (picX == 885)
                                {
                                    picY += 70;
                                    picX = 125;
                                }
                                else
                                {
                                    picX += 190;
                                }
                            }
                        }
                    }
                }
            }
            else if (Aboxof == "25")
            {
                //載入嘜頭資料
                using (conn = new SqlConnection(myConnectionString))
                {
                    conn.Open();

                    selectCmd = "SELECT  [Client],(vchPrint + ' ' + ProductName + ' ' + Marking) PartDescription,isnull([CustomerProductName],''),isnull([CustomerProductNo],''),[vchBoxs],isnull(PalletNo,'') FROM [ShippingHead] where [ListDate]='" + ListDate_LB.SelectedItem + "' and [ProductName]='" + ProductName_CB.SelectedItem + "'and [vchBoxs]='" + WhereBox_LB.SelectedItem + "'";
                    cmd = new SqlCommand(selectCmd, conn);
                    using (reader = cmd.ExecuteReader())
                    {
                        while (reader.Read())
                        {
                            Client = reader.GetString(0).Trim();
                            //載入客戶產品名稱
                            oSheet.Cells[1, 7] = reader.GetString(reader.GetOrdinal("PartDescription"));

                            //載入客戶產品型號
                            oSheet.Cells[2, 7] = reader.GetString(3);

                            //載入一箱幾隻
                            oSheet.Cells[4, 7] = Getcount;

                            //載入箱號
                            oSheet.Cells[11, 2] = reader.GetString(4);

                            //20200410 加入PO
                            oSheet.Cells[5, 11] = CustomerPO_L.Text;

                            //載入客戶名稱
                            oSheet.Cells[3, 7] = reader.GetString(0);

                            //載入棧板號
                            oSheet.Cells[11, 10] = reader.GetString(5);

                            if (reader.GetString(0).Trim().CompareTo("Wicked Sportz") == 0)
                            {
                                //該客戶要其自己的logo
                                oSheet.Shapes.AddPicture(Application.StartupPath + @".\LOGO-ENAIRGY_Wicked Sportz.jpg", Microsoft.Office.Core.MsoTriState.msoFalse,
                                                Microsoft.Office.Core.MsoTriState.msoTrue, 2, 17, 212, 125);
                                // Application.StartupPath + @".\LOGO-ENAIRGY_Wicked Sportz.jpg"
                            }
                            if (reader.GetString(0).Trim().CompareTo("達成數位") == 0)
                            {
                                //該客戶要其自己的logo
                                oSheet.Shapes.AddPicture(Application.StartupPath + @".\LOGO_DCT.jpg", Microsoft.Office.Core.MsoTriState.msoFalse,
                                                Microsoft.Office.Core.MsoTriState.msoTrue, 2, 17, 212, 125);
                                // Application.StartupPath + @".\LOGO-ENAIRGY_Wicked Sportz.jpg"
                            }
                            else if (reader.GetString(reader.GetOrdinal("Client")).Trim().Contains("ADRENALICIA S.L."))
                            {
                                oSheet.Shapes.AddPicture(Application.StartupPath + @".\LOGO_RogerSports.png", Microsoft.Office.Core.MsoTriState.msoFalse,
                            Microsoft.Office.Core.MsoTriState.msoTrue, 2, 17, 212, 125);
                            }
                            else if (Client.ToUpper().StartsWith("EMB"))
                            {
                                oSheet.Shapes.AddPicture(Application.StartupPath + @".\LOGO_EMB.png", Microsoft.Office.Core.MsoTriState.msoFalse,
                                                    Microsoft.Office.Core.MsoTriState.msoTrue, 2, 17, 212, 125);
                            }

                            else if (Client.ToUpper().StartsWith("PAINTBALL SPORTS"))
                            {
                                oSheet.Shapes.AddPicture(Application.StartupPath + @".\LOGO_Paintball Sports GmbH.png", Microsoft.Office.Core.MsoTriState.msoFalse,
                                                    Microsoft.Office.Core.MsoTriState.msoTrue, 2, 17, 212, 125);
                            }
                        }
                    }
                }


                int serialnooneX = 8, serialnooneY = 192;
                string serialnooneadd = @"C:\SerialNoCode\";
                string FirstCNO = "";

                //載入嘜頭氣瓶序號位子
                using (conn = new SqlConnection(myConnectionString))
                {
                    conn.Open();

                    selectCmd = "SELECT  * FROM [ShippingBody] where [ListDate]='" + ListDate_LB.SelectedItem + "' and [ProductName]='" + ProductName_CB.SelectedItem + "'and [WhereBox]='" + WhereBox_LB.SelectedItem + "'";
                    cmd = new SqlCommand(selectCmd, conn);
                    using (reader = cmd.ExecuteReader())
                    {
                        while (reader.Read())
                        {
                            serialnooneX = 3; serialnooneY = 189;
                            switch (reader.GetString(5))
                            {
                                case "1":
                                    oSheet.Cells[6, 1] = reader.GetString(3);
                                    FirstCNO = reader.GetString(3);
                                    break;

                                case "2":
                                    oSheet.Cells[6, 3] = reader.GetString(3);
                                    break;

                                case "3":
                                    oSheet.Cells[6, 5] = reader.GetString(3);
                                    break;

                                case "4":
                                    oSheet.Cells[6, 7] = reader.GetString(3);
                                    break;

                                case "5":
                                    oSheet.Cells[6, 9] = reader.GetString(3);
                                    break;

                                case "6":
                                    oSheet.Cells[7, 1] = reader.GetString(3);
                                    break;

                                case "7":
                                    oSheet.Cells[7, 3] = reader.GetString(3);
                                    break;

                                case "8":
                                    oSheet.Cells[7, 5] = reader.GetString(3);
                                    break;

                                case "9":
                                    oSheet.Cells[7, 7] = reader.GetString(3);
                                    break;

                                case "10":
                                    oSheet.Cells[7, 9] = reader.GetString(3);
                                    break;

                                case "11":
                                    oSheet.Cells[8, 1] = reader.GetString(3);
                                    break;

                                case "12":
                                    oSheet.Cells[8, 3] = reader.GetString(3);
                                    break;

                                case "13":
                                    oSheet.Cells[8, 5] = reader.GetString(3);
                                    break;

                                case "14":
                                    oSheet.Cells[8, 7] = reader.GetString(3);
                                    break;

                                case "15":
                                    oSheet.Cells[8, 9] = reader.GetString(3);
                                    break;

                                case "16":
                                    oSheet.Cells[9, 1] = reader.GetString(3);
                                    break;

                                case "17":
                                    oSheet.Cells[9, 3] = reader.GetString(3);
                                    break;

                                case "18":
                                    oSheet.Cells[9, 5] = reader.GetString(3);
                                    break;

                                case "19":
                                    oSheet.Cells[9, 7] = reader.GetString(3);
                                    break;

                                case "20":
                                    oSheet.Cells[9, 9] = reader.GetString(3);
                                    break;

                                case "21":
                                    oSheet.Cells[10, 1] = reader.GetString(3);
                                    break;

                                case "22":
                                    oSheet.Cells[10, 3] = reader.GetString(3);
                                    break;

                                case "23":
                                    oSheet.Cells[10, 5] = reader.GetString(3);
                                    break;

                                case "24":
                                    oSheet.Cells[10, 7] = reader.GetString(3);
                                    break;

                                case "25":
                                    oSheet.Cells[10, 9] = reader.GetString(3);
                                    break;
                            }
                            serialnooneX = serialnooneX + ((Convert.ToInt32(reader.GetString(5)) + 3) % 5) * 144;
                            serialnooneY = serialnooneY + ((Convert.ToInt32(reader.GetString(5)) - 1) / 5) * 47;
                            oSheet.Shapes.AddPicture(serialnooneadd + reader.GetString(3) + ".png", Microsoft.Office.Core.MsoTriState.msoFalse,
                            Microsoft.Office.Core.MsoTriState.msoTrue, serialnooneX, serialnooneY, 40, 40);//, 130, 20);
                        }
                    }

                    if ((Client.Contains("Scientific Gas Australia Pty Ltd") || Client == "Airtanks Limited") && PackingMarks.Trim().CompareTo("SGA-GLADIATAIR") == 0)
                    {
                        string ProductNO = "";

                        //該客戶要其自己的logo  PartNo   Part Description
                        selectCmd = "SELECT  Product_NO FROM MSNBody,Manufacturing where [CylinderNo]='" + FirstCNO + "' and vchManufacturingNo=  Manufacturing_NO";
                        cmd = new SqlCommand(selectCmd, conn);
                        using (reader = cmd.ExecuteReader())
                        {
                            if (reader.Read())
                            {
                                ProductNO = reader.GetValue(0).ToString();
                            }
                        }

                        selectCmd = "SELECT  ProductCode, ProductDescription FROM CustomerPackingMark where ProductNo='" + ProductNO + "' and LogoType='" + (PackingMarks.Trim().Contains("-") == true ? PackingMarks.Trim().Split('-')[1].Trim().ToUpper() : PackingMarks.Trim()) + "'";
                        cmd = new SqlCommand(selectCmd, conn);
                        using (reader = cmd.ExecuteReader())
                        {
                            if (reader.Read())
                            {
                                //載入客戶產品名稱
                                oSheet.Cells[1, 7] = reader.GetString(1);

                                //載入客戶產品型號
                                oSheet.Cells[2, 7] = reader.GetString(0);
                            }
                        }

                        //LOGO
                        oSheet.Shapes.AddPicture(Application.StartupPath + @".\LOGO_SGA_GLADIATAIR.png", Microsoft.Office.Core.MsoTriState.msoFalse,
                                            Microsoft.Office.Core.MsoTriState.msoTrue, 2, 17, 212, 125);
                    }
                    else if (Client.ToUpper().StartsWith("EMB"))
                    {
                        oSheet.Shapes.AddPicture(Application.StartupPath + @".\LOGO_EMB.png", Microsoft.Office.Core.MsoTriState.msoFalse,
                                            Microsoft.Office.Core.MsoTriState.msoTrue, 2, 17, 212, 125);
                    }
                    else if ((Client.ToUpper().StartsWith("HATSAN") == true) && PackingMarks.Trim().CompareTo("HATSAN") == 0)
                    {
                        //LOGO
                        oSheet.Shapes.AddPicture(Application.StartupPath + @".\LOGO_HATSAN.png", Microsoft.Office.Core.MsoTriState.msoFalse,
                                            Microsoft.Office.Core.MsoTriState.msoTrue, 2, 17, 212, 125);
                    }

                    else if (Client.ToUpper().StartsWith("PAINTBALL SPORTS"))
                    {
                        oSheet.Shapes.AddPicture(Application.StartupPath + @".\LOGO_Paintball Sports GmbH.png", Microsoft.Office.Core.MsoTriState.msoFalse,
                                            Microsoft.Office.Core.MsoTriState.msoTrue, 2, 17, 212, 125);
                    }

                    //if (StorageStatus == "N")//20190212
                    {

                        //預設位子在X:680,Y:155
                        //預設QRCODE圖片大小250*250

                        //插入二維條碼
                        int picX = 730, picY = 179;
                        string picadd = @"C:\QRCode\";

                        selectCmd = "SELECT  * FROM [ShippingHead] where [ListDate]='" + ListDate_LB.SelectedItem + "' and [ProductName]='" + ProductName_CB.Text + "' and [vchBoxs]='" + WhereBox_LB.SelectedItem + "'";
                        cmd = new SqlCommand(selectCmd, conn);
                        using (reader = cmd.ExecuteReader())
                        {
                            while (reader.Read())
                            {
                                Excel.Worksheet xSheet = (Excel.Worksheet)oWB.Sheets[1];
                                oSheet.Shapes.AddPicture(picadd + (reader.GetString(0) + reader.GetString(1) + reader.GetString(3)) + ".png", Microsoft.Office.Core.MsoTriState.msoFalse,
                                Microsoft.Office.Core.MsoTriState.msoTrue, picX, picY, 250, 250);
                                if (picX == 885)
                                {
                                    picY += 70;
                                    picX = 125;
                                }
                                else
                                {
                                    picX += 190;
                                }
                            }
                        }
                    }
                }
            }
            else if (Aboxof == "30")
            {
                //載入嘜頭資料
                using (conn = new SqlConnection(myConnectionString))
                {
                    conn.Open();

                    selectCmd = "SELECT  [Client],(vchPrint + ' ' + ProductName + ' ' + Marking) PartDescription,isnull([CustomerProductName],''),isnull([CustomerProductNo],''),[vchBoxs],isnull(PalletNo,'') FROM [ShippingHead] where [ListDate]='" + ListDate_LB.SelectedItem + "' and [ProductName]='" + ProductName_CB.SelectedItem + "'and [vchBoxs]='" + WhereBox_LB.SelectedItem + "'";
                    cmd = new SqlCommand(selectCmd, conn);
                    using (reader = cmd.ExecuteReader())
                    {
                        while (reader.Read())
                        {
                            Client = reader.GetString(0).Trim();
                            //載入客戶產品名稱
                            oSheet.Cells[1, 7] = reader.GetString(reader.GetOrdinal("PartDescription"));

                            //載入客戶產品型號
                            oSheet.Cells[2, 7] = reader.GetString(3);

                            //載入一箱幾隻
                            oSheet.Cells[4, 7] = Getcount;

                            //載入箱號
                            oSheet.Cells[12, 2] = reader.GetString(4);

                            //20200410 加入PO
                            oSheet.Cells[5, 11] = CustomerPO_L.Text;

                            //載入客戶名稱
                            oSheet.Cells[3, 7] = reader.GetString(0);

                            //載入箱號
                            oSheet.Cells[12, 10] = reader.GetString(5);

                            if (reader.GetString(0).Trim().CompareTo("Wicked Sportz") == 0)
                            {
                                //該客戶要其自己的logo
                                oSheet.Shapes.AddPicture(Application.StartupPath + @".\LOGO-ENAIRGY_Wicked Sportz.jpg", Microsoft.Office.Core.MsoTriState.msoFalse,
                                                Microsoft.Office.Core.MsoTriState.msoTrue, 2, 17, 212, 125);
                                // Application.StartupPath + @".\LOGO-ENAIRGY_Wicked Sportz.jpg"
                            }
                            else if (Client.ToUpper().StartsWith("EMB"))
                            {
                                oSheet.Shapes.AddPicture(Application.StartupPath + @".\LOGO_EMB.png", Microsoft.Office.Core.MsoTriState.msoFalse,
                                                    Microsoft.Office.Core.MsoTriState.msoTrue, 2, 17, 212, 125);
                            }
                            if (reader.GetString(0).Trim().CompareTo("達成數位") == 0)
                            {
                                //該客戶要其自己的logo
                                oSheet.Shapes.AddPicture(Application.StartupPath + @".\LOGO_DCT.jpg", Microsoft.Office.Core.MsoTriState.msoFalse,
                                                Microsoft.Office.Core.MsoTriState.msoTrue, 2, 17, 212, 125);
                                // Application.StartupPath + @".\LOGO-ENAIRGY_Wicked Sportz.jpg"
                            }
                            else if (reader.GetString(reader.GetOrdinal("Client")).Trim().Contains("ADRENALICIA S.L."))
                            {
                                oSheet.Shapes.AddPicture(Application.StartupPath + @".\LOGO_RogerSports.png", Microsoft.Office.Core.MsoTriState.msoFalse,
                            Microsoft.Office.Core.MsoTriState.msoTrue, 2, 17, 212, 125);
                            }
                            else if (Client.ToUpper().StartsWith("PAINTBALL SPORTS"))
                            {
                                oSheet.Shapes.AddPicture(Application.StartupPath + @".\LOGO_Paintball Sports GmbH.png", Microsoft.Office.Core.MsoTriState.msoFalse,
                                                    Microsoft.Office.Core.MsoTriState.msoTrue, 2, 17, 212, 125);
                            }
                        }
                    }
                }

                string FirstCNO = "";

                //載入嘜頭氣瓶序號位子
                using (conn = new SqlConnection(myConnectionString))
                {
                    conn.Open();

                    selectCmd = "SELECT  * FROM [ShippingBody] where [ListDate]='" + ListDate_LB.SelectedItem + "' and [ProductName]='" + ProductName_CB.SelectedItem + "'and [WhereBox]='" + WhereBox_LB.SelectedItem + "'";
                    cmd = new SqlCommand(selectCmd, conn);
                    using (reader = cmd.ExecuteReader())
                    {
                        while (reader.Read())
                        {
                            if (FirstCNO == "")
                            {
                                FirstCNO = reader.GetString(3);
                            }

                            if (reader.IsDBNull(5) == false && (Convert.ToInt32(reader.GetString(5)) >= 1 && Convert.ToInt32(reader.GetString(5)) <= 30))
                            {
                                if ((Convert.ToInt32(reader.GetString(5)) - 1) % 5 <= 5)
                                {
                                    oSheet.Cells[6 + (Convert.ToInt32(reader.GetString(5)) - 1) / 5, 1 + ((Convert.ToInt32(reader.GetString(5)) - 1) % 5) * 2] = reader.GetString(3);
                                }

                                //oSheet.Cells[6 + ((Convert.ToInt32(reader.GetString(5)) - 1) / 9), 1 + ((Convert.ToInt32(reader.GetString(5)) - 1) % 9) * 2] = reader.GetString(3);
                            }

                            //serialnooneX = serialnooneX + ((Convert.ToInt32(reader.GetString(5)) + 3) % 5) * 143;
                            //serialnooneY = serialnooneY + ((Convert.ToInt32(reader.GetString(5)) - 1) / 5) * 46;
                            //oSheet.Shapes.AddPicture(serialnooneadd + reader.GetString(3) + ".png", Microsoft.Office.Core.MsoTriState.msoFalse,
                            //Microsoft.Office.Core.MsoTriState.msoTrue, serialnooneX, serialnooneY, 130, 20);
                        }
                    }

                    if ((Client.Contains("Scientific Gas Australia Pty Ltd") || Client == "Airtanks Limited") && PackingMarks.Trim().CompareTo("SGA-GLADIATAIR") == 0)
                    {
                        string ProductNO = "";
                        //該客戶要其自己的logo  PartNo   Part Description
                        selectCmd = "SELECT  Product_NO FROM MSNBody,Manufacturing where [CylinderNo]='" + FirstCNO + "' and vchManufacturingNo=  Manufacturing_NO";
                        cmd = new SqlCommand(selectCmd, conn);
                        using (reader = cmd.ExecuteReader())
                        {
                            if (reader.Read())
                            {
                                ProductNO = reader.GetValue(0).ToString();
                            }
                        }

                        selectCmd = "SELECT  ProductCode, ProductDescription FROM CustomerPackingMark where ProductNo='" + ProductNO + "' and LogoType='" + (PackingMarks.Trim().Contains("-") == true ? PackingMarks.Trim().Split('-')[1].Trim().ToUpper() : PackingMarks.Trim()) + "'";
                        cmd = new SqlCommand(selectCmd, conn);
                        using (reader = cmd.ExecuteReader())
                        {
                            if (reader.Read())
                            {
                                //載入客戶產品名稱
                                oSheet.Cells[1, 7] = reader.GetString(1);

                                //載入客戶產品型號
                                oSheet.Cells[2, 7] = reader.GetString(0);
                            }
                        }

                        //LOGO
                        oSheet.Shapes.AddPicture(Application.StartupPath + @".\LOGO_SGA_GLADIATAIR.png", Microsoft.Office.Core.MsoTriState.msoFalse,
                                            Microsoft.Office.Core.MsoTriState.msoTrue, 3, 17, 212, 125);
                    }
                    else if ((Client.ToUpper().StartsWith("HATSAN") == true) && PackingMarks.Trim().CompareTo("HATSAN") == 0)
                    {
                        //LOGO
                        oSheet.Shapes.AddPicture(Application.StartupPath + @".\LOGO_HATSAN.png", Microsoft.Office.Core.MsoTriState.msoFalse,
                                            Microsoft.Office.Core.MsoTriState.msoTrue, 3, 17, 212, 125);
                    }
                    else if (Client.ToUpper().StartsWith("EMB"))
                    {
                        oSheet.Shapes.AddPicture(Application.StartupPath + @".\LOGO_EMB.png", Microsoft.Office.Core.MsoTriState.msoFalse,
                                            Microsoft.Office.Core.MsoTriState.msoTrue, 2, 17, 212, 125);
                    }

                    //if (StorageStatus == "N")//20190212
                    {
                        //預設位子在X:680,Y:155
                        //預設QRCODE圖片大小250*250

                        //插入二維條碼
                        int picX = 730, picY = 179;
                        string picadd = @"C:\QRCode\";

                        selectCmd = "SELECT  * FROM [ShippingHead] where [ListDate]='" + ListDate_LB.SelectedItem + "' and [ProductName]='" + ProductName_CB.Text + "' and [vchBoxs]='" + WhereBox_LB.SelectedItem + "'";
                        cmd = new SqlCommand(selectCmd, conn);
                        using (reader = cmd.ExecuteReader())
                        {
                            while (reader.Read())
                            {
                                Excel.Worksheet xSheet = (Excel.Worksheet)oWB.Sheets[1];
                                oSheet.Shapes.AddPicture(picadd + (reader.GetString(0) + reader.GetString(1) + reader.GetString(3)) + ".png", Microsoft.Office.Core.MsoTriState.msoFalse,
                                Microsoft.Office.Core.MsoTriState.msoTrue, picX, picY, 250, 250);
                                if (picX == 885)
                                {
                                    picY += 70;
                                    picX = 125;
                                }
                                else
                                {
                                    picX += 190;
                                }
                            }
                        }
                    }
                }
            }
            else if (Aboxof == "117")
            {
                //載入嘜頭資料
                using (conn = new SqlConnection(myConnectionString))
                {
                    conn.Open();

                    selectCmd = "SELECT  [Client],(vchPrint + ' ' + ProductName + ' ' + Marking) PartDescription,isnull([CustomerProductName],''),isnull([CustomerProductNo],''),[vchBoxs],isnull(PalletNo,'') FROM [ShippingHead] where [ListDate]='" + ListDate_LB.SelectedItem + "' and [ProductName]='" + ProductName_CB.SelectedItem + "'and [vchBoxs]='" + WhereBox_LB.SelectedItem + "'";
                    cmd = new SqlCommand(selectCmd, conn);
                    using (reader = cmd.ExecuteReader())
                    {
                        while (reader.Read())
                        {
                            Client = reader.GetString(0).Trim();
                            //載入客戶產品名稱
                            oSheet.Cells[1, 9] = reader.GetString(2);

                            //載入客戶產品型號
                            oSheet.Cells[2, 9] = reader.GetString(3);

                            //載入一箱幾隻
                            oSheet.Cells[4, 9] = Getcount;

                            //載入箱號
                            oSheet.Cells[19, 2] = reader.GetString(4);

                            //20200410 加入PO
                            oSheet.Cells[5, 9] = CustomerPO_L.Text;

                            //載入客戶名稱
                            oSheet.Cells[3, 9] = reader.GetString(0);
                            oSheet.Cells[19, 11] = reader.GetString(5);

                            if (reader.GetString(0).Trim().CompareTo("Wicked Sportz") == 0)
                            {
                                //該客戶要其自己的logo
                                oSheet.Shapes.AddPicture(Application.StartupPath + @".\LOGO-ENAIRGY_Wicked Sportz.jpg", Microsoft.Office.Core.MsoTriState.msoFalse,
                                                Microsoft.Office.Core.MsoTriState.msoTrue, 2, 17, 212, 125);
                                // Application.StartupPath + @".\LOGO-ENAIRGY_Wicked Sportz.jpg"
                            }
                            if (reader.GetString(0).Trim().CompareTo("達成數位") == 0)
                            {
                                //該客戶要其自己的logo
                                oSheet.Shapes.AddPicture(Application.StartupPath + @".\LOGO_DCT.jpg", Microsoft.Office.Core.MsoTriState.msoFalse,
                                                Microsoft.Office.Core.MsoTriState.msoTrue, 2, 17, 212, 125);
                                // Application.StartupPath + @".\LOGO-ENAIRGY_Wicked Sportz.jpg"
                            }
                            else if (reader.GetString(reader.GetOrdinal("Client")).Trim().Contains("ADRENALICIA S.L."))
                            {
                                oSheet.Shapes.AddPicture(Application.StartupPath + @".\LOGO_RogerSports.png", Microsoft.Office.Core.MsoTriState.msoFalse,
                            Microsoft.Office.Core.MsoTriState.msoTrue, 2, 17, 212, 125);
                            }
                            else if (Client.ToUpper().StartsWith("EMB"))
                            {
                                oSheet.Shapes.AddPicture(Application.StartupPath + @".\LOGO_EMB.png", Microsoft.Office.Core.MsoTriState.msoFalse,
                                                    Microsoft.Office.Core.MsoTriState.msoTrue, 2, 17, 212, 125);
                            }
                        }
                    }
                }

                string FirstCNO = "";

                //載入嘜頭氣瓶序號位子
                using (conn = new SqlConnection(myConnectionString))
                {
                    conn.Open();

                    selectCmd = "SELECT  * FROM [ShippingBody] where [ListDate]='" + ListDate_LB.SelectedItem + "' and [ProductName]='" + ProductName_CB.SelectedItem + "'and [WhereBox]='" + WhereBox_LB.SelectedItem + "'";
                    cmd = new SqlCommand(selectCmd, conn);
                    using (reader = cmd.ExecuteReader())
                    {
                        while (reader.Read())
                        {
                            if (FirstCNO == "")
                            {
                                FirstCNO = reader.GetString(3);
                            }
                            if (reader.IsDBNull(5) == false && (Convert.ToInt32(reader.GetString(5)) >= 1 && Convert.ToInt32(reader.GetString(5)) <= 117))
                            {
                                oSheet.Cells[6 + ((Convert.ToInt32(reader.GetString(5)) - 1) / 9), 1 + ((Convert.ToInt32(reader.GetString(5)) - 1) % 9) * 2] = reader.GetString(3);
                            }
                        }
                    }

                    if ((Client.Contains("Scientific Gas Australia Pty Ltd") || Client == "Airtanks Limited") && PackingMarks.Trim().CompareTo("SGA-GLADIATAIR") == 0)
                    {
                        string ProductNO = "";

                        //該客戶要其自己的logo  PartNo   Part Description
                        selectCmd = "SELECT  Product_NO FROM MSNBody,Manufacturing where [CylinderNo]='" + FirstCNO + "' and vchManufacturingNo=  Manufacturing_NO";
                        cmd = new SqlCommand(selectCmd, conn);
                        using (reader = cmd.ExecuteReader())
                        {
                            if (reader.Read())
                            {
                                ProductNO = reader.GetValue(0).ToString();
                            }
                        }

                        selectCmd = "SELECT  ProductCode, ProductDescription FROM CustomerPackingMark where ProductNo='" + ProductNO + "' and LogoType='" + (PackingMarks.Trim().Contains("-") == true ? PackingMarks.Trim().Split('-')[1].Trim().ToUpper() : PackingMarks.Trim()) + "'";
                        cmd = new SqlCommand(selectCmd, conn);
                        using (reader = cmd.ExecuteReader())
                        {
                            if (reader.Read())
                            {
                                //載入客戶產品名稱
                                oSheet.Cells[1, 9] = reader.GetString(1);

                                //載入客戶產品型號
                                oSheet.Cells[2, 9] = reader.GetString(0);
                            }
                        }

                        //LOGO
                        oSheet.Shapes.AddPicture(Application.StartupPath + @".\LOGO_SGA_GLADIATAIR.png", Microsoft.Office.Core.MsoTriState.msoFalse,
                                            Microsoft.Office.Core.MsoTriState.msoTrue, 3, 17, 212, 125);
                    }
                    else if ((Client.ToUpper().StartsWith("HATSAN") == true) && PackingMarks.Trim().CompareTo("HATSAN") == 0)
                    {
                        //LOGO
                        oSheet.Shapes.AddPicture(Application.StartupPath + @".\LOGO_HATSAN.png", Microsoft.Office.Core.MsoTriState.msoFalse,
                                            Microsoft.Office.Core.MsoTriState.msoTrue, 3, 17, 212, 125);
                    }
                    else if (Client.ToUpper().StartsWith("EMB"))
                    {
                        oSheet.Shapes.AddPicture(Application.StartupPath + @".\LOGO_EMB.png", Microsoft.Office.Core.MsoTriState.msoFalse,
                                            Microsoft.Office.Core.MsoTriState.msoTrue, 2, 17, 212, 125);
                    }
                }


                //Aboxof == "117"其資料太長，造成QR code 無法全部紀錄，僅序號最多41組
                //if (StorageStatus == "N")
                //{

                //    //預設位子在X:680,Y:155
                //    //預設QRCODE圖片大小250*250

                //    //插入二維條碼

                //    string picadd = @"C:\QRCode\";
                //    
                //    selectCmd = "SELECT  * FROM [ShippingHead] where [ListDate]='" + ListDateListBox.SelectedItem + "' and [ProductName]='" + ProductComboBox.Text + "' and [vchBoxs]='" + BoxsListBox.SelectedItem + "'";
                //    conn = new SqlConnection(myConnectionString);
                //    conn.Open();
                //    cmd = new SqlCommand(selectCmd, conn);
                //    reader = cmd.ExecuteReader();
                //    while (reader.Read())
                //    {
                //        Excel.Worksheet xSheet = (Excel.Worksheet)oWB.Sheets[1];
                //        oSheet.Shapes.AddPicture(picadd + (reader.GetString(0) + reader.GetString(1) + reader.GetString(3)) + ".png", Microsoft.Office.Core.MsoTriState.msoFalse,
                //        Microsoft.Office.Core.MsoTriState.msoTrue, 825, 0, 160, 160);
                //    }
                //}
            }
            else if (Aboxof == "111")
            {
                //載入嘜頭資料
                using (conn = new SqlConnection(myConnectionString))
                {
                    conn.Open();

                    selectCmd = "SELECT  [Client],(vchPrint + ' ' + ProductName + ' ' + Marking) PartDescription,isnull([CustomerProductName],''),isnull([CustomerProductNo],''),[vchBoxs] ,isnull(PalletNo,'') FROM [ShippingHead] where [ListDate]='" + ListDate_LB.SelectedItem + "' and [ProductName]='" + ProductName_CB.SelectedItem + "'and [vchBoxs]='" + WhereBox_LB.SelectedItem + "'";
                    cmd = new SqlCommand(selectCmd, conn);
                    using (reader = cmd.ExecuteReader())
                    {
                        while (reader.Read())
                        {
                            Client = reader.GetString(0).Trim();
                            //載入客戶產品名稱
                            oSheet.Cells[1, 9] = reader.GetString(2);

                            //載入客戶產品型號
                            oSheet.Cells[2, 9] = reader.GetString(3);

                            //載入一箱幾隻
                            oSheet.Cells[4, 9] = Getcount;

                            //載入箱號
                            oSheet.Cells[19, 2] = reader.GetString(4);

                            //20200410 加入PO
                            oSheet.Cells[5, 9] = CustomerPO_L.Text;

                            //載入客戶名稱
                            oSheet.Cells[3, 9] = reader.GetString(0);

                            //載入
                            oSheet.Cells[19, 11] = reader.GetString(5);

                            if (reader.GetString(0).Trim().CompareTo("Wicked Sportz") == 0)
                            {
                                //該客戶要其自己的logo
                                oSheet.Shapes.AddPicture(Application.StartupPath + @".\LOGO-ENAIRGY_Wicked Sportz.jpg", Microsoft.Office.Core.MsoTriState.msoFalse,
                                                Microsoft.Office.Core.MsoTriState.msoTrue, 2, 17, 212, 125);
                                // Application.StartupPath + @".\LOGO-ENAIRGY_Wicked Sportz.jpg"
                            }
                            if (reader.GetString(0).Trim().CompareTo("達成數位") == 0)
                            {
                                //該客戶要其自己的logo
                                oSheet.Shapes.AddPicture(Application.StartupPath + @".\LOGO_DCT.jpg", Microsoft.Office.Core.MsoTriState.msoFalse,
                                                Microsoft.Office.Core.MsoTriState.msoTrue, 2, 17, 212, 125);
                                // Application.StartupPath + @".\LOGO-ENAIRGY_Wicked Sportz.jpg"
                            }
                            else if (reader.GetString(reader.GetOrdinal("Client")).Trim().Contains("ADRENALICIA S.L."))
                            {
                                oSheet.Shapes.AddPicture(Application.StartupPath + @".\LOGO_RogerSports.png", Microsoft.Office.Core.MsoTriState.msoFalse,
                            Microsoft.Office.Core.MsoTriState.msoTrue, 2, 17, 212, 125);
                            }

                            else if (Client.ToUpper().StartsWith("EMB"))
                            {
                                oSheet.Shapes.AddPicture(Application.StartupPath + @".\LOGO_EMB.png", Microsoft.Office.Core.MsoTriState.msoFalse,
                                                    Microsoft.Office.Core.MsoTriState.msoTrue, 2, 17, 212, 125);
                            }
                        }
                    }
                }

                string FirstCNO = "";

                //載入嘜頭氣瓶序號位子
                using (conn = new SqlConnection(myConnectionString))
                {
                    conn.Open();

                    selectCmd = "SELECT  * FROM [ShippingBody] where [ListDate]='" + ListDate_LB.SelectedItem + "' and [ProductName]='" + ProductName_CB.SelectedItem + "'and [WhereBox]='" + WhereBox_LB.SelectedItem + "'";
                    cmd = new SqlCommand(selectCmd, conn);
                    using (reader = cmd.ExecuteReader())
                    {
                        while (reader.Read())
                        {
                            if (FirstCNO == "")
                            {
                                FirstCNO = reader.GetString(3);
                            }
                            if (reader.IsDBNull(5) == false && (Convert.ToInt32(reader.GetString(5)) >= 1 && Convert.ToInt32(reader.GetString(5)) <= 111))
                            {
                                if ((Convert.ToInt32(reader.GetString(5)) - 1) % 17 <= 8)
                                {
                                    //9
                                    oSheet.Cells[6 + (Convert.ToInt32(reader.GetString(5)) - 1) / 17 * 2, 1 + ((Convert.ToInt32(reader.GetString(5)) - 1) % 17) * 2] = reader.GetString(3);
                                }
                                else
                                {
                                    //8
                                    oSheet.Cells[6 + (Convert.ToInt32(reader.GetString(5)) - 1) / 17 * 2 + 1, 2 + ((((Convert.ToInt32(reader.GetString(5)) - 1) % 17) - 8) % 8) * 2] = reader.GetString(3);
                                }
                                //oSheet.Cells[6 + ((Convert.ToInt32(reader.GetString(5)) - 1) / 9), 1 + ((Convert.ToInt32(reader.GetString(5)) - 1) % 9) * 2] = reader.GetString(3);
                            }
                        }
                    }

                    if ((Client.Contains("Scientific Gas Australia Pty Ltd") || Client == "Airtanks Limited") && PackingMarks.Trim().CompareTo("SGA-GLADIATAIR") == 0)
                    {
                        string ProductNO = "";

                        //該客戶要其自己的logo  PartNo   Part Description
                        selectCmd = "SELECT  Product_NO FROM MSNBody,Manufacturing where [CylinderNo]='" + FirstCNO + "' and vchManufacturingNo=  Manufacturing_NO";
                        cmd = new SqlCommand(selectCmd, conn);
                        using (reader = cmd.ExecuteReader())
                        {
                            if (reader.Read())
                            {
                                ProductNO = reader.GetValue(0).ToString();
                            }
                        }

                        selectCmd = "SELECT  ProductCode, ProductDescription FROM CustomerPackingMark where ProductNo='" + ProductNO + "' and LogoType='" + (PackingMarks.Trim().Contains("-") == true ? PackingMarks.Trim().Split('-')[1].Trim().ToUpper() : PackingMarks.Trim()) + "'";
                        cmd = new SqlCommand(selectCmd, conn);
                        using (reader = cmd.ExecuteReader())
                        {
                            if (reader.Read())
                            {
                                //載入客戶產品名稱
                                oSheet.Cells[1, 9] = reader.GetString(1);

                                //載入客戶產品型號
                                oSheet.Cells[2, 9] = reader.GetString(0);
                            }
                        }

                        //LOGO
                        oSheet.Shapes.AddPicture(Application.StartupPath + @".\LOGO_SGA_GLADIATAIR.png", Microsoft.Office.Core.MsoTriState.msoFalse,
                                            Microsoft.Office.Core.MsoTriState.msoTrue, 2, 17, 212, 125);
                    }
                    else if ((Client.ToUpper().StartsWith("HATSAN") == true) && PackingMarks.Trim().CompareTo("HATSAN") == 0)
                    {
                        //LOGO
                        oSheet.Shapes.AddPicture(Application.StartupPath + @".\LOGO_HATSAN.png", Microsoft.Office.Core.MsoTriState.msoFalse,
                                            Microsoft.Office.Core.MsoTriState.msoTrue, 2, 17, 212, 125);
                    }
                    else if (Client.ToUpper().StartsWith("EMB"))
                    {
                        oSheet.Shapes.AddPicture(Application.StartupPath + @".\LOGO_EMB.png", Microsoft.Office.Core.MsoTriState.msoFalse,
                                            Microsoft.Office.Core.MsoTriState.msoTrue, 2, 17, 212, 125);
                    }
                }
            }
            else if (Aboxof == "5")
            {
                string HowMuch = "";
                int Cumulative = 0;
                int Total = 0;

                //載入嘜頭資料
                using (conn = new SqlConnection(myConnectionString))
                {
                    conn.Open();

                    selectCmd = "SELECT  * FROM [ShippingHead] where [ListDate]='" + ListDate_LB.SelectedItem + "' and [ProductName]='" + ProductName_CB.SelectedItem + "'";
                    cmd = new SqlCommand(selectCmd, conn);
                    using (reader = cmd.ExecuteReader())
                    {
                        while (reader.Read())
                        {
                            HowMuch = reader.GetString(4);
                            Cumulative++;
                        }
                    }

                    Total = Convert.ToInt32(HowMuch) * Cumulative;

                    //載入嘜頭資料
                    selectCmd = "SELECT  [Client],(vchPrint + ' ' + ProductName + ' ' + Marking) PartDescription,isnull([CustomerProductName],''),isnull([CustomerProductNo],''),[vchBoxs],isnull(PalletNo,'') FROM [ShippingHead] where [ListDate]='" + ListDate_LB.SelectedItem + "' and [ProductName]='" + ProductName_CB.SelectedItem + "'and [vchBoxs]='" + WhereBox_LB.SelectedItem + "'";
                    cmd = new SqlCommand(selectCmd, conn);
                    using (reader = cmd.ExecuteReader())
                    {
                        while (reader.Read())
                        {
                            Client = reader.GetString(0).Trim();
                            //載入客戶產品名稱
                            oSheet.Cells[1, 7] = reader.GetString(reader.GetOrdinal("PartDescription"));

                            //載入客戶產品型號
                            oSheet.Cells[2, 7] = reader.GetString(3);

                            //載入一箱幾隻
                            oSheet.Cells[4, 7] = Getcount;

                            //載入箱號
                            oSheet.Cells[9, 2] = reader.GetString(4);

                            //20200410 加入PO
                            oSheet.Cells[5, 11] = CustomerPO_L.Text;

                            //載入客戶名稱
                            oSheet.Cells[3, 7] = reader.GetString(0);

                            //棧板號
                            oSheet.Cells[9, 10] = reader.GetString(5);

                            if (reader.GetString(0).Trim().CompareTo("Wicked Sportz") == 0)
                            {
                                //該客戶要其自己的logo
                                oSheet.Shapes.AddPicture(Application.StartupPath + @".\LOGO-ENAIRGY_Wicked Sportz.jpg", Microsoft.Office.Core.MsoTriState.msoFalse,
                                                Microsoft.Office.Core.MsoTriState.msoTrue, 3, 17, 212, 125);
                                // Application.StartupPath + @".\LOGO-ENAIRGY_Wicked Sportz.jpg"
                            }
                            if (reader.GetString(0).Trim().CompareTo("達成數位") == 0)
                            {
                                //該客戶要其自己的logo
                                oSheet.Shapes.AddPicture(Application.StartupPath + @".\LOGO_DCT.jpg", Microsoft.Office.Core.MsoTriState.msoFalse,
                                                Microsoft.Office.Core.MsoTriState.msoTrue, 3, 17, 212, 125);
                                // Application.StartupPath + @".\LOGO-ENAIRGY_Wicked Sportz.jpg"
                            }
                            else if (reader.GetString(reader.GetOrdinal("Client")).Trim().Contains("ADRENALICIA S.L."))
                            {
                                oSheet.Shapes.AddPicture(Application.StartupPath + @".\LOGO_RogerSports.png", Microsoft.Office.Core.MsoTriState.msoFalse,
                            Microsoft.Office.Core.MsoTriState.msoTrue, 2, 17, 212, 125);
                            }
                            else if (Client.ToUpper().StartsWith("EMB"))
                            {
                                oSheet.Shapes.AddPicture(Application.StartupPath + @".\LOGO_EMB.png", Microsoft.Office.Core.MsoTriState.msoFalse,
                                                    Microsoft.Office.Core.MsoTriState.msoTrue, 2, 17, 212, 125);
                            }
                            else if (Client.ToUpper().StartsWith("PAINTBALL SPORTS"))
                            {
                                oSheet.Shapes.AddPicture(Application.StartupPath + @".\LOGO_Paintball Sports GmbH.png", Microsoft.Office.Core.MsoTriState.msoFalse,
                                                    Microsoft.Office.Core.MsoTriState.msoTrue, 2, 17, 212, 125);
                            }
                        }
                    }
                }

                int serialnooneX = 7, serialnooneY = 209;
                string serialnooneadd = @"C:\SerialNoCode\";

                string FirstCNO = "";

                //載入嘜頭氣瓶序號位子
                using (conn = new SqlConnection(myConnectionString))
                {
                    conn.Open();

                    selectCmd = "SELECT  * FROM [ShippingBody] where [ListDate]='" + ListDate_LB.SelectedItem + "' and [ProductName]='" + ProductName_CB.SelectedItem + "'and [WhereBox]='" + WhereBox_LB.SelectedItem + "'";
                    cmd = new SqlCommand(selectCmd, conn);
                    using (reader = cmd.ExecuteReader())
                    {
                        while (reader.Read())
                        {
                            serialnooneX = 3; serialnooneY = 211;
                            switch (reader.GetString(5))
                            {
                                case "1":
                                    oSheet.Cells[6, 1] = reader.GetString(3);
                                    FirstCNO = reader.GetString(3);
                                    break;

                                case "2":
                                    oSheet.Cells[6, 3] = reader.GetString(3);
                                    break;

                                case "3":
                                    oSheet.Cells[6, 5] = reader.GetString(3);
                                    break;

                                case "4":
                                    oSheet.Cells[6, 7] = reader.GetString(3);
                                    break;

                                case "5":
                                    oSheet.Cells[6, 9] = reader.GetString(3);
                                    break;

                                case "6":
                                    oSheet.Cells[7, 1] = reader.GetString(3);
                                    break;

                                case "7":
                                    oSheet.Cells[7, 3] = reader.GetString(3);
                                    break;

                                case "8":
                                    oSheet.Cells[7, 5] = reader.GetString(3);
                                    break;

                                case "9":
                                    oSheet.Cells[7, 7] = reader.GetString(3);
                                    break;

                                case "10":
                                    oSheet.Cells[7, 9] = reader.GetString(3);
                                    break;

                                case "11":
                                    oSheet.Cells[8, 1] = reader.GetString(3);
                                    break;

                                case "12":
                                    oSheet.Cells[8, 3] = reader.GetString(3);
                                    break;

                                case "13":
                                    oSheet.Cells[8, 5] = reader.GetString(3);
                                    break;

                                case "14":
                                    oSheet.Cells[8, 7] = reader.GetString(3);
                                    break;

                                case "15":
                                    oSheet.Cells[8, 9] = reader.GetString(3);
                                    break;
                            }
                            serialnooneX = serialnooneX + ((Convert.ToInt32(reader.GetString(5)) + 4) % 5) * 145;
                            serialnooneY = serialnooneY + ((Convert.ToInt32(reader.GetString(5)) - 1) / 5) * 75;
                            oSheet.Shapes.AddPicture(serialnooneadd + reader.GetString(3) + ".png", Microsoft.Office.Core.MsoTriState.msoFalse,
                            Microsoft.Office.Core.MsoTriState.msoTrue, serialnooneX, serialnooneY, 44, 44);//, 130, 25);
                        }
                    }


                    if ((Client.Contains("Scientific Gas Australia Pty Ltd") || Client == "Airtanks Limited") && PackingMarks.Trim().CompareTo("SGA-GLADIATAIR") == 0)
                    {
                        string ProductNO = "";

                        //該客戶要其自己的logo  PartNo   Part Description
                        selectCmd = "SELECT  Product_NO FROM MSNBody,Manufacturing where [CylinderNo]='" + FirstCNO + "' and vchManufacturingNo=  Manufacturing_NO";
                        cmd = new SqlCommand(selectCmd, conn);
                        using (reader = cmd.ExecuteReader())
                        {
                            if (reader.Read())
                            {
                                ProductNO = reader.GetValue(0).ToString();
                            }
                        }

                        selectCmd = "SELECT  ProductCode, ProductDescription FROM CustomerPackingMark where ProductNo='" + ProductNO + "' and LogoType='" + (PackingMarks.Trim().Contains("-") == true ? PackingMarks.Trim().Split('-')[1].Trim().ToUpper() : PackingMarks.Trim()) + "'";
                        cmd = new SqlCommand(selectCmd, conn);
                        using (reader = cmd.ExecuteReader())
                        {
                            if (reader.Read())
                            {
                                //載入客戶產品名稱
                                oSheet.Cells[1, 7] = reader.GetString(1);

                                //載入客戶產品型號
                                oSheet.Cells[2, 7] = reader.GetString(0);
                            }
                        }

                        //LOGO
                        oSheet.Shapes.AddPicture(Application.StartupPath + @".\LOGO_SGA_GLADIATAIR.png", Microsoft.Office.Core.MsoTriState.msoFalse,
                                            Microsoft.Office.Core.MsoTriState.msoTrue, 3, 17, 212, 125);
                    }
                    else if ((Client.ToUpper().StartsWith("HATSAN") == true) && PackingMarks.Trim().CompareTo("HATSAN") == 0)
                    {
                        //LOGO
                        oSheet.Shapes.AddPicture(Application.StartupPath + @".\LOGO_HATSAN.png", Microsoft.Office.Core.MsoTriState.msoFalse,
                                            Microsoft.Office.Core.MsoTriState.msoTrue, 3, 17, 212, 125);
                    }
                    else if (Client.ToUpper().StartsWith("EMB"))
                    {
                        oSheet.Shapes.AddPicture(Application.StartupPath + @".\LOGO_EMB.png", Microsoft.Office.Core.MsoTriState.msoFalse,
                                            Microsoft.Office.Core.MsoTriState.msoTrue, 2, 17, 212, 125);
                    }
                    else if (Client.ToUpper().StartsWith("PAINTBALL SPORTS"))
                    {
                        oSheet.Shapes.AddPicture(Application.StartupPath + @".\LOGO_Paintball Sports GmbH.png", Microsoft.Office.Core.MsoTriState.msoFalse,
                                            Microsoft.Office.Core.MsoTriState.msoTrue, 2, 17, 212, 125);
                    }

                    //if (StorageStatus == "N")//20190212
                    {
                        //預設位子在X:680,Y:155
                        //預設QRCODE圖片大小250*250

                        //插入圖片
                        int picX = 732, picY = 187;
                        string picadd = @"C:\QRCode\";

                        selectCmd = "SELECT  * FROM [ShippingHead] where [ListDate]='" + ListDate_LB.SelectedItem + "' and [ProductName]='" + ProductName_CB.Text + "' and [vchBoxs]='" + WhereBox_LB.SelectedItem + "'";
                        cmd = new SqlCommand(selectCmd, conn);
                        using (reader = cmd.ExecuteReader())
                        {
                            while (reader.Read())
                            {
                                Excel.Worksheet xSheet = (Excel.Worksheet)oWB.Sheets[1];
                                oSheet.Shapes.AddPicture(picadd + (reader.GetString(0) + reader.GetString(1) + reader.GetString(3)) + ".png", Microsoft.Office.Core.MsoTriState.msoFalse,
                                Microsoft.Office.Core.MsoTriState.msoTrue, picX, picY, 250, 250);
                                if (picX == 885)
                                {
                                    picY += 70;
                                    picX = 125;
                                }
                                else
                                {
                                    picX += 190;
                                }
                            }
                        }
                    }
                }
            }

            if (Aboxof == "1")
            {
                //載入嘜頭資料
                using (conn = new SqlConnection(myConnectionString))
                {
                    conn.Open();

                    selectCmd = "SELECT  [Client],(vchPrint + ' ' + ProductName + ' ' + Marking) PartDescription,isnull([CustomerProductName],''),isnull([CustomerProductNo],''),[vchBoxs],isnull(PalletNo,'') FROM [ShippingHead] where [ListDate]='" + ListDate_LB.SelectedItem + "' and [ProductName]='" + ProductName_CB.SelectedItem + "'and [vchBoxs]='" + WhereBox_LB.SelectedItem + "'";
                    cmd = new SqlCommand(selectCmd, conn);
                    using (reader = cmd.ExecuteReader())
                    {
                        while (reader.Read())
                        {
                            Client = reader.GetString(0).Trim();
                            //載入客戶產品名稱
                            oSheet.Cells[1, 7] = reader.GetString(reader.GetOrdinal("PartDescription"));

                            //載入客戶產品型號
                            oSheet.Cells[2, 7] = reader.GetString(3);

                            //載入一箱幾隻
                            oSheet.Cells[4, 7] = Getcount;

                            //載入箱號
                            //oSheet.Cells[11, 3] = reader.GetString(4);

                            // if (StorageStatus == "N")//20190213
                            {
                                //載入客戶名稱
                                oSheet.Cells[3, 7] = reader.GetString(0);

                                //載入訂單編號(PO)
                                //oSheet.Cells[5, 13] = reader.GetString(1);

                                //載入箱號
                                oSheet.Cells[10, 2] = reader.GetString(4);

                                //20200410 加入PO
                                oSheet.Cells[10, 8] = CustomerPO_L.Text;
                            }
                            if (reader.GetString(0).Trim().CompareTo("Wicked Sportz") == 0)
                            {
                                //該客戶要其自己的logo
                                oSheet.Shapes.AddPicture(Application.StartupPath + @".\LOGO-ENAIRGY_Wicked Sportz.jpg", Microsoft.Office.Core.MsoTriState.msoFalse,
                                                Microsoft.Office.Core.MsoTriState.msoTrue, 12, 17, 212, 125);
                                // Application.StartupPath + @".\LOGO-ENAIRGY_Wicked Sportz.jpg"
                            }
                            if (reader.GetString(0).Trim().CompareTo("達成數位") == 0)
                            {
                                //該客戶要其自己的logo
                                oSheet.Shapes.AddPicture(Application.StartupPath + @".\LOGO_DCT.jpg", Microsoft.Office.Core.MsoTriState.msoFalse,
                                                Microsoft.Office.Core.MsoTriState.msoTrue, 12, 17, 212, 125);
                                // Application.StartupPath + @".\LOGO-ENAIRGY_Wicked Sportz.jpg"
                            }
                            else if (reader.GetString(reader.GetOrdinal("Client")).Trim().Contains("ADRENALICIA S.L."))
                            {
                                oSheet.Shapes.AddPicture(Application.StartupPath + @".\LOGO_RogerSports.png", Microsoft.Office.Core.MsoTriState.msoFalse,
                            Microsoft.Office.Core.MsoTriState.msoTrue, 2, 17, 212, 125);
                            }
                            else if (Client.ToUpper().StartsWith("EMB"))
                            {
                                oSheet.Shapes.AddPicture(Application.StartupPath + @".\LOGO_EMB.png", Microsoft.Office.Core.MsoTriState.msoFalse,
                                                    Microsoft.Office.Core.MsoTriState.msoTrue, 2, 17, 212, 125);
                            }
                        }
                    }
                }

                string FirstCNO = "";

                //載入嘜頭氣瓶序號位子
                using (conn = new SqlConnection(myConnectionString))
                {
                    conn.Open();

                    selectCmd = "SELECT  * FROM [ShippingBody] where [ListDate]='" + ListDate_LB.SelectedItem + "' and [ProductName]='" + ProductName_CB.SelectedItem + "'and [WhereBox]='" + WhereBox_LB.SelectedItem + "'";
                    cmd = new SqlCommand(selectCmd, conn);
                    using (reader = cmd.ExecuteReader())
                    {
                        while (reader.Read())
                        {
                            switch (reader.GetString(5))
                            {
                                case "1":
                                    oSheet.Cells[6, 1] = reader.GetString(3);
                                    FirstCNO = reader.GetString(3);
                                    break;
                            }
                            //serialnooneX = serialnooneX + ((Convert.ToInt32(reader.GetString(5)) + 3) % 5) * 143;
                            //serialnooneY = serialnooneY + ((Convert.ToInt32(reader.GetString(5)) - 1) / 5) * 46;
                            //oSheet.Shapes.AddPicture(serialnooneadd + reader.GetString(3) + ".png", Microsoft.Office.Core.MsoTriState.msoFalse,
                            //Microsoft.Office.Core.MsoTriState.msoTrue, serialnooneX, serialnooneY, 130, 20);
                        }
                    }

                    if ((Client.Contains("Scientific Gas Australia Pty Ltd") || Client == "Airtanks Limited") && PackingMarks.Trim().CompareTo("SGA-GLADIATAIR") == 0)
                    {
                        string ProductNO = "";
                        //該客戶要其自己的logo  PartNo   Part Description
                        selectCmd = "SELECT  Product_NO FROM MSNBody,Manufacturing where [CylinderNo]='" + FirstCNO + "' and vchManufacturingNo=  Manufacturing_NO";
                        cmd = new SqlCommand(selectCmd, conn);
                        using (reader = cmd.ExecuteReader())
                        {
                            if (reader.Read())
                            {
                                ProductNO = reader.GetValue(0).ToString();
                            }
                        }

                        selectCmd = "SELECT  ProductCode, ProductDescription FROM CustomerPackingMark where ProductNo='" + ProductNO + "' and LogoType='" + (PackingMarks.Trim().Contains("-") == true ? PackingMarks.Trim().Split('-')[1].Trim().ToUpper() : PackingMarks.Trim()) + "'";
                        cmd = new SqlCommand(selectCmd, conn);
                        using (reader = cmd.ExecuteReader())
                        {
                            if (reader.Read())
                            {
                                //載入客戶產品名稱
                                oSheet.Cells[1, 7] = reader.GetString(1);

                                //載入客戶產品型號
                                oSheet.Cells[2, 7] = reader.GetString(0);
                            }
                        }

                        //LOGO
                        oSheet.Shapes.AddPicture(Application.StartupPath + @".\LOGO_SGA_GLADIATAIR.png", Microsoft.Office.Core.MsoTriState.msoFalse,
                                            Microsoft.Office.Core.MsoTriState.msoTrue, 3, 17, 212, 125);
                    }
                    else if ((Client.ToUpper().StartsWith("HATSAN") == true) && PackingMarks.Trim().CompareTo("HATSAN") == 0)
                    {
                        //LOGO
                        oSheet.Shapes.AddPicture(Application.StartupPath + @".\LOGO_HATSAN.png", Microsoft.Office.Core.MsoTriState.msoFalse,
                                            Microsoft.Office.Core.MsoTriState.msoTrue, 3, 17, 212, 125);
                    }
                    else if (Client.ToUpper().StartsWith("EMB"))
                    {
                        oSheet.Shapes.AddPicture(Application.StartupPath + @".\LOGO_EMB.png", Microsoft.Office.Core.MsoTriState.msoFalse,
                                            Microsoft.Office.Core.MsoTriState.msoTrue, 2, 17, 212, 125);
                    }
                    else if (Client.ToUpper().StartsWith("PAINTBALL SPORTS"))
                    {
                        oSheet.Shapes.AddPicture(Application.StartupPath + @".\LOGO_Paintball Sports GmbH.png", Microsoft.Office.Core.MsoTriState.msoFalse,
                                            Microsoft.Office.Core.MsoTriState.msoTrue, 2, 17, 212, 125);
                    }
                }


                //if (StorageStatus == "N")
                //{

                //    //預設位子在X:680,Y:155
                //    //預設QRCODE圖片大小250*250

                //    //插入二維條碼
                //    int picX = 730, picY = 179;
                //    string picadd = @"C:\QRCode\";
                //    
                //    selectCmd = "SELECT  * FROM [ShippingHead] where [ListDate]='" + ListDateListBox.SelectedItem + "' and [ProductName]='" + ProductComboBox.Text + "' and [vchBoxs]='" + BoxsListBox.SelectedItem + "'";
                //    conn = new SqlConnection(myConnectionString);
                //    conn.Open();
                //    cmd = new SqlCommand(selectCmd, conn);
                //    reader = cmd.ExecuteReader();
                //    while (reader.Read())
                //    {
                //        Excel.Worksheet xSheet = (Excel.Worksheet)oWB.Sheets[1];
                //        oSheet.Shapes.AddPicture(picadd + (reader.GetString(0) + reader.GetString(1) + reader.GetString(3)) + ".png", Microsoft.Office.Core.MsoTriState.msoFalse,
                //        Microsoft.Office.Core.MsoTriState.msoTrue, picX, picY, 250, 250);
                //        if (picX == 885)
                //        {
                //            picY += 70;
                //            picX = 125;
                //        }
                //        else
                //        {
                //            picX += 190;
                //        }
                //    }
                //}
            }
            else if (Aboxof == "4" || Aboxof == "3")
            {
                string HowMuch = "";
                int Cumulative = 0;
                int Total = 0;

                //載入嘜頭資料
                using (conn = new SqlConnection(myConnectionString))
                {
                    conn.Open();

                    selectCmd = "SELECT  * FROM [ShippingHead] where [ListDate]='" + ListDate_LB.SelectedItem + "' and [ProductName]='" + ProductName_CB.SelectedItem + "'";
                    cmd = new SqlCommand(selectCmd, conn);
                    using (reader = cmd.ExecuteReader())
                    {
                        while (reader.Read())
                        {
                            HowMuch = reader.GetString(4);
                            Cumulative++;
                        }
                    }

                    Total = Convert.ToInt32(HowMuch) * Cumulative;

                    //載入嘜頭資料
                    selectCmd = "SELECT  [Client],(vchPrint + ' ' + ProductName + ' ' + Marking) PartDescription,isnull([CustomerProductName],''),isnull([CustomerProductNo],''),[vchBoxs],isnull(PalletNo,'') FROM [ShippingHead] where [ListDate]='" + ListDate_LB.SelectedItem + "' and [ProductName]='" + ProductName_CB.SelectedItem + "'and [vchBoxs]='" + WhereBox_LB.SelectedItem + "'";
                    cmd = new SqlCommand(selectCmd, conn);
                    using (reader = cmd.ExecuteReader())
                    {
                        while (reader.Read())
                        {
                            Client = reader.GetString(0).Trim();
                            //載入客戶產品名稱
                            oSheet.Cells[1, 7] = reader.GetString(reader.GetOrdinal("PartDescription"));

                            //載入客戶產品型號
                            oSheet.Cells[2, 7] = reader.GetString(3);

                            //載入一箱幾隻
                            oSheet.Cells[4, 7] = Getcount;

                            //載入箱號
                            oSheet.Cells[10, 2] = reader.GetString(4);

                            //20200410 加入PO
                            oSheet.Cells[5, 9] = CustomerPO_L.Text;

                            //載入客戶名稱
                            oSheet.Cells[3, 7] = reader.GetString(0);

                            //載入箱號
                            oSheet.Cells[10, 8] = reader.GetString(5);

                            if (reader.GetString(0).Trim().CompareTo("Wicked Sportz") == 0)
                            {
                                //該客戶要其自己的logo
                                oSheet.Shapes.AddPicture(Application.StartupPath + @".\LOGO-ENAIRGY_Wicked Sportz.jpg", Microsoft.Office.Core.MsoTriState.msoFalse,
                                                Microsoft.Office.Core.MsoTriState.msoTrue, 12, 17, 212, 125);
                                // Application.StartupPath + @".\LOGO-ENAIRGY_Wicked Sportz.jpg"
                            }
                            if (reader.GetString(0).Trim().CompareTo("達成數位") == 0)
                            {
                                //該客戶要其自己的logo
                                oSheet.Shapes.AddPicture(Application.StartupPath + @".\LOGO_DCT.jpg", Microsoft.Office.Core.MsoTriState.msoFalse,
                                                Microsoft.Office.Core.MsoTriState.msoTrue, 12, 17, 212, 125);
                                // Application.StartupPath + @".\LOGO-ENAIRGY_Wicked Sportz.jpg"
                            }
                            else if (reader.GetString(reader.GetOrdinal("Client")).Trim().Contains("ADRENALICIA S.L."))
                            {
                                oSheet.Shapes.AddPicture(Application.StartupPath + @".\LOGO_RogerSports.png", Microsoft.Office.Core.MsoTriState.msoFalse,
                            Microsoft.Office.Core.MsoTriState.msoTrue, 2, 17, 212, 125);
                            }
                            else if (Client.ToUpper().StartsWith("EMB"))
                            {
                                oSheet.Shapes.AddPicture(Application.StartupPath + @".\LOGO_EMB.png", Microsoft.Office.Core.MsoTriState.msoFalse,
                                                    Microsoft.Office.Core.MsoTriState.msoTrue, 2, 17, 212, 125);
                            }
                        }
                    }
                }

                int serialnooneX = 10, serialnooneY = 239;
                string serialnooneadd = @"C:\SerialNoCode\";

                string FirstCNO = "";

                //載入嘜頭氣瓶序號位子
                using (conn = new SqlConnection(myConnectionString))
                {
                    conn.Open();

                    selectCmd = "SELECT  * FROM [ShippingBody] where [ListDate]='" + ListDate_LB.SelectedItem + "' and [ProductName]='" + ProductName_CB.SelectedItem + "'and [WhereBox]='" + WhereBox_LB.SelectedItem + "'";
                    cmd = new SqlCommand(selectCmd, conn);
                    using (reader = cmd.ExecuteReader())
                    {
                        while (reader.Read())
                        {
                            serialnooneX = 3; serialnooneY = 209;
                            switch (reader.GetString(5))
                            {
                                case "1":
                                    oSheet.Cells[6, 1] = reader.GetString(3);
                                    FirstCNO = reader.GetString(3);
                                    break;

                                case "2":
                                    oSheet.Cells[6, 5] = reader.GetString(3);
                                    break;

                                case "3":
                                    oSheet.Cells[8, 1] = reader.GetString(3);
                                    break;

                                case "4":
                                    oSheet.Cells[8, 5] = reader.GetString(3);
                                    break;

                            }
                            serialnooneX = serialnooneX + ((Convert.ToInt32(reader.GetString(5)) + 3) % 2) * 315;
                            serialnooneY = serialnooneY + ((Convert.ToInt32(reader.GetString(5)) - 1) / 2) * 111;
                            oSheet.Shapes.AddPicture(serialnooneadd + reader.GetString(3) + ".png", Microsoft.Office.Core.MsoTriState.msoFalse,
                            Microsoft.Office.Core.MsoTriState.msoTrue, serialnooneX, serialnooneY, 90, 90);//, 150, 30);
                        }
                    }

                    if ((Client.Contains("Scientific Gas Australia Pty Ltd") || Client == "Airtanks Limited") && PackingMarks.Trim().CompareTo("SGA-GLADIATAIR") == 0)
                    {
                        string ProductNO = "";

                        //該客戶要其自己的logo  PartNo   Part Description
                        selectCmd = "SELECT  Product_NO FROM MSNBody,Manufacturing where [CylinderNo]='" + FirstCNO + "' and vchManufacturingNo=  Manufacturing_NO";
                        cmd = new SqlCommand(selectCmd, conn);
                        using (reader = cmd.ExecuteReader())
                        {
                            if (reader.Read())
                            {
                                ProductNO = reader.GetValue(0).ToString();
                            }
                        }

                        selectCmd = "SELECT  ProductCode, ProductDescription FROM CustomerPackingMark where ProductNo='" + ProductNO + "' and LogoType='" + (PackingMarks.Trim().Contains("-") == true ? PackingMarks.Trim().Split('-')[1].Trim().ToUpper() : PackingMarks.Trim()) + "'";
                        cmd = new SqlCommand(selectCmd, conn);
                        using (reader = cmd.ExecuteReader())
                        {
                            if (reader.Read())
                            {
                                //載入客戶產品名稱
                                oSheet.Cells[1, 7] = reader.GetString(1);

                                //載入客戶產品型號
                                oSheet.Cells[2, 7] = reader.GetString(0);
                            }
                        }

                        //LOGO
                        oSheet.Shapes.AddPicture(Application.StartupPath + @".\LOGO_SGA_GLADIATAIR.png", Microsoft.Office.Core.MsoTriState.msoFalse,
                                            Microsoft.Office.Core.MsoTriState.msoTrue, 3, 17, 212, 125);
                    }
                    else if ((Client.ToUpper().StartsWith("HATSAN") == true) && PackingMarks.Trim().CompareTo("HATSAN") == 0)
                    {
                        //LOGO
                        oSheet.Shapes.AddPicture(Application.StartupPath + @".\LOGO_HATSAN.png", Microsoft.Office.Core.MsoTriState.msoFalse,
                                            Microsoft.Office.Core.MsoTriState.msoTrue, 3, 17, 212, 125);
                    }
                    else if (Client.ToUpper().StartsWith("EMB"))
                    {
                        oSheet.Shapes.AddPicture(Application.StartupPath + @".\LOGO_EMB.png", Microsoft.Office.Core.MsoTriState.msoFalse,
                                            Microsoft.Office.Core.MsoTriState.msoTrue, 2, 17, 212, 125);
                    }

                    //if (StorageStatus == "N")//20190212
                    {

                        //預設位子在X:680,Y:155
                        //預設QRCODE圖片大小250*250

                        //插入圖片
                        int picX = 680, picY = 185;
                        string picadd = @"C:\QRCode\";

                        selectCmd = "SELECT  * FROM [ShippingHead] where [ListDate]='" + ListDate_LB.SelectedItem + "' and [ProductName]='" + ProductName_CB.Text + "' and [vchBoxs]='" + WhereBox_LB.SelectedItem + "'";
                        cmd = new SqlCommand(selectCmd, conn);
                        using (reader = cmd.ExecuteReader())
                        {
                            while (reader.Read())
                            {
                                Excel.Worksheet xSheet = (Excel.Worksheet)oWB.Sheets[1];
                                oSheet.Shapes.AddPicture(picadd + (reader.GetString(0) + reader.GetString(1) + reader.GetString(3)) + ".png", Microsoft.Office.Core.MsoTriState.msoFalse,
                                Microsoft.Office.Core.MsoTriState.msoTrue, picX, picY, 250, 250);
                                if (picX == 885)
                                {
                                    picY += 70;
                                    picX = 125;
                                }
                                else
                                {
                                    picX += 190;
                                }
                            }
                        }
                    }
                }
            }
            else if (Aboxof == "2")
            {
                string HowMuch = "";
                int Cumulative = 0;
                int Total = 0;

                //載入嘜頭資料
                using (conn = new SqlConnection(myConnectionString))
                {
                    conn.Open();

                    selectCmd = "SELECT  * FROM [ShippingHead] where [ListDate]='" + ListDate_LB.SelectedItem + "' and [ProductName]='" + ProductName_CB.SelectedItem + "'";
                    cmd = new SqlCommand(selectCmd, conn);
                    using (reader = cmd.ExecuteReader())
                    {
                        while (reader.Read())
                        {
                            HowMuch = reader.GetString(4);
                            Cumulative++;
                        }
                    }

                    Total = Convert.ToInt32(HowMuch) * Cumulative;

                    //載入嘜頭資料
                    selectCmd = "SELECT  [Client],(vchPrint + ' ' + ProductName + ' ' + Marking) PartDescription,isnull([CustomerProductName],''),isnull([CustomerProductNo],''),[vchBoxs],isnull(PalletNo,'') FROM [ShippingHead] where [ListDate]='" + ListDate_LB.SelectedItem + "' and [ProductName]='" + ProductName_CB.SelectedItem + "'and [vchBoxs]='" + WhereBox_LB.SelectedItem + "'";
                    cmd = new SqlCommand(selectCmd, conn);
                    using (reader = cmd.ExecuteReader())
                    {
                        while (reader.Read())
                        {
                            Client = reader.GetString(0).Trim();
                            //載入客戶產品名稱
                            oSheet.Cells[1, 7] = reader.GetString(reader.GetOrdinal("PartDescription"));

                            //載入客戶產品型號
                            oSheet.Cells[2, 7] = reader.GetString(3);

                            //載入一箱幾隻
                            oSheet.Cells[4, 7] = Getcount;

                            //載入箱號
                            oSheet.Cells[10, 2] = reader.GetString(4);

                            //20200410 加入PO
                            oSheet.Cells[5, 9] = CustomerPO_L.Text;

                            //載入客戶名稱
                            oSheet.Cells[3, 7] = reader.GetString(0);

                            //載入箱號
                            oSheet.Cells[10, 8] = reader.GetString(5);

                            if (reader.GetString(0).Trim().CompareTo("Wicked Sportz") == 0)
                            {
                                //該客戶要其自己的logo
                                oSheet.Shapes.AddPicture(Application.StartupPath + @".\LOGO-ENAIRGY_Wicked Sportz.jpg", Microsoft.Office.Core.MsoTriState.msoFalse,
                                                Microsoft.Office.Core.MsoTriState.msoTrue, 12, 17, 212, 125);
                                // Application.StartupPath + @".\LOGO-ENAIRGY_Wicked Sportz.jpg"
                            }
                            if (reader.GetString(0).Trim().CompareTo("達成數位") == 0)
                            {
                                //該客戶要其自己的logo
                                oSheet.Shapes.AddPicture(Application.StartupPath + @".\LOGO_DCT.jpg", Microsoft.Office.Core.MsoTriState.msoFalse,
                                                Microsoft.Office.Core.MsoTriState.msoTrue, 12, 17, 212, 125);
                                // Application.StartupPath + @".\LOGO-ENAIRGY_Wicked Sportz.jpg"
                            }
                            else if (reader.GetString(reader.GetOrdinal("Client")).Trim().Contains("ADRENALICIA S.L."))
                            {
                                oSheet.Shapes.AddPicture(Application.StartupPath + @".\LOGO_RogerSports.png", Microsoft.Office.Core.MsoTriState.msoFalse,
                            Microsoft.Office.Core.MsoTriState.msoTrue, 2, 17, 212, 125);
                            }
                            else if (Client.ToUpper().StartsWith("EMB"))
                            {
                                oSheet.Shapes.AddPicture(Application.StartupPath + @".\LOGO_EMB.png", Microsoft.Office.Core.MsoTriState.msoFalse,
                                                    Microsoft.Office.Core.MsoTriState.msoTrue, 2, 17, 212, 125);
                            }
                        }
                    }
                }

                int serialnooneX = 10, serialnooneY = 239;
                string serialnooneadd = @"C:\SerialNoCode\";

                string FirstCNO = "";

                //載入嘜頭氣瓶序號位子
                using (conn = new SqlConnection(myConnectionString))
                {
                    conn.Open();

                    selectCmd = "SELECT  * FROM [ShippingBody] where [ListDate]='" + ListDate_LB.SelectedItem + "' and [ProductName]='" + ProductName_CB.SelectedItem + "'and [WhereBox]='" + WhereBox_LB.SelectedItem + "'";
                    cmd = new SqlCommand(selectCmd, conn);
                    using (reader = cmd.ExecuteReader())
                    {
                        while (reader.Read())
                        {
                            serialnooneX = 3; serialnooneY = 270;
                            switch (reader.GetString(5))
                            {
                                case "1":
                                    oSheet.Cells[6, 1] = reader.GetString(3);
                                    FirstCNO = reader.GetString(3);
                                    break;
                                case "2":
                                    oSheet.Cells[6, 5] = reader.GetString(3);
                                    break;

                            }
                            serialnooneX = serialnooneX + ((Convert.ToInt32(reader.GetString(5)) + 3) % 2) * 315;
                            serialnooneY = serialnooneY;// +((Convert.ToInt32(reader.GetString(5))) / 2) * 1111;
                            oSheet.Shapes.AddPicture(serialnooneadd + reader.GetString(3) + ".png", Microsoft.Office.Core.MsoTriState.msoFalse,
                            Microsoft.Office.Core.MsoTriState.msoTrue, serialnooneX, serialnooneY, 90, 90);//, 150, 30);
                        }
                    }

                    if ((Client.Contains("Scientific Gas Australia Pty Ltd") || Client == "Airtanks") && PackingMarks.Trim().CompareTo("SGA-GLADIATAIR") == 0)
                    {
                        string ProductNO = "";

                        //該客戶要其自己的logo  PartNo   Part Description
                        selectCmd = "SELECT  Product_NO FROM MSNBody,Manufacturing where [CylinderNo]='" + FirstCNO + "' and vchManufacturingNo=  Manufacturing_NO";
                        cmd = new SqlCommand(selectCmd, conn);
                        using (reader = cmd.ExecuteReader())
                        {
                            if (reader.Read())
                            {
                                ProductNO = reader.GetValue(0).ToString();
                            }
                        }

                        selectCmd = "SELECT  ProductCode, ProductDescription FROM CustomerPackingMark where ProductNo='" + ProductNO + "' and LogoType='" + (PackingMarks.Trim().Contains("-") == true ? PackingMarks.Trim().Split('-')[1].Trim().ToUpper() : PackingMarks.Trim()) + "'";
                        cmd = new SqlCommand(selectCmd, conn);
                        using (reader = cmd.ExecuteReader())
                        {
                            if (reader.Read())
                            {
                                //載入客戶產品名稱
                                oSheet.Cells[1, 7] = reader.GetString(1);

                                //載入客戶產品型號
                                oSheet.Cells[2, 7] = reader.GetString(0);
                            }
                        }

                        //LOGO
                        oSheet.Shapes.AddPicture(Application.StartupPath + @".\LOGO_SGA_GLADIATAIR.png", Microsoft.Office.Core.MsoTriState.msoFalse,
                                            Microsoft.Office.Core.MsoTriState.msoTrue, 12, 17, 212, 125);
                    }
                    else if (Client.ToUpper().StartsWith("EMB"))
                    {
                        oSheet.Shapes.AddPicture(Application.StartupPath + @".\LOGO_EMB.png", Microsoft.Office.Core.MsoTriState.msoFalse,
                                            Microsoft.Office.Core.MsoTriState.msoTrue, 2, 17, 212, 125);
                    }
                    else if ((Client.ToUpper().StartsWith("HATSAN") == true) && PackingMarks.Trim().CompareTo("HATSAN") == 0)
                    {
                        //LOGO
                        oSheet.Shapes.AddPicture(Application.StartupPath + @".\LOGO_HATSAN.png", Microsoft.Office.Core.MsoTriState.msoFalse,
                                            Microsoft.Office.Core.MsoTriState.msoTrue, 12, 17, 212, 125);
                    }

                    else if (Client.ToUpper().StartsWith("PAINTBALL SPORTS"))
                    {
                        oSheet.Shapes.AddPicture(Application.StartupPath + @".\LOGO_Paintball Sports GmbH.png", Microsoft.Office.Core.MsoTriState.msoFalse,
                                            Microsoft.Office.Core.MsoTriState.msoTrue, 2, 17, 212, 125);
                    }

                    //if (StorageStatus == "N")//20190212
                    {
                        //預設位子在X:680,Y:155
                        //預設QRCODE圖片大小250*250

                        //插入圖片
                        int picX = 680, picY = 183;
                        string picadd = @"C:\QRCode\";

                        selectCmd = "SELECT  * FROM [ShippingHead] where [ListDate]='" + ListDate_LB.SelectedItem + "' and [ProductName]='" + ProductName_CB.Text + "' and [vchBoxs]='" + WhereBox_LB.SelectedItem + "'";
                        cmd = new SqlCommand(selectCmd, conn);
                        using (reader = cmd.ExecuteReader())
                        {
                            while (reader.Read())
                            {
                                Excel.Worksheet xSheet = (Excel.Worksheet)oWB.Sheets[1];
                                oSheet.Shapes.AddPicture(picadd + (reader.GetString(0) + reader.GetString(1) + reader.GetString(3)) + ".png", Microsoft.Office.Core.MsoTriState.msoFalse,
                                Microsoft.Office.Core.MsoTriState.msoTrue, picX, picY, 250, 250);
                                if (picX == 885)
                                {
                                    picY += 70;
                                    picX = 125;
                                }
                                else
                                {
                                    picX += 190;
                                }
                            }
                        }
                    }
                }
            }

            Excel.Sheets excelSheets = oWB.Worksheets;

            //顯示EXCEL
            oXL.Visible = true;

            if (AutoPrintCheckBox.Checked == true)
            {
                //列印EXCEL
                oWB.PrintOutEx(Missing.Value, Missing.Value, Missing.Value, Missing.Value, Missing.Value, Missing.Value, Missing.Value, Missing.Value, Missing.Value);
            }
            oXL.DisplayAlerts = false;

            if (AutoPrintCheckBox.Checked == true)
            {
                //關閉EXCEL
                oWB.Close(Type.Missing, Type.Missing, Type.Missing);
            }
            //釋放EXCEL資源
            oXL = null;
            oWB = null;
            oSheet = null;
        }

        private void Customer_SGA_Form(string Aboxof, string PackingMarks)
        {
            if (Aboxof != "1")
            {
                MessageBox.Show("客製化需求未定義裝箱數為" + Aboxof + "之嘜頭表格");
                return;
            }
            //客戶SGA需求嘜頭表格
            //公司定義的嘜頭表格
            Excel.Application oXL = new Excel.Application();
            Excel.Workbook oWB;
            Excel.Worksheet oSheet;

            string srcFileName = "";

            if (Aboxof == "1")
            {
                srcFileName = Application.StartupPath + @".\SGAForm_1.xlsx";//EXCEL檔案路徑
            }

            try
            {
                //產生一個Workbook物件，並加入Application//改成.open以及在()中輸入開啟位子
                oWB = oXL.Workbooks.Open(srcFileName, Type.Missing, Type.Missing, Type.Missing, Type.Missing,
                                        Type.Missing, Type.Missing, Type.Missing, Type.Missing,
                                        Type.Missing, Type.Missing, Type.Missing, Type.Missing,
                                        Type.Missing, Type.Missing);
            }
            catch
            {
                MessageBox.Show(@"找不到EXCEL檔案！", "Warning");
                return;
            }

            GetThisBoxMaxCount();

            //設定工作表
            oSheet = (Excel.Worksheet)oWB.ActiveSheet;

            if (Aboxof == "1")
            {
                //插入1維條碼
                //預設位子在X:255,Y:412
                //預設1維條碼圖片大小170*35            
                int oneX = 255, oneY = 411;
                string oneadd = @"C:\Code\";
                int serialnooneX = 308, serialnooneY = 128;
                string serialnooneadd = @"C:\SerialNoCode\";
                string CylinderNo = "", HydrostaticTestDate = "", ProductNO = "";

                oSheet.Shapes.AddPicture(oneadd + WhereBox_LB.SelectedItem + ".png", Microsoft.Office.Core.MsoTriState.msoFalse,
                                    Microsoft.Office.Core.MsoTriState.msoTrue, oneX, oneY, 170, 35);
                //載入嘜頭資料
                using (conn = new SqlConnection(myConnectionString))
                {
                    conn.Open();

                    selectCmd = "SELECT  isnull(CustomerPO,'') CustomerPO, isnull(CustomerProductName,'') CustomerProductName, isnull(CustomerProductNo,'') CustomerProductNo, vchBoxs FROM [ShippingHead] " +
                        "where [ListDate]='" + ListDate_LB.SelectedItem + "' and [ProductName]='" + ProductName_CB.SelectedItem + "'and [vchBoxs]='" + WhereBox_LB.SelectedItem + "'";
                    cmd = new SqlCommand(selectCmd, conn);
                    using (reader = cmd.ExecuteReader())
                    {
                        if (reader.Read())
                        {
                            //描述
                            oSheet.Cells[5, 3] = reader.GetString(reader.GetOrdinal("CustomerProductName"));

                            //品號
                            oSheet.Cells[6, 3] = reader.GetString(reader.GetOrdinal("CustomerProductNo"));

                            //載入P/O No.
                            oSheet.Cells[8, 3] = reader.GetString(reader.GetOrdinal("CustomerPO"));

                            //載入一箱幾隻
                            oSheet.Cells[7, 3] = Getcount;

                            //載入箱號
                            oSheet.Cells[9, 3] = reader.GetString(reader.GetOrdinal("vchBoxs"));
                        }
                    }

                    //載入嘜頭氣瓶序號位子v
                    selectCmd = "SELECT WhereSeat, CylinderNumbers FROM [ShippingBody] where [ListDate]='" + ListDate_LB.SelectedItem + "' and [ProductName]='" + ProductName_CB.SelectedItem + "'and [WhereBox]='" + WhereBox_LB.SelectedItem + "'";
                    cmd = new SqlCommand(selectCmd, conn);
                    using (reader = cmd.ExecuteReader())
                    {
                        while (reader.Read())
                        {
                            switch (reader.GetString(reader.GetOrdinal("WhereSeat")))
                            {
                                case "1":
                                    oSheet.Cells[2, 4] = reader.GetString(reader.GetOrdinal("CylinderNumbers"));
                                    CylinderNo = reader.GetString(reader.GetOrdinal("CylinderNumbers"));
                                    MarkSerialNoBarCode(CylinderNo);
                                    oSheet.Shapes.AddPicture(serialnooneadd + reader.GetString(reader.GetOrdinal("CylinderNumbers")) + ".png", Microsoft.Office.Core.MsoTriState.msoFalse,
                                    Microsoft.Office.Core.MsoTriState.msoTrue, serialnooneX, serialnooneY, 255, 44);
                                    break;
                            }
                        }
                    }

                    //載入由序號找水壓年月
                    selectCmd = "SELECT  Product_NO, vchHydrostaticTestDate FROM MSNBody,Manufacturing " +
                        "where [CylinderNo]='" + CylinderNo + "' and vchManufacturingNo=  Manufacturing_NO";
                    cmd = new SqlCommand(selectCmd, conn);
                    using (reader = cmd.ExecuteReader())
                    {
                        if (reader.Read())
                        {
                            HydrostaticTestDate = reader.GetValue(1).ToString();
                            ProductNO = reader.GetValue(0).ToString();
                        }
                    }
                }

                if (HydrostaticTestDate.Contains("/") == true)
                {
                    oSheet.Cells[9, 7] = HydrostaticTestDate.Split('/')[1] + HydrostaticTestDate.Split('/')[0].Substring(2, 2);
                }
                else
                {
                    oSheet.Cells[9, 7] = HydrostaticTestDate;
                }

                //由序號找出產品型號再找出Part Description、Part No.
                using (conn = new SqlConnection(myConnectionString))
                {
                    conn.Open();

                    selectCmd = "SELECT ProductCode, ProductDescription FROM CustomerPackingMark " +
                        "where ProductNo='" + ProductNO + "' and LogoType='" + (PackingMarks.Trim().Contains("-") == true ? PackingMarks.Trim().Split('-')[1].Trim().ToUpper() : PackingMarks.Trim()) + "'";
                    cmd = new SqlCommand(selectCmd, conn);
                    using (reader = cmd.ExecuteReader())
                    {
                        if (reader.Read())
                        {
                            //載入Part Description
                            oSheet.Cells[5, 3] = reader.GetString(reader.GetOrdinal("ProductDescription")).Contains("-") == true ? reader.GetString(reader.GetOrdinal("ProductDescription")).Replace("- ", "\n") : reader.GetString(reader.GetOrdinal("ProductDescription"));

                            if (reader.GetString(reader.GetOrdinal("ProductDescription")).Contains("-") == true)
                            {
                                oSheet.get_Range("C5").Font.Size = 22;
                            }
                            //oSheet.get_Range("C5").ShrinkToFit = true;// '設定為縮小字型以適合欄寬
                            //載入Part No.
                            oSheet.Cells[6, 3] = reader.GetString(reader.GetOrdinal("ProductCode"));
                        }
                    }

                    //預設位子在X:446,Y:228
                    //預設QRCODE圖片大小190*190

                    //插入二維條碼
                    int picX = 444, picY = 228;
                    string picadd = @"C:\QRCode\";

                    selectCmd = "SELECT ListDate, ProductName, vchBoxs FROM [ShippingHead] " +
                        "where [ListDate]='" + ListDate_LB.SelectedItem + "' and [ProductName]='" + ProductName_CB.Text + "' and [vchBoxs]='" + WhereBox_LB.SelectedItem + "'";
                    cmd = new SqlCommand(selectCmd, conn);
                    using (reader = cmd.ExecuteReader())
                    {
                        if (reader.Read())
                        {
                            oSheet.Shapes.AddPicture(picadd + reader.GetString(reader.GetOrdinal("ListDate")) + reader.GetString(reader.GetOrdinal("ProductName")) + reader.GetString(reader.GetOrdinal("vchBoxs")) + ".png", Microsoft.Office.Core.MsoTriState.msoFalse,
                            Microsoft.Office.Core.MsoTriState.msoTrue, picX, picY, 190, 190);
                        }
                    }
                }

                if (PackingMarks.Trim().CompareTo("SGA-SHOOTAIR") == 0)
                {
                    //該客戶要其自己的logo
                    oSheet.Shapes.AddPicture(Application.StartupPath + @".\LOGO_SGA_Shootair.jpg", Microsoft.Office.Core.MsoTriState.msoFalse,
                                    Microsoft.Office.Core.MsoTriState.msoTrue, 24, 5, 194, (float)168.9);
                }
                else if (PackingMarks.Trim().CompareTo("SGA-BREATHEAIR") == 0)
                {
                    //該客戶要其自己的logo
                    oSheet.Shapes.AddPicture(Application.StartupPath + @".\LOGO_SGA_Breatheair.jpg", Microsoft.Office.Core.MsoTriState.msoFalse,
                                    Microsoft.Office.Core.MsoTriState.msoTrue, 17, 5, (float)205.4, (float)167.9);
                }
                else if (PackingMarks.Trim().CompareTo("SGA-SCUBAIR") == 0)
                {
                    //該客戶要其自己的logo
                    oSheet.Shapes.AddPicture(Application.StartupPath + @".\LOGO_SGA_SCUBAIR.jpg", Microsoft.Office.Core.MsoTriState.msoFalse,
                                    Microsoft.Office.Core.MsoTriState.msoTrue, 26, 3, (float)180.5, (float)172.3);
                }
                else if (PackingMarks.Trim().CompareTo("SGA-SPIROTEK") == 0)
                {
                    //該客戶要其自己的logo
                    oSheet.Shapes.AddPicture(Application.StartupPath + @".\LOGO_SGA_SPIROTEK.jpg", Microsoft.Office.Core.MsoTriState.msoFalse,
                                    Microsoft.Office.Core.MsoTriState.msoTrue, 8, 20, 219, (float)133.9);
                }
                else if (PackingMarks.Trim().CompareTo("SGA-SGA") == 0)
                {
                    //該客戶要其自己的logo
                    oSheet.Shapes.AddPicture(Application.StartupPath + @".\LOGO_SGA_SGA.jpg", Microsoft.Office.Core.MsoTriState.msoFalse,
                                    Microsoft.Office.Core.MsoTriState.msoTrue, 8, 20, 219, (float)133.9);
                }
                else if (PackingMarks.Trim().CompareTo("SGA-GLADIATAIR") == 0)
                {
                    oSheet.Shapes.AddPicture(Application.StartupPath + @".\LOGO_SGA_GLADIATAIR.png", Microsoft.Office.Core.MsoTriState.msoFalse,
                                        Microsoft.Office.Core.MsoTriState.msoTrue, 8, 20, 219, (float)133.9);
                }
            }

            Excel.Sheets excelSheets = oWB.Worksheets;
            //顯示EXCEL
            oXL.Visible = true;

            if (AutoPrintCheckBox.Checked == true)
            {
                //列印EXCEL
                oWB.PrintOutEx(Missing.Value, Missing.Value, Missing.Value, Missing.Value, Missing.Value, Missing.Value, Missing.Value, Missing.Value, Missing.Value);
            }
            oXL.DisplayAlerts = false;

            if (AutoPrintCheckBox.Checked == true)
            {
                //關閉EXCEL
                oWB.Close(Type.Missing, Type.Missing, Type.Missing);
            }
            //釋放EXCEL資源
            oXL = null;
            oWB = null;
            oSheet = null;
        }

        private void Customer_Estratego_Form(string Aboxof, string PackingMarks)
        {
            Excel.Application oXL = new Excel.Application();
            Excel.Workbook oWB;
            Excel.Worksheet oSheet;

            string srcFileName = "";
            srcFileName = Application.StartupPath + @".\EstrategoForm_1.xlsx";//EXCEL檔案路徑

            try
            {
                //產生一個Workbook物件，並加入Application//改成.open以及在()中輸入開啟位子
                oWB = oXL.Workbooks.Open(srcFileName, Type.Missing, Type.Missing, Type.Missing, Type.Missing,
                                        Type.Missing, Type.Missing, Type.Missing, Type.Missing,
                                        Type.Missing, Type.Missing, Type.Missing, Type.Missing,
                                        Type.Missing, Type.Missing);
            }
            catch
            {
                MessageBox.Show(@"找不到EXCEL檔案！", "Warning");
                return;
            }

            //設定工作表
            oSheet = (Excel.Worksheet)oWB.ActiveSheet;

            //7*3.5
            if (PackingMarks.Trim().CompareTo("Regulator 3000psi") == 0)
            {
                //該客戶要其自己的logo
                oSheet.Shapes.AddPicture(Application.StartupPath + @".\LOGO_Regulator 3000psi.png", Microsoft.Office.Core.MsoTriState.msoFalse,
                                Microsoft.Office.Core.MsoTriState.msoTrue, 0, 0, 545, 450);
            }
            else if (PackingMarks.Trim().CompareTo("Estratego-48ci 3000psi+Regulator") == 0)
            {
                //該客戶要其自己的logo
                oSheet.Shapes.AddPicture(Application.StartupPath + @".\LOGO_Estratego_48ci 3000psi_Regulator.png", Microsoft.Office.Core.MsoTriState.msoFalse,
                                Microsoft.Office.Core.MsoTriState.msoTrue, 0, 0, 545, 450);
            }
            else if (PackingMarks.Trim().CompareTo("Estratego-48ci 3000psi") == 0)
            {
                //該客戶要其自己的logo
                oSheet.Shapes.AddPicture(Application.StartupPath + @".\LOGO_Estratego_48ci 3000psi.png", Microsoft.Office.Core.MsoTriState.msoFalse,
                                Microsoft.Office.Core.MsoTriState.msoTrue, 0, 0, 545, 450);
            }
            else if (PackingMarks.Trim().CompareTo("Estratego-38ci 3000psi") == 0)
            {
                //該客戶要其自己的logo
                oSheet.Shapes.AddPicture(Application.StartupPath + @".\LOGO_Estratego_38ci 3000psi.png", Microsoft.Office.Core.MsoTriState.msoFalse,
                                Microsoft.Office.Core.MsoTriState.msoTrue, 0, 0, 545, 450);
            }
            else if (PackingMarks.Trim().CompareTo("Estratego-13ci 3000psi+Regulator") == 0)
            {
                //該客戶要其自己的logo
                oSheet.Shapes.AddPicture(Application.StartupPath + @".\LOGO_Estratego_13ci 3000psi_Regulator.png", Microsoft.Office.Core.MsoTriState.msoFalse,
                                Microsoft.Office.Core.MsoTriState.msoTrue, 0, 0, 545, 450);
            }
            else if (PackingMarks.Trim().CompareTo("Estratego-13ci 3000psi") == 0)
            {
                //該客戶要其自己的logo
                oSheet.Shapes.AddPicture(Application.StartupPath + @".\LOGO_Estratego_13ci 3000psi.png", Microsoft.Office.Core.MsoTriState.msoFalse,
                                Microsoft.Office.Core.MsoTriState.msoTrue, 0, 0, 545, 450);
            }
            else if (PackingMarks.Trim().CompareTo("Estratego-12oz") == 0)
            {
                //該客戶要其自己的logo
                oSheet.Shapes.AddPicture(Application.StartupPath + @".\LOGO_Estratego_12oz.png", Microsoft.Office.Core.MsoTriState.msoFalse,
                                Microsoft.Office.Core.MsoTriState.msoTrue, 0, 0, 545, 450);
            }
            else if (PackingMarks.Trim().CompareTo("Estratego-20oz") == 0)
            {
                //該客戶要其自己的logo
                oSheet.Shapes.AddPicture(Application.StartupPath + @".\LOGO_Estratego_20oz.png", Microsoft.Office.Core.MsoTriState.msoFalse,
                                Microsoft.Office.Core.MsoTriState.msoTrue, 0, 0, 545, 450);
            }
            else if (PackingMarks.Trim().CompareTo("Estratego-68ci(Assault) 4500psi") == 0)
            {
                //該客戶要其自己的logo
                oSheet.Shapes.AddPicture(Application.StartupPath + @".\LOGO_Estratego_68ci(Assault) 4500psi.png", Microsoft.Office.Core.MsoTriState.msoFalse,
                                Microsoft.Office.Core.MsoTriState.msoTrue, 0, 0, 545, 450);
            }
            else if (PackingMarks.Trim().CompareTo("Estratego-68ci(Snow White) 4500psi") == 0)
            {
                //該客戶要其自己的logo
                oSheet.Shapes.AddPicture(Application.StartupPath + @".\LOGO_Estratego_68ci(Snow White) 4500psi.png", Microsoft.Office.Core.MsoTriState.msoFalse,
                                Microsoft.Office.Core.MsoTriState.msoTrue, 0, 0, 545, 450);
            }
            else if (PackingMarks.Trim().CompareTo("Estratego_68ci_UL") == 0)
            {
                //該客戶要其自己的logo
                oSheet.Shapes.AddPicture(Application.StartupPath + @".\LOGO_Estratego_68ci_UL.png", Microsoft.Office.Core.MsoTriState.msoFalse,
                                Microsoft.Office.Core.MsoTriState.msoTrue, 0, 0, 545, 450);
            }
            else if (PackingMarks.Trim().CompareTo("Estratego-26ci 3000psi_Regulator") == 0)
            {
                //該客戶要其自己的logo
                oSheet.Shapes.AddPicture(Application.StartupPath + @".\LOGO_Estratego_26ci_Regulator.png", Microsoft.Office.Core.MsoTriState.msoFalse,
                                Microsoft.Office.Core.MsoTriState.msoTrue, 0, 0, 545, 450);
            }
            else if (PackingMarks.Trim().CompareTo("Estratego-BL-xx-10_1.8K") == 0)
            {
                //該客戶要其自己的logo
                oSheet.Shapes.AddPicture(Application.StartupPath + @".\LOGO_Estratego_BL-xx-10.png", Microsoft.Office.Core.MsoTriState.msoFalse,
                                Microsoft.Office.Core.MsoTriState.msoTrue, 0, 0, 545, 450);
            }
            else if (PackingMarks.Trim().CompareTo("Estratego-BL-xx-13_5K") == 0)
            {
                //該客戶要其自己的logo
                oSheet.Shapes.AddPicture(Application.StartupPath + @".\LOGO_Estratego_BL-xx-13.png", Microsoft.Office.Core.MsoTriState.msoFalse,
                                Microsoft.Office.Core.MsoTriState.msoTrue, 0, 0, 545, 450);
            }
            else if (PackingMarks.Trim().CompareTo("Estratego-Nipple") == 0)
            {
                //該客戶要其自己的logo
                oSheet.Shapes.AddPicture(Application.StartupPath + @".\LOGO_Estratego_Nipple.png", Microsoft.Office.Core.MsoTriState.msoFalse,
                                Microsoft.Office.Core.MsoTriState.msoTrue, 0, 0, 545, 450);
            }
            else if (PackingMarks.Trim().CompareTo("Estratego-Regulator Gauge") == 0)
            {
                //該客戶要其自己的logo
                oSheet.Shapes.AddPicture(Application.StartupPath + @".\LOGO_Estratego_Regulator Gauge.png", Microsoft.Office.Core.MsoTriState.msoFalse,
                                Microsoft.Office.Core.MsoTriState.msoTrue, 0, 0, 545, 450);
            }
            else if (PackingMarks.Trim().CompareTo("Estratego-Totem_Air_UL_Tank") == 0)
            {
                //該客戶要其自己的logo
                oSheet.Shapes.AddPicture(Application.StartupPath + @".\LOGO_Estratego_Totem_Air_UL_Tank.png", Microsoft.Office.Core.MsoTriState.msoFalse,
                                Microsoft.Office.Core.MsoTriState.msoTrue, 0, 0, 545, 450);
            }

            Excel.Sheets excelSheets = oWB.Worksheets;
            //顯示EXCEL
            oXL.Visible = true;
            if (AutoPrintCheckBox.Checked == true)
            {
                //列印EXCEL
                oWB.PrintOutEx(Missing.Value, Missing.Value, Missing.Value, Missing.Value, Missing.Value, Missing.Value, Missing.Value, Missing.Value, Missing.Value);
            }
            oXL.DisplayAlerts = false;

            if (AutoPrintCheckBox.Checked == true)
            {
                //關閉EXCEL
                oWB.Close(Type.Missing, Type.Missing, Type.Missing);
            }
            //釋放EXCEL資源
            oXL = null;
            oWB = null;
            oSheet = null;
        }

        private void RefreshhButton_Click(object sender, EventArgs e)
        {
            //重新刷新LISTBOX
            LoadListDate();
        }

        private void button3_Click(object sender, EventArgs e)
        {
            //切換讀取位子
            using (conn = new SqlConnection(myConnectionString))
            {
                conn.Open();

                selectCmd = "Update [LaserMarkDirection] SET  [vchWhere] = 0";
                cmd = new SqlCommand(selectCmd, conn);
                cmd.ExecuteNonQuery();
            }
        }

        private void button4_Click(object sender, EventArgs e)
        {
            //更新氣瓶相關資料進入MSNBody資料表
            using (conn = new SqlConnection(myConnectionString))
            {
                conn.Open();

                selectCmd = "Update [LaserMarkDirection] SET  [vchWhere] = 1";
                cmd = new SqlCommand(selectCmd, conn);
                cmd.ExecuteNonQuery();
            }
        }

        private void timer1_Tick(object sender, EventArgs e)
        {
            if (BottleTextBox.Text != "" && BottomTextBox.Text != "" && Message == false)
            {
                DateTime ResrictionDate = new DateTime();
                DateTime HydroDate = new DateTime();

                //bool HydrostaticPass = false;
                bool ProductAcceptance = false;

                string SpecialUses = "N";
                string NowSeat = "";
                string HydrostaticTestDate = "";
                string CustomerName = "";
                string LotNumber = string.Empty;
                string MarkingType = string.Empty;
                string CylinderNumbers = string.Empty;
                string Error = string.Empty;
                string ProductNo = string.Empty;
                string ProductType = string.Empty;

                string Bottle = string.Empty;
                string Bottom = string.Empty;

                bool HydroLabelPass = false;

                if (BottleTextBox.Text != "")
                {
                    Bottle = BottleTextBox.Text;
                }
                if (BottomTextBox.Text != "")
                {
                    Bottom = BottomTextBox.Text;
                }

                if (Bottle != Bottom)
                {
                    return;
                }

                CylinderNumbers = Bottle;

                //判斷是否滿箱
                using (conn = new SqlConnection(myConnectionString))
                {
                    conn.Open();

                    selectCmd = "SELECT WhereSeat FROM [ShippingBody] where [ListDate]='" + ListDate_LB.SelectedItem + "' and [ProductName]='" + ProductName_CB.SelectedItem + "' and [WhereBox]='" + WhereBox_LB.SelectedItem + "' order by Convert(INT,[WhereSeat]) DESC ";
                    cmd = new SqlCommand(selectCmd, conn);
                    using (reader = cmd.ExecuteReader())
                    {
                        if (reader.HasRows)
                        {
                            if (reader.Read())
                            {
                                NowSeat = reader.GetString(reader.GetOrdinal("WhereSeat"));
                                WhereSeatLabel.Text = (Convert.ToInt32(reader.GetString(reader.GetOrdinal("WhereSeat"))) + 2).ToString();

                                if (NowSeat == Aboxof())
                                {
                                    BottomTextBox.Text = "";
                                    MessageBox.Show("此嘜頭已滿箱", "警告", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                                    return;
                                }
                            }
                        }
                        else
                        {
                            NowSeat = "0";
                        }
                    }
                }

                //20200420
                try
                {
                    var v = (from p in DT.AsEnumerable()
                             where p.Field<string>("CylinderNo").Trim() == CylinderNumbers
                             select p).First();

                    LotNumber = v.Field<string>("vchManufacturingNo");
                    HydrostaticTestDate = v.Field<string>("vchHydrostaticTestDate");
                    CustomerName = v.Field<string>("ClientName");
                    MarkingType = v.Field<string>("vchMarkingType");
                    HydroLabelPass = v.Field<bool>("HydroLabelPass");
                }
                catch (Exception)
                {
                    BottomTextBox.Text = "";
                    MessageBox.Show("查無序號，請聯繫MIS", "警告", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                    return;
                }

                ProductType = Product_L.Text;

                using (conn = new SqlConnection(myConnectionString))
                {
                    conn.Open();

                    //抓產品型號
                    selectCmd = "select [Product_NO] from [Manufacturing] where [Manufacturing_NO] = '" + LotNumber + "'";
                    cmd = new SqlCommand(selectCmd, conn);
                    using (reader = cmd.ExecuteReader())
                    {
                        if (reader.Read())
                        {
                            ProductNo = reader.GetString(0);
                        }
                    }

                    //20210930拿掉
                    //selectCmd = "SELECT isnull([H_SpecialUses],'N') FROM [Manufacturing] where [Manufacturing_NO]='" + LotNumber + "' ";
                    //cmd = new SqlCommand(selectCmd, conn);
                    //using (reader = cmd.ExecuteReader())
                    //{
                    //    if (reader.Read())
                    //    {
                    //        if (reader.GetValue(0).ToString() == "Y")
                    //        {
                    //            SpecialUses = "Y";
                    //        }
                    //    }
                    //}

                    //判斷是否有成品檢驗報告
                    selectCmd = "SELECT * FROM [QC_ProductAcceptanceHead]" +
                        " WHERE ManufacturingNo = @LotNo AND QualifiedQuantity > 0 ";
                    cmd = new SqlCommand(selectCmd, conn);
                    cmd.Parameters.AddWithValue("@LotNo", LotNumber);
                    using (reader = cmd.ExecuteReader())
                    {
                        if (reader.Read())
                        {
                            ProductAcceptance = true;
                        }
                    }

                    //報廢
                    selectCmd = "SELECT  * FROM [RePortScrapReason] where [ScrapCylinderNO] = @CylinderNo";
                    cmd = new SqlCommand(selectCmd, conn);
                    cmd.Parameters.AddWithValue("@CylinderNo", CylinderNumbers);
                    using (reader = cmd.ExecuteReader())
                    {
                        if (reader.Read())
                        {
                            Error += "Code：101 此序號之氣瓶為報廢氣瓶，不允許加入\n";
                        }
                    }

                    //隔離
                    selectCmd = "SELECT [ID] FROM [ManufacturingIsolation] WHERE [CylinderNo] = @CylinderNo";
                    cmd = new SqlCommand(selectCmd, conn);
                    cmd.Parameters.AddWithValue("@CylinderNo", CylinderNumbers);
                    using (reader = cmd.ExecuteReader())
                    {
                        if (reader.Read())
                        {
                            Error += "Code：201 氣瓶已被隔離，不允許加入，請聯絡品保\n";
                        }
                    }

                    using (conn1 = new SqlConnection(myConnectionString30))
                    {
                        conn1.Open();

                        DateTime HydrostaticDate = Convert.ToDateTime(HydrostaticTestDate);

                        if (HydroLabelPass == true)
                        {
                            selectCmd1 = "SELECT [TestDate] FROM [PPT_Hydro_Details]" +
                                " WHERE [SerialNo] = @SN  order by id desc ";
                            cmd1 = new SqlCommand(selectCmd1, conn1);
                            cmd1.Parameters.AddWithValue("@SN", CylinderNumbers);
                            using (reader1 = cmd1.ExecuteReader())
                            {
                                if (reader1.Read())
                                {
                                    HydroDate = reader1.GetDateTime(reader1.GetOrdinal("TestDate"));
                                }
                                else
                                {
                                    //內膽不檢查水壓報告
                                    if (!ProductNo.Contains("-L-"))
                                    {
                                        Error += "Code：103 無水壓報告資料，請聯繫品保\n";
                                    }
                                }
                            }
                        }
                        else
                        {
                            selectCmd1 = "SELECT [TestDate] FROM [PPT_Hydro_Details]" +
                                " WHERE [SerialNo] = @SN and [TestDate] between '" + HydrostaticDate.ToString("yyyy-MM-dd") + "' and '" + HydrostaticDate.AddMonths(2).ToString("yyyy-MM-dd") + "' order by id desc ";
                            cmd1 = new SqlCommand(selectCmd1, conn1);
                            cmd1.Parameters.AddWithValue("@SN", CylinderNumbers);
                            using (reader1 = cmd1.ExecuteReader())
                            {
                                if (reader1.Read())
                                {
                                    HydroDate = reader1.GetDateTime(reader1.GetOrdinal("TestDate"));
                                }
                                else
                                {
                                    //內膽不檢查水壓報告
                                    if (!ProductNo.Contains("-L-"))
                                    {
                                        Error += "Code：103 無水壓報告資料，請聯繫品保\n";
                                    }
                                }
                            }
                        }
                    }

                    //判斷水壓年月是否大於規定範圍
                    selectCmd = "SELECT [HydroDate] FROM [ShippingHydroDateRestrictions] WHERE [BoxNo] = @BN";
                    cmd = new SqlCommand(selectCmd, conn);
                    cmd.Parameters.AddWithValue("@BN", WhereBox_LB.SelectedItem);
                    using (reader = cmd.ExecuteReader())
                    {
                        if (reader.HasRows)
                        {
                            if (reader.Read())
                            {
                                ResrictionDate = reader.GetDateTime(reader.GetOrdinal("HydroDate"));

                                if (HydroDate < ResrictionDate)
                                {
                                    Error += "Code：104 此序號水壓年月不在規定範圍內，請聯繫生管\n";
                                }
                            }
                        }
                    }

                    //判斷是否已經有相同的序號入嘜頭
                    selectCmd = "SELECT  * FROM [ShippingBody] where [CylinderNumbers]='" + CylinderNumbers + "'";
                    cmd = new SqlCommand(selectCmd, conn);
                    using (reader = cmd.ExecuteReader())
                    {
                        if (reader.Read())
                        {
                            Error += "Code：102 此序號已入嘜\n";
                        }
                    }

                    //檢查打字形式是否相同
                    selectCmd = "SELECT [Marking] FROM [ShippingHead] WHERE [Marking] = @Marking AND [vchBoxs] = @Box";
                    cmd = new SqlCommand(selectCmd, conn);
                    cmd.Parameters.AddWithValue("@Marking", MarkingType);
                    cmd.Parameters.AddWithValue("@Box", WhereBox_LB.SelectedItem);
                    using (reader = cmd.ExecuteReader())
                    {
                        if (reader.HasRows)
                        {
                            ;
                        }
                        else
                        {
                            Error += "Code：105 氣瓶打印形式與訂單不符，請聯繫生管\n";
                        }
                    }
                }

                // 照片檢查
                if (Product_L.Text.Contains("Composite") == true)
                {
                    using (conn = new SqlConnection(myConnectionString30))
                    {
                        conn.Open();
                        selectCmd = "select ID from CH_ShippingInspectionPhoto where MNO='" + LotNumber + "'" +
                        " and DATEDIFF(MONTH,([HydrostaticTestDate]+'/01'),@HydrostaticTestDate) BETWEEN -1 AND 0 and CustomerName='" + CustomerName + "'";
                        cmd = new SqlCommand(selectCmd, conn);
                        cmd.Parameters.Add("@HydrostaticTestDate", SqlDbType.VarChar).Value = HydrostaticTestDate + "/01";
                        using (reader = cmd.ExecuteReader())
                        {
                            if (!reader.HasRows)
                            {
                                if (!ProductNo.Contains("-L-"))
                                {
                                    Error += "Code：106 沒有客戶產品照片，請聯繫品保\n";
                                }
                            }
                        }
                    }
                }
                else if (Product_L.Text.Contains("Aluminum") == true)
                {
                    using (conn = new SqlConnection(myConnectionString30))
                    {
                        conn.Open();

                        selectCmd = "select ID from ProductPhotoCheck where [ManufacturingNo] = '" + LotNumber + "'" +
                            " and HydrostaticTestDate = @HydrostaticTestDate ";
                        cmd = new SqlCommand(selectCmd, conn);
                        cmd.Parameters.Add("@HydrostaticTestDate", SqlDbType.VarChar).Value = HydrostaticTestDate;
                        using (reader = cmd.ExecuteReader())
                        {
                            if (!reader.HasRows)
                            {
                                Error += "Code：124 沒有產品照片，請聯繫品保\n";
                            }
                        }
                    }
                }

                if (SpecialUses == "N")
                {/*
                    using (conn = new SqlConnection(myConnectionString))
                    {
                        conn.Open();

                        selectCmd = "SELECT  * FROM [HydrostaticPass] where [ManufacturingNo]='" + LotNumber + "' and [CylinderNo]='" + CylinderNumbers + "' and [HydrostaticPass]='Y'";
                        cmd = new SqlCommand(selectCmd, conn);
                        using (reader = cmd.ExecuteReader())
                        {
                            if (reader.Read())
                            {
                                HydrostaticPass = true;
                            }
                        }

                        if (HydrostaticPass == false)
                        {
                            //找對應的舊序號，若有序號則依此序號查是否有做過水壓
                            string OriCNo = "", OriMNO = "";

                            selectCmd = "SELECT  OriCylinderNo, OriManufacturingNo FROM [ChangeCylinderNo] where [NewManufacturingNo]='" + LotNumber + "' and [NewCylinderNo]='" + CylinderNumbers + "' ";
                            cmd = new SqlCommand(selectCmd, conn);
                            using (reader = cmd.ExecuteReader())
                            {
                                if (reader.Read())
                                {
                                    OriCNo = reader.GetString(reader.GetOrdinal("OriCylinderNo"));
                                    OriMNO = reader.GetString(reader.GetOrdinal("OriManufacturingNo"));
                                }
                            }

                            if (OriCNo != "")
                            {
                                selectCmd = "SELECT  * FROM [HydrostaticPass] where [ManufacturingNo]='" + OriMNO + "' and [CylinderNo]='" + OriCNo + "' and [HydrostaticPass]='Y'";
                                cmd = new SqlCommand(selectCmd, conn);
                                using (reader = cmd.ExecuteReader())
                                {
                                    if (reader.Read())
                                    {
                                        HydrostaticPass = true;
                                    }
                                }
                            }
                        }
                    }

                    if (HydrostaticPass == false)
                    {
                        Error += "Code：107 此序號查詢不到水壓測試資料，請聯繫品保\n";
                    }*/
                }

                //判別是否有做過成品檢驗
                //研發瓶轉正式出貨產品時，有可能之前的研發瓶試認證瓶所以沒有成品檢驗，因此要有成品檢驗的記錄
                if (ProductAcceptance == false)
                {
                    string OriMNO = "";

                    using (conn = new SqlConnection(myConnectionString))
                    {
                        conn.Open();
                        //找是否有對應之批號，有則依此搜尋是否有做成品檢驗
                        selectCmd = "SELECT  OriManufacturingNo FROM [TransformProductNo] where TransManufacturingNo='" + LotNumber + "'";
                        cmd = new SqlCommand(selectCmd, conn);
                        using (reader = cmd.ExecuteReader())
                        {
                            if (reader.Read())
                            {
                                OriMNO = reader.GetString(reader.GetOrdinal("OriManufacturingNo"));
                            }
                        }

                        if (OriMNO != "")
                        {
                            selectCmd = "SELECT   * FROM [QC_ProductAcceptanceHead] where ManufacturingNo='" + OriMNO + "' and QualifiedQuantity > 0";
                            cmd = new SqlCommand(selectCmd, conn);
                            using (reader = cmd.ExecuteReader())
                            {
                                if (reader.Read())
                                {
                                    ProductAcceptance = true;
                                }
                            }
                        }
                    }

                    if (ProductAcceptance == false)
                    {
                        Error += "Code：108 此序號查詢不到成品檢驗資料，請聯繫品保\n";
                    }
                }

                //判別產品類型
                if (ProductType.Contains("Aluminum"))
                {
                    if (ProductNo != "")
                    {
                        if (ProductNo.Contains("1-A-") == true)
                        {
                            using (conn = new SqlConnection(myConnectionString30))
                            {
                                conn.Open();

                                //彎曲
                                selectCmd = "SELECT * FROM [PPT_FlatBend] WHERE [ManufacturingNo] = '" + LotNumber + "' and FinalResult='PASS' and (Method='彎曲' or Method='壓扁') ";
                                cmd = new SqlCommand(selectCmd, conn);
                                using (reader = cmd.ExecuteReader())
                                {
                                    if (reader.HasRows)
                                    {
                                        ;
                                    }
                                    else
                                    {
                                        Error += "Code：109 無彎曲或壓扁資料，請聯繫品保\n";
                                    }
                                }

                                //拉伸
                                selectCmd = "SELECT  * FROM [PPT_Tensile] WHERE [ManufacturingNo] = '" + LotNumber + "' and FinalResult='PASS' ";
                                cmd = new SqlCommand(selectCmd, conn);
                                using (reader = cmd.ExecuteReader())
                                {
                                    if (reader.HasRows)
                                    {
                                        ;
                                    }
                                    else
                                    {
                                        Error += "Code：111 無拉伸資料，請聯繫品保\n";
                                    }
                                }

                                //硬度
                                selectCmd = "SELECT * FROM QCDocument INNER JOIN Esign2 ON QCDocument.AcceptanceNo = Esign2.AcceptanceNo WHERE (QCDocument.LotNo = '" + LotNumber + "') AND (Esign2.Type LIKE '硬度%')";
                                cmd = new SqlCommand(selectCmd, conn);
                                using (reader = cmd.ExecuteReader())
                                {
                                    if (reader.HasRows)
                                    {
                                        ;
                                    }
                                    else
                                    {
                                        Error += "Code：112 無硬度資料，請聯繫品保\n";
                                    }
                                }

                                //爆破
                                selectCmd = "SELECT  * FROM [PPT_Burst] WHERE [ManufacturingNo] = '" + LotNumber + "' and FinalResult='PASS' order by AcceptanceNo desc";
                                cmd = new SqlCommand(selectCmd, conn);
                                using (reader = cmd.ExecuteReader())
                                {
                                    if (reader.HasRows)
                                    {
                                        ;
                                    }
                                    else
                                    {
                                        Error += "Code：113 無爆破資料，請聯繫品保\n";
                                    }
                                }
                            }
                        }
                        else if (ProductNo.Contains("3-A-") == true)
                        {
                            using (conn = new SqlConnection(myConnectionString30))
                            {
                                conn.Open();

                                //拉伸
                                selectCmd = "SELECT  * FROM [PPT_Tensile] WHERE [ManufacturingNo] = '" + LotNumber + "' and FinalResult='PASS' ";
                                cmd = new SqlCommand(selectCmd, conn);
                                using (reader = cmd.ExecuteReader())
                                {
                                    if (reader.HasRows)
                                    {
                                        ;
                                    }
                                    else
                                    {
                                        Error += "Code：111 無拉伸資料，請聯繫品保\n";
                                    }
                                }

                                //壓扁
                                selectCmd = "SELECT * FROM [PPT_FlatBend] WHERE [ManufacturingNo] = '" + LotNumber + "' and FinalResult='PASS' and Method='壓扁' ";
                                cmd = new SqlCommand(selectCmd, conn);
                                using (reader = cmd.ExecuteReader())
                                {
                                    if (reader.HasRows)
                                    {
                                        ;
                                    }
                                    else
                                    {
                                        Error += "Code：110 無壓扁資料，請聯繫品保\n";
                                    }
                                }
                            }
                        }
                        else if (ProductNo.Contains("5-A-") == true)
                        {
                            using (conn = new SqlConnection(myConnectionString30))
                            {
                                conn.Open();

                                //爆破
                                selectCmd = "SELECT  * FROM [PPT_Burst] WHERE [ManufacturingNo] = '" + LotNumber + "' and FinalResult='PASS' order by AcceptanceNo desc";
                                cmd = new SqlCommand(selectCmd, conn);
                                using (reader = cmd.ExecuteReader())
                                {
                                    if (reader.HasRows)
                                    {
                                        ;
                                    }
                                    else
                                    {
                                        Error += "Code：113 無爆破資料，請聯繫品保\n";
                                    }
                                }
                            }
                        }
                    }
                }
                else if (ProductType.Contains("Composite"))
                {
                    string ResinLotNo = "", CarbonLotNo = "", GlassLotNo = "";
                    string CarbonSpec = "", GlassSpec = "";

                    using (conn = new SqlConnection(myConnectionString30))
                    {
                        conn.Open();

                        //判別是否有做出貨檢驗，無出貨檢驗資料不允許包裝
                        selectCmd = "SELECT  * FROM  CH_ShippingInspection where SerialNo='" + CylinderNumbers + "'";
                        cmd = new SqlCommand(selectCmd, conn);
                        using (reader = cmd.ExecuteReader())
                        {
                            if (reader.Read())
                            {
                                ;
                            }
                            else
                            {
                                Error += "Code：114 無出貨檢驗資料，請聯繫品保\n";
                            }
                        }

                        //爆破
                        selectCmd = "SELECT  * FROM [PPT_Burst] WHERE [ManufacturingNo] = '" + LotNumber + "' and FinalResult='PASS' order by AcceptanceNo desc";
                        cmd = new SqlCommand(selectCmd, conn);
                        using (reader = cmd.ExecuteReader())
                        {
                            if (reader.HasRows)
                            {
                                ;
                            }
                            else
                            {
                                Error += "Code：113 無爆破資料，請聯繫品保\n";
                            }
                        }

                        //循環
                        selectCmd = "SELECT  * FROM [PPT_Cycling] WHERE [LotNo] = '" + LotNumber + "' and FinalResult='PASS'";
                        cmd = new SqlCommand(selectCmd, conn);
                        using (reader = cmd.ExecuteReader())
                        {
                            if (reader.HasRows)
                            {
                                ;
                            }
                            else
                            {
                                if (!ProductNo.Contains("-L-"))
                                {
                                    Error += "Code：117 無循環資料，請聯繫品保\n";
                                }
                            }
                        }

                        selectCmd = "SELECT  ResinLotNo, CarbonLotNo, GlassLotNo, CarbonSpec, GlassSpec FROM [FilamentWinding] WHERE [LotNo] = '" + LotNumber + "' order by id desc";
                        cmd = new SqlCommand(selectCmd, conn);
                        using (reader = cmd.ExecuteReader())
                        {
                            if (reader.Read())
                            {
                                ResinLotNo = reader.GetString(0);
                                CarbonLotNo = reader.GetString(1);
                                GlassLotNo = reader.GetString(2);
                                CarbonSpec = reader.GetString(3);
                                GlassSpec = reader.GetString(4);
                            }
                        }
                    }

                    if (ResinLotNo != "" && CarbonLotNo != "" && GlassLotNo != "")
                    {
                        using (conn = new SqlConnection(myConnectionString30))
                        {
                            conn.Open();

                            //碳纖
                            selectCmd = "SELECT * FROM [IQC] A, [Esign2] B WHERE A.[AcceptanceNo]=B.[AcceptanceNo] AND A.[Type] = '碳纖' AND A.[LotNo] LIKE '%" + CarbonLotNo + "%'";
                            cmd = new SqlCommand(selectCmd, conn);
                            using (reader = cmd.ExecuteReader())
                            {
                                if (reader.Read())
                                {
                                    ;
                                }
                                else
                                {
                                    Error += "Code：118 無碳纖檢驗資料，請聯繫品保\n";
                                }
                            }

                            //玻纖
                            selectCmd = "SELECT * FROM [PPT] A, [Esign2] B WHERE A.[AcceptanceNo]=B.[AcceptanceNo] AND A.[Type] = '玻纖' AND A.[LotNo] LIKE '%" + GlassLotNo + "%'";
                            cmd = new SqlCommand(selectCmd, conn);
                            using (reader = cmd.ExecuteReader())
                            {
                                if (reader.Read())
                                {
                                    ;
                                }
                                else
                                {
                                    Error += "Code：119 無玻纖檢驗資料，請聯繫品保\n";
                                }
                            }

                            //樹酯
                            selectCmd = "SELECT * FROM [PPT] A, [Esign2] B WHERE A.[AcceptanceNo]=B.[AcceptanceNo] AND A.[Type] = '樹脂' AND A.[LotNo] LIKE '%" + ResinLotNo + "%' and FiberType ='玻' and (FiberLotNo like '%" + GlassLotNo + "%' or FiberSpec like '%" + GlassSpec + "%')";//20180912品保系統檢驗組組長 說只要規格一樣沒有對應批號也可以。當初為CE0086有問題
                            cmd = new SqlCommand(selectCmd, conn);
                            using (reader = cmd.ExecuteReader())
                            {
                                if (reader.Read())
                                {
                                    ;
                                }
                                else
                                {
                                    Error += "Code：120 無樹脂檢驗資料，請聯繫品保\n";
                                }
                            }

                            selectCmd = "SELECT * FROM [PPT] A, [Esign2] B WHERE A.[AcceptanceNo]=B.[AcceptanceNo] AND A.[Type] = '樹脂' AND A.[LotNo] LIKE '%" + ResinLotNo + "%' and FiberType ='碳' and (FiberLotNo like '%" + CarbonLotNo + "%' or FiberSpec like '%" + CarbonSpec + "%')";//20180912品保系統檢驗組組長 說只要規格一樣沒有對應批號也可以。當初為CE0086有問題
                            cmd = new SqlCommand(selectCmd, conn);
                            using (reader = cmd.ExecuteReader())
                            {
                                if (reader.Read())
                                {
                                    ;
                                }
                                else
                                {
                                    Error += "Code：120 無樹脂檢驗資料，請聯繫品保\n";
                                }
                            }
                        }
                    }

                    //對應內膽  拉伸、爆破
                    //找出對應內膽批號
                    string BuildUp = "";

                    using (conn = new SqlConnection(AMS21_ConnectionString))
                    {
                        conn.Open();
                        selectCmd = "SELECT [LinerLotNo] FROM [AMS_DATA].[dbo].[ComCylinderNo]" +
                            " WHERE [CylinderNo] = @CylinderNo";
                        cmd = new SqlCommand(selectCmd, conn);
                        cmd.Parameters.AddWithValue("@CylinderNo", CylinderNumbers);
                        using (reader = cmd.ExecuteReader())
                        {
                            if (reader.Read())
                            {
                                BuildUp = reader.GetString(reader.GetOrdinal("LinerLotNo"));
                            }
                        }
                    }

                    if (BuildUp != "")
                    {
                        using (conn = new SqlConnection(myConnectionString30))
                        {
                            conn.Open();

                            selectCmd = "SELECT  * FROM [PPT_Tensile]" +
                                " WHERE [ManufacturingNo] = @LotNo" +
                                " AND FinalResult = 'PASS' ";
                            cmd = new SqlCommand(selectCmd, conn);
                            cmd.Parameters.AddWithValue("@LotNo", BuildUp);
                            using (reader = cmd.ExecuteReader())
                            {
                                if (reader.HasRows)
                                {
                                    ;
                                }
                                else
                                {
                                    Error += "Code：115 無對應內膽(" + BuildUp + ")拉伸資料，請聯繫品保\n";
                                }
                            }

                            selectCmd = "SELECT  * FROM [PPT_Burst]" +
                                " WHERE [ManufacturingNo] = @LotNo" +
                                " AND [FinalResult] ='PASS' order by AcceptanceNo desc";
                            cmd = new SqlCommand(selectCmd, conn);
                            cmd.Parameters.AddWithValue("@LotNo", BuildUp);
                            using (reader = cmd.ExecuteReader())
                            {
                                if (reader.HasRows)
                                {
                                    ;
                                }
                                else
                                {
                                    Error += "Code：116 無對應內膽(" + BuildUp + ")爆破資料，請聯繫品保\n";
                                }
                            }
                        }
                    }
                }

                //20200702 客戶序號檢查
                string CustomerCylinderNo = string.Empty;

                using (conn = new SqlConnection(myConnectionString))
                {
                    conn.Open();

                    selectCmd = "select isnull([CustomerCylinderNo],'N') CustomerCylinderNo from [MSNBody] where [CylinderNo] = @CylinderNo";
                    cmd = new SqlCommand(selectCmd, conn);
                    cmd.Parameters.Add("CylinderNo", SqlDbType.VarChar).Value = CylinderNumbers;
                    using (reader = cmd.ExecuteReader())
                    {
                        if (reader.Read())
                        {
                            CustomerCylinderNo = reader.GetString(reader.GetOrdinal("CustomerCylinderNo"));
                        }
                    }

                    if (CustomerCylinderNo != "N" && CustomerCylinderNo != "")
                    {
                        selectCmd = "select count(ID) count from MSNBody where CustomerCylinderNo = @CustomerCylinderNo ";
                        cmd = new SqlCommand(selectCmd, conn);
                        cmd.Parameters.Add("@CustomerCylinderNo", SqlDbType.VarChar).Value = CustomerCylinderNo;
                        using (reader = cmd.ExecuteReader())
                        {
                            if (reader.Read())
                            {
                                if (reader.GetInt32(reader.GetOrdinal("count")) > 1)
                                {
                                    Error += "Code：121 此客戶序號以重複";
                                }
                            }
                        }

                        using (conn1 = new SqlConnection(myConnectionString21_AMS_check))
                        {
                            conn1.Open();

                            selectCmd1 = "select ID from [CylinderNoCheck_Q] where CylinderNo = @CylinderNo ";
                            cmd1 = new SqlCommand(selectCmd1, conn1);
                            cmd1.Parameters.Add("@CylinderNo", SqlDbType.VarChar).Value = CylinderNumbers;
                            using (reader1 = cmd1.ExecuteReader())
                            {
                                if (reader1.HasRows)
                                {
                                    ;
                                }
                                else
                                {
                                    //Error += "Code：122 品保未確認客戶序號";
                                }
                            }

                            selectCmd1 = "select ID from [CylinderNoCheck_P] where CylinderNo = @CylinderNo ";
                            cmd1 = new SqlCommand(selectCmd1, conn1);
                            cmd1.Parameters.Add("@CylinderNo", SqlDbType.VarChar).Value = CylinderNumbers;
                            using (reader1 = cmd1.ExecuteReader())
                            {
                                if (reader1.HasRows)
                                {
                                    ;
                                }
                                else
                                {
                                    //Error += "Code：123 生產未確認客戶序號";
                                }
                            }
                        }
                    }
                }

                if (Error.Any())
                {
                    BottomTextBox.Text = "";
                    Message = true;

                    DialogResult result = MessageBox.Show(Error, "警告", MessageBoxButtons.OK, MessageBoxIcon.Warning);

                    if (result == DialogResult.OK)
                    {
                        Message = false;
                        return;
                    }
                }


                //20200617 新增客戶序號確認
                using (conn = new SqlConnection(myConnectionString))
                {
                    conn.Open();

                    selectCmd = "select isnull([CustomerCylinderNo],'N') CustomerCylinderNo from [MSNBody] where [CylinderNo] = @CylinderNo";
                    cmd = new SqlCommand(selectCmd, conn);
                    cmd.Parameters.Add("CylinderNo", SqlDbType.VarChar).Value = CylinderNumbers;
                    using (reader = cmd.ExecuteReader())
                    {
                        if (reader.Read())
                        {
                            if (reader.GetString(reader.GetOrdinal("CustomerCylinderNo")) != "N" && reader.GetString(reader.GetOrdinal("CustomerCylinderNo")) != "")
                            {
                                DialogResult result = MessageBox.Show("請確認客戶序號：" + reader.GetString(reader.GetOrdinal("CustomerCylinderNo")), "確認", MessageBoxButtons.OKCancel, MessageBoxIcon.Warning);

                                if (result == DialogResult.Cancel)
                                {
                                    return;
                                }
                            }
                        }
                    }
                }


                using (conn = new SqlConnection(myConnectionString))
                {
                    conn.Open();

                    selectCmd = "select vchManufacturingNo from MSNBody where CylinderNo = @CylinderNo ";
                    cmd = new SqlCommand(selectCmd, conn);
                    cmd.Parameters.Add("@CylinderNo", SqlDbType.VarChar).Value = CylinderNumbers;
                    using (reader = cmd.ExecuteReader())
                    {
                        if (reader.Read())
                        {
                            if (LotNumber != reader.GetString(reader.GetOrdinal("vchManufacturingNo")))
                            {
                                Message = true;

                                DialogResult result = MessageBox.Show("請聯繫MIS", "警告", MessageBoxButtons.OK, MessageBoxIcon.Warning);

                                if (result == DialogResult.OK)
                                {
                                    Message = false;
                                    return;
                                }
                            }
                        }
                    }
                }

                int InsertSB = 0, UpdateLP = 0, UpdateMsn = 0;

                using (TransactionScope scope = new TransactionScope())
                {
                    using (conn = new SqlConnection(myConnectionString))
                    {
                        conn.Open();

                        //雷刻掃描完確認瓶身瓶底相同後載入資料
                        selectCmd = "INSERT INTO [ShippingBody] ( ListDate, ProductName, CylinderNumbers, WhereBox, WhereSeat, vchUser, Time, LotNumber )" +
                            "VALUES ( @ListDate, @ProductName, @CylinderNumbers, @WhereBox, @WhereSeat, @vchUser, @Time, @LotNumber )";
                        cmd = new SqlCommand(selectCmd, conn);

                        cmd.Parameters.Add("@ListDate", SqlDbType.VarChar).Value = ListDate_LB.SelectedItem;
                        cmd.Parameters.Add("@ProductName", SqlDbType.VarChar).Value = ProductName_CB.SelectedItem;
                        cmd.Parameters.Add("@CylinderNumbers", SqlDbType.VarChar).Value = CylinderNumbers;
                        cmd.Parameters.Add("@WhereBox", SqlDbType.VarChar).Value = WhereBox_LB.SelectedItem;
                        cmd.Parameters.Add("@WhereSeat", SqlDbType.VarChar).Value = Convert.ToInt32(NowSeat) + 1;
                        cmd.Parameters.Add("@vchUser", SqlDbType.VarChar).Value = User_LB.Text.Remove(0, 7);
                        cmd.Parameters.Add("@Time", SqlDbType.VarChar).Value = DateTime.Now.ToLocalTime().ToString();
                        cmd.Parameters.Add("@LotNumber", SqlDbType.VarChar).Value = LotNumber;

                        InsertSB = cmd.ExecuteNonQuery();

                        //更新登出時間
                        selectCmd = "UPDATE [LoginPackage] SET  [LogoutTime] = '" + DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss") + "' , [IsUpdate]='0' " +
                            "WHERE [ID] = '" + toolStripStatusLabel1.Text + "'";
                        cmd = new SqlCommand(selectCmd, conn);
                        UpdateLP = cmd.ExecuteNonQuery();

                        selectCmd = "INSERT INTO [WorkTimePackage] ( CylinderNo, Operator, OperatorId, AddTime, Date, WorkType, ProcessNO ) " +
                            "VALUES ( @CylinderNo, @Operator, @OperatorId, @AddTime, @Date, @WorkType, @ProcessNO )";
                        cmd = new SqlCommand(selectCmd, conn);

                        cmd.Parameters.Add("@CylinderNo", SqlDbType.VarChar).Value = CylinderNumbers;
                        cmd.Parameters.Add("@Operator", SqlDbType.VarChar).Value = User;
                        cmd.Parameters.Add("@OperatorId", SqlDbType.VarChar).Value = ID;
                        cmd.Parameters.Add("@AddTime", SqlDbType.VarChar).Value = DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss");
                        cmd.Parameters.Add("@Date", SqlDbType.VarChar).Value = DateTime.Now.ToString("yyyy-MM-dd");
                        cmd.Parameters.Add("@WorkType", SqlDbType.VarChar).Value = worktype;
                        cmd.Parameters.Add("@ProcessNO", SqlDbType.VarChar).Value = ProcessNo;

                        //cmd.ExecuteNonQuery();

                        selectCmd = "update [MSNBody] set [Package]='1' where [CylinderNo]='" + CylinderNumbers + "'";
                        cmd = new SqlCommand(selectCmd, conn);
                        UpdateMsn = cmd.ExecuteNonQuery();
                    }

                    if (InsertSB != 0 && UpdateLP != 0 && UpdateMsn != 0)
                    {
                        scope.Complete();
                    }
                    else
                    {
                        MessageBox.Show("新增失敗，請重新新增", "警告", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                    }
                }

                time = 420;

                string BoxsListBoxIndex = "";
                string NowSeat2 = "";

                //用來自動跳下一箱     
                using (conn = new SqlConnection(myConnectionString))
                {
                    conn.Open();

                    selectCmd = "SELECT WhereSeat FROM [ShippingBody] where [ListDate]='" + ListDate_LB.SelectedItem + "' and [ProductName]='" + ProductName_CB.SelectedItem + "' and [WhereBox]='" + WhereBox_LB.SelectedItem + "' order by Convert(INT,[WhereSeat]) DESC ";
                    cmd = new SqlCommand(selectCmd, conn);
                    using (reader = cmd.ExecuteReader())
                    {
                        if (reader.Read())
                        {
                            NowSeat2 = reader.GetString(reader.GetOrdinal("WhereSeat"));
                            BoxsListBoxIndex = WhereBox_LB.SelectedIndex.ToString();

                            //如果箱號已經超過最大箱數則不自動跳箱
                            if ((Convert.ToInt32(BoxsListBoxIndex) >= (WhereBox_LB.Items.Count - 1)) && WhereBox_LB.Items.Count != 1 && NowSeat2 == Aboxof())
                            {
                                //ABoxofLabel
                                MessageBox.Show("此日期嘜頭已經完全結束", "提示");
                                BottleTextBox.Focus();
                                return;
                            }

                            if (NowSeat2 == Aboxof())
                            {
                                if (PrintCheckBox.Checked == true)
                                {
                                    PrintButton.PerformClick();
                                }
                                else
                                {
                                    WhereBox_LB.SelectedIndex = (Convert.ToInt32(BoxsListBoxIndex) + 1);
                                }
                            }
                        }
                    }
                }

                //載入入箱狀況的圖片
                LoadPictrue();

                //載入dataGridView資料
                LoadSQLDate();

                //清除TextBox
                BottleTextBox.Text = "";
                BottomTextBox.Text = "";

                BottleTextBox.Focus();

                //直接略過正常程序載入資料 N代表取消 Y代表執行
                Pass = "N";

                //用來表示已經完成一輪,需等待雷刻瓶身才可繼續輸入
                BeGin = "N";

                //提示此序號已經載入嘜頭
                TipTextLabel.Visible = true;

                AutoCheckTimer.Enabled = false;
            }
        }

        private void BottomTextBox_TextChanged(object sender, EventArgs e)
        {
            if ((BottomTextBox.Text == BottleTextBox.Text) && ((BottomTextBox.Text != "") || (BottleTextBox.Text != "")))
            {
                AutoCheckTimer.Enabled = true;
            }
        }

        private void BottleTextBox_TextChanged(object sender, EventArgs e)
        {
            TipTextLabel.Visible = false;
            if ((BottomTextBox.Text == BottleTextBox.Text) && ((BottomTextBox.Text != "") || (BottleTextBox.Text != "")))
            {
                AutoCheckTimer.Enabled = true;
            }
        }

        private void BottomTextBox_KeyDown(object sender, KeyEventArgs e)
        {
            if ((e.KeyValue != 16) && (e.KeyValue != 13))//16=SHIFT 13=ENTER
            {
                str += Convert.ToChar(e.KeyValue);

                if (str == TempStr2)
                {
                    return;
                }
            }

            if (e.KeyValue == 13)
            {
                TempStr2 = str;

                BottomTextBox.Text = str;
                str = "";
            }
        }

        private void BottleTextBox_KeyDown(object sender, KeyEventArgs e)
        {
            if ((e.KeyValue != 16) && (e.KeyValue != 13))//16=SHIFT 13=ENTER
            {
                str += Convert.ToChar(e.KeyValue);

                if (str == TempStr1)
                {
                    return;
                }
            }

            if (e.KeyValue == 13)
            {
                TempStr1 = str;

                BottleTextBox.Text = str;
                str = "";

                if (LinkLMCheckBox.Checked == false)
                {
                    BottomTextBox.Focus();
                }
            }
        }

        private void AutoLoadPictureTimer_Tick(object sender, EventArgs e)
        {
            //載入入箱狀況的圖片
            LoadPictrue();

            //載入dataGridView資料
            LoadSQLDate();
        }

        private void HistoryListBox_SelectedIndexChanged(object sender, EventArgs e)
        {
            HistoryListBox.SelectedItem = HistoryListBox.Items[HistoryListBox.Items.Count - 1];
        }

        private void BottleTextBox_KeyPress(object sender, KeyPressEventArgs e)
        {
            if (BeGin == "N" && LinkLMCheckBox.Checked == true)
            {
                if (e.KeyChar != (char)Keys.Back)
                {//如果按下的不是回退键，则取消本次(按键)动作
                    e.Handled = true;
                }
            }

            if (e.KeyChar == (char)Keys.Back)
            {
                e.Handled = true;
            }
        }

        private void BottomTextBox_KeyPress(object sender, KeyPressEventArgs e)
        {
            if (BeGin == "N" && LinkLMCheckBox.Checked == true)
            {
                if (e.KeyChar != (char)Keys.Back)
                {
                    //如果按下的不是回退键，则取消本次(按键)动作
                    e.Handled = true;
                }
            }

            if (e.KeyChar == (char)Keys.Back)
            {
                e.Handled = true;
            }
        }

        private void checkBox1_CheckedChanged(object sender, EventArgs e)
        {
            if (NoLMCheckBox.Checked == true)
            {
                LinkLMCheckBox.Checked = false;
                LinkLMCheckBox.Enabled = false;
                KeyInGroupBox.Visible = false;
                NoLMGroupBox.Visible = true;
            }
            else
            {
                LinkLMCheckBox.Enabled = true;
                KeyInGroupBox.Visible = true;
                NoLMGroupBox.Visible = false;
            }
        }

        private void NoLMCylinderNOTextBox_KeyDown(object sender, KeyEventArgs e)
        {
            if (e.KeyValue == 13)//16=SHIFT 13=ENTER
            {
                NoLMCylinderNOTextBox.Text = NoLMCylinderNOTextBox.Text.Trim();//移除前後空白，以防止找不到資料或系統Error
                //NoLMCylinderNOTextBox.Text = NoLMCylinderNOTextBox.Text.TrimEnd(' ');
                if (NoLMCylinderNOTextBox.Text == "")
                {
                    MessageBox.Show("請輸入第一隻氣瓶序號！", "警告W-004");
                    return;
                }
                else
                {
                    //20141029 修改成不跳出視窗，直接在該畫面作操作。因有不連號(跳號)，原方式耗時
                    //以按Enter表示某汽瓶序號裝箱，但系統不自動跳號(+1)；以按Enter表示某汽瓶序號裝箱，且系統自動跳號(+1)
                    if (ShippingCNO() == false)
                    {
                        return;
                    }

                    if (CheckCylinderNOTextBox() == true)
                    {
                        AutoAccumulate();
                        NoLMCylinderNOTextBox.SelectAll();
                    }
                }
            }
            else if (e.KeyValue == 32)//32=SPACE
            {
                //20141029 修改成不跳出視窗，直接在該畫面作操作。因有不連號(跳號)，原方式耗時
                //以按Enter表示某汽瓶序號裝箱，但系統不自動跳號(+1)；以按Enter表示某汽瓶序號裝箱，且系統自動跳號(+1)

                //讓序號加1
                NoLMCylinderNOTextBox.Text = NoLMCylinderNOTextBox.Text.Trim();
                //NoLMCylinderNOTextBox.Text = NoLMCylinderNOTextBox.Text.TrimEnd(' ');
                if (NoLMCylinderNOTextBox.Text == "")
                {
                    MessageBox.Show("請輸入第一隻氣瓶序號！", "警告W-004");
                    return;
                }
                else
                {
                    if (ShippingCNO() == false)
                    {
                        return;
                    }
                    if (CheckCylinderNOTextBox() == true)
                    {
                        AutoAccumulate();
                        //序號往下累加
                        NextNumber();
                    }
                }
            }
        }

        private bool ShippingCNO()
        {
            //找出客戶、國家
            string Client = "", City = "";

            using (conn = new SqlConnection(myConnectionString))
            {
                conn.Open();

                selectCmd = "SELECT  Client, City FROM ShippingHead where  ListDate='" + ListDate_LB.SelectedItem.ToString() + "' and  ProductName='" + ProductName_CB.Text.Trim().ToString() + "'and vchBoxs='" + WhereBox_LB.SelectedItem.ToString() + "' and City is not null";
                cmd = new SqlCommand(selectCmd, conn);
                using (reader = cmd.ExecuteReader())
                {
                    if (reader.Read())
                    {
                        Client = reader.GetString(reader.GetOrdinal("Client"));
                        City = reader.GetString(reader.GetOrdinal("City"));
                    }
                }

                //找出對應的
                selectCmd = "SELECT   Client, City FROM  ShippingCityCNo WHERE  ('" + NoLMCylinderNOTextBox.Text.Trim() + "' >= SCNO) AND ('" + NoLMCylinderNOTextBox.Text.Trim() + "' <= ECNO)";
                cmd = new SqlCommand(selectCmd, conn);
                using (reader = cmd.ExecuteReader())
                {
                    if (reader.Read())
                    {
                        if (Client == reader.GetString(reader.GetOrdinal("Client")) && City == reader.GetString(reader.GetOrdinal("City")))
                        {
                            return true;
                        }
                        else
                        {
                            MessageBox.Show("該序號歸屬" + reader.GetString(reader.GetOrdinal("City")));
                            return false;
                        }
                    }
                }
            }

            return true;
        }

        private bool CheckCylinderNOTextBox()
        {
            if (NoLMCylinderNOTextBox.Text.Length < 6 || NoLMCylinderNOTextBox.Text.Length > 12)
            {
                MessageBox.Show("所輸入之氣瓶序號長度錯誤，請重新輸入!", "提示");
                return false;
            }

            return true;
        }

        private void AutoAccumulate()
        {
            DateTime ResrictionDate = new DateTime();
            DateTime HydroDate = new DateTime();

            string MarkingType = string.Empty;
            string HydrostaticTestDate = string.Empty;
            string CustomerName = string.Empty;
            string NowSeat = string.Empty;
            string LotNumber = string.Empty;
            string Error = string.Empty;
            string CylinderNO = string.Empty;
            string ProductNo = string.Empty;
            string ProductType = string.Empty;

            bool ProductAcceptance = false;
            bool SpecialUses = false;
            bool HydroLabelPass = false;

            //判斷是否滿箱
            using (conn = new SqlConnection(myConnectionString))
            {
                conn.Open();

                selectCmd = "SELECT WhereSeat FROM [ShippingBody] where [ListDate]='" + ListDate_LB.SelectedItem + "' and [ProductName]='" + ProductName_CB.SelectedItem + "' and [WhereBox]='" + WhereBox_LB.SelectedItem + "' order by Convert(INT,[WhereSeat]) DESC ";
                cmd = new SqlCommand(selectCmd, conn);
                using (reader = cmd.ExecuteReader())
                {
                    if (reader.HasRows)
                    {
                        if (reader.Read())
                        {
                            NowSeat = reader.GetString(reader.GetOrdinal("WhereSeat"));
                            WhereSeatLabel.Text = (Convert.ToInt32(reader.GetString(reader.GetOrdinal("WhereSeat"))) + 2).ToString();

                            if (NowSeat == Aboxof())
                            {
                                MessageBox.Show("此嘜頭已滿箱", "警告", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                                return;
                            }
                        }
                    }
                    else
                    {
                        NowSeat = "0";
                    }
                }
            }

            try
            {
                var v = (from p in DT.AsEnumerable()
                         where p.Field<string>("CylinderNo").Trim() == NoLMCylinderNOTextBox.Text
                         select p).First();

                LotNumber = v.Field<string>("vchManufacturingNo");
                MarkingType = v.Field<string>("vchMarkingType");
                HydrostaticTestDate = v.Field<string>("vchHydrostaticTestDate");
                CustomerName = v.Field<string>("ClientName");
                CylinderNO = NoLMCylinderNOTextBox.Text;
                HydroLabelPass = v.Field<bool>("HydroLabelPass");
            }
            catch (Exception)
            {
                MessageBox.Show("查無序號，請聯繫MIS", "警告", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                return;
            }

            ProductType = Product_L.Text;

            using (conn = new SqlConnection(myConnectionString))
            {
                conn.Open();

                //抓產品型號
                selectCmd = "select [Product_NO] from [Manufacturing] where [Manufacturing_NO] = '" + LotNumber + "'";
                cmd = new SqlCommand(selectCmd, conn);
                using (reader = cmd.ExecuteReader())
                {
                    if (reader.Read())
                    {
                        ProductNo = reader.GetString(0);
                    }
                }

                //取得製造批號
                selectCmd = "SELECT isnull([H_SpecialUses],'N') FROM [Manufacturing] where [Manufacturing_NO]='" + LotNumber + "'";
                cmd = new SqlCommand(selectCmd, conn);
                using (reader = cmd.ExecuteReader())
                {
                    if (reader.Read())
                    {
                        if (reader.GetString(0) == "Y")
                        {
                            SpecialUses = true;
                        }
                    }
                }

                //判斷是否有成品檢驗報告
                selectCmd = "SELECT * FROM [QC_ProductAcceptanceHead] where ManufacturingNo='" + LotNumber + "' and QualifiedQuantity > 0 ";
                cmd = new SqlCommand(selectCmd, conn);
                using (reader = cmd.ExecuteReader())
                {
                    if (reader.Read())
                    {
                        ProductAcceptance = true;
                    }
                }

                //判別是否為報廢氣瓶
                selectCmd = "SELECT  * FROM [RePortScrapReason] where [ScrapCylinderNO]='" + CylinderNO + "'";
                cmd = new SqlCommand(selectCmd, conn);
                using (reader = cmd.ExecuteReader())
                {
                    if (reader.Read())
                    {
                        Error += "Code：101 此序號之氣瓶為報廢氣瓶，不允許加入\n";
                    }
                }

                //隔離
                selectCmd = "SELECT [ID] FROM [ManufacturingIsolation] WHERE [CylinderNo] = @CylinderNo";
                cmd = new SqlCommand(selectCmd, conn);
                cmd.Parameters.Add("@CylinderNo", SqlDbType.VarChar).Value = CylinderNO;
                using (reader = cmd.ExecuteReader())
                {
                    if (reader.Read())
                    {
                        Error += "Code：201 氣瓶已被隔離，不允許加入，請聯絡品保\n";
                    }
                }

                //判斷是否已經有相同的序號入嘜頭
                selectCmd = "SELECT  * FROM [ShippingBody] where [CylinderNumbers]='" + CylinderNO + "'";
                cmd = new SqlCommand(selectCmd, conn);
                using (reader = cmd.ExecuteReader())
                {
                    if (reader.Read())
                    {
                        Error += "Code：102 此序號已入嘜\n";
                    }
                }

                //判斷是否有水壓報告
                using (conn1 = new SqlConnection(myConnectionString30))
                {
                    conn1.Open();

                    DateTime HydrostaticDate = Convert.ToDateTime(HydrostaticTestDate);

                    if (HydroLabelPass == true)
                    {
                        selectCmd1 = "SELECT [TestDate] FROM [PPT_Hydro_Details] WHERE [SerialNo] = @SN order by id desc";
                        cmd1 = new SqlCommand(selectCmd1, conn1);
                        cmd1.Parameters.AddWithValue("@SN", CylinderNO);
                        using (reader1 = cmd1.ExecuteReader())
                        {
                            if (reader1.HasRows)
                            {
                                if (reader1.Read())
                                {
                                    HydroDate = reader1.GetDateTime(reader1.GetOrdinal("TestDate"));
                                }
                            }
                            else
                            {
                                //內膽不檢查水壓報告
                                if (!ProductNo.Contains("-L-"))
                                {
                                    Error += "Code：103 無水壓報告資料，請聯繫品保\n";
                                }
                            }
                        }
                    }
                    else
                    {
                        selectCmd1 = "SELECT [TestDate] FROM [PPT_Hydro_Details] WHERE [SerialNo] = @SN and [TestDate] between '" + HydrostaticDate.ToString("yyyy-MM-dd") + "' and '" + HydrostaticDate.AddMonths(2).ToString("yyyy-MM-dd") + "' order by id desc";
                        cmd1 = new SqlCommand(selectCmd1, conn1);
                        cmd1.Parameters.AddWithValue("@SN", CylinderNO);
                        using (reader1 = cmd1.ExecuteReader())
                        {
                            if (reader1.HasRows)
                            {
                                if (reader1.Read())
                                {
                                    HydroDate = reader1.GetDateTime(reader1.GetOrdinal("TestDate"));
                                }
                            }
                            else
                            {
                                //內膽不檢查水壓報告
                                if (!ProductNo.Contains("-L-"))
                                {
                                    Error += "Code：103 無水壓報告資料，請聯繫品保\n";
                                }
                            }
                        }
                    }
                }

                //判斷水壓年月是否大於規定範圍
                selectCmd = "SELECT [HydroDate] FROM [ShippingHydroDateRestrictions] WHERE [BoxNo] = @BN";
                cmd = new SqlCommand(selectCmd, conn);
                cmd.Parameters.AddWithValue("@BN", WhereBox_LB.SelectedItem);
                using (reader = cmd.ExecuteReader())
                {
                    if (reader.HasRows)
                    {
                        if (reader.Read())
                        {
                            ResrictionDate = reader.GetDateTime(reader.GetOrdinal("HydroDate"));

                            if (HydroDate < ResrictionDate)
                            {
                                Error += "Code：104 此序號水壓年月不在規定範圍內，請聯繫生管\n";
                            }
                        }
                    }
                }

                //檢查打字形式是否相同
                selectCmd = "SELECT [Marking] FROM [ShippingHead] WHERE [Marking] = @Marking AND [vchBoxs] = @Box";
                cmd = new SqlCommand(selectCmd, conn);
                cmd.Parameters.AddWithValue("@Marking", MarkingType);
                cmd.Parameters.AddWithValue("@Box", WhereBox_LB.SelectedItem);
                using (reader = cmd.ExecuteReader())
                {
                    if (reader.HasRows)
                    {
                        ;
                    }
                    else
                    {
                        Error += "Code：105 氣瓶打印形式與訂單不符，請聯繫生管\n";
                    }
                }
            }

            //照片檢查
            if (Product_L.Text.Contains("Composite") == true)
            {
                using (conn = new SqlConnection(myConnectionString30))
                {
                    conn.Open();
                    selectCmd = "select ID from CH_ShippingInspectionPhoto where MNO='" + LotNumber + "'" +
                        " and DATEDIFF(MONTH,([HydrostaticTestDate]+'/01'),@HydrostaticTestDate) BETWEEN -1 AND 0 and CustomerName='" + CustomerName + "'";
                    cmd = new SqlCommand(selectCmd, conn);
                    cmd.Parameters.Add("@HydrostaticTestDate", SqlDbType.VarChar).Value = HydrostaticTestDate + "/01";
                    using (reader = cmd.ExecuteReader())
                    {
                        if (!reader.HasRows)
                        {
                            if (!ProductNo.Contains("-L-"))
                            {
                                Error += "Code：106 沒有客戶產品照片，請聯繫品保\n";
                            }
                        }
                    }
                }
            }
            else if (Product_L.Text.Contains("Aluminum") == true)
            {
                using (conn = new SqlConnection(myConnectionString30))
                {
                    conn.Open();

                    selectCmd = "select ID from ProductPhotoCheck where [ManufacturingNo] = '" + LotNumber + "'" +
                        " and HydrostaticTestDate = @HydrostaticTestDate ";
                    cmd = new SqlCommand(selectCmd, conn);
                    cmd.Parameters.Add("@HydrostaticTestDate", SqlDbType.VarChar).Value = HydrostaticTestDate;
                    using (reader = cmd.ExecuteReader())
                    {
                        if (!reader.HasRows)
                        {
                            if (!ProductNo.Contains("-L-"))
                            {
                                Error += "Code：124 沒有產品照片，請聯繫品保\n";
                            }
                        }
                    }
                }
            }

            if (SpecialUses == false)
            {/*
                using (conn = new SqlConnection(myConnectionString))
                {
                    conn.Open();

                    selectCmd = "SELECT  * FROM [HydrostaticPass] where [ManufacturingNo]='" + LotNumber + "' and [CylinderNo]='" + NoLMCylinderNOTextBox.Text + "' and [HydrostaticPass]='Y'";
                    cmd = new SqlCommand(selectCmd, conn);
                    using (reader = cmd.ExecuteReader())
                    {
                        if (reader.Read())
                        {
                            HydrostaticPass = true;
                        }
                    }

                    if (HydrostaticPass == false)
                    {
                        //找對應的舊序號，若有序號則依此序號查是否有做過水壓
                        string OriCNo = "", OriMNO = "";

                        selectCmd = "SELECT  OriCylinderNo, OriManufacturingNo FROM [ChangeCylinderNo] where [NewManufacturingNo]='" + LotNumber + "' and [NewCylinderNo]='" + NoLMCylinderNOTextBox.Text + "' ";
                        cmd = new SqlCommand(selectCmd, conn);
                        using (reader = cmd.ExecuteReader())
                        {
                            if (reader.Read())
                            {
                                OriCNo = reader.GetString(reader.GetOrdinal("OriCylinderNo"));
                                OriMNO = reader.GetString(reader.GetOrdinal("OriManufacturingNo"));
                            }
                        }

                        if (OriCNo != "")
                        {
                            selectCmd = "SELECT  * FROM [HydrostaticPass] where [ManufacturingNo]='" + OriMNO + "' and [CylinderNo]='" + OriCNo + "' and [HydrostaticPass]='Y'";
                            cmd = new SqlCommand(selectCmd, conn);
                            using (reader = cmd.ExecuteReader())
                            {
                                if (reader.Read())
                                {
                                    HydrostaticPass = true;
                                }
                            }
                        }
                    }
                }

                if (HydrostaticPass == false)
                {
                    Error += "Code：107 此序號查詢不到水壓測試資料，請聯繫品保\n";
                }*/
            }

            //判別是否有做過成品檢驗
            //研發瓶轉正式出貨產品時，有可能之前的研發瓶試認證瓶所以沒有成品檢驗，因此要有成品檢驗的記錄
            if (ProductAcceptance == false)
            {
                string OriMNO = "";

                using (conn = new SqlConnection(myConnectionString))
                {
                    conn.Open();
                    //找是否有對應之批號，有則依此搜尋是否有做成品檢驗
                    selectCmd = "SELECT  OriManufacturingNo FROM [TransformProductNo] where TransManufacturingNo='" + LotNumber + "'";
                    cmd = new SqlCommand(selectCmd, conn);
                    using (reader = cmd.ExecuteReader())
                    {
                        if (reader.Read())
                        {
                            OriMNO = reader.GetString(reader.GetOrdinal("OriManufacturingNo"));
                        }
                    }

                    if (OriMNO != "")
                    {
                        selectCmd = "SELECT   * FROM [QC_ProductAcceptanceHead] where ManufacturingNo='" + OriMNO + "' and QualifiedQuantity > 0";
                        cmd = new SqlCommand(selectCmd, conn);
                        using (reader = cmd.ExecuteReader())
                        {
                            if (reader.Read())
                            {
                                ProductAcceptance = true;
                            }
                        }
                    }
                }

                if (ProductAcceptance == false)
                {
                    Error += "Code：108 此序號查詢不到成品檢驗資料，請聯繫品保\n";
                }
            }

            //判別產品類型
            if (ProductType.Contains("Aluminum"))
            {
                if (ProductNo != "")
                {
                    if (ProductNo.Contains("1-A-") == true)
                    {
                        using (conn = new SqlConnection(myConnectionString30))
                        {
                            conn.Open();

                            //彎曲
                            selectCmd = "SELECT * FROM [PPT_FlatBend] WHERE [ManufacturingNo] = '" + LotNumber + "' and FinalResult='PASS' and ( Method='彎曲' or Method='壓扁') ";
                            cmd = new SqlCommand(selectCmd, conn);
                            using (reader = cmd.ExecuteReader())
                            {
                                if (reader.HasRows)
                                {
                                    ;
                                }
                                else
                                {
                                    Error += "Code：109 無彎曲或壓扁資料，請聯繫品保\n";
                                }
                            }

                            //拉伸
                            selectCmd = "SELECT  * FROM [PPT_Tensile] WHERE [ManufacturingNo] = '" + LotNumber + "' and FinalResult='PASS' ";
                            cmd = new SqlCommand(selectCmd, conn);
                            using (reader = cmd.ExecuteReader())
                            {
                                if (reader.HasRows)
                                {
                                    ;
                                }
                                else
                                {
                                    Error += "Code：111 無拉伸資料，請聯繫品保\n";
                                }
                            }

                            //硬度
                            selectCmd = "SELECT * FROM QCDocument INNER JOIN Esign2 ON QCDocument.AcceptanceNo = Esign2.AcceptanceNo WHERE (QCDocument.LotNo = '" + LotNumber + "') AND (Esign2.Type LIKE '硬度%')";
                            cmd = new SqlCommand(selectCmd, conn);
                            using (reader = cmd.ExecuteReader())
                            {
                                if (reader.HasRows)
                                {
                                    ;
                                }
                                else
                                {
                                    Error += "Code：112 無硬度資料，請聯繫品保\n";
                                }
                            }

                            //爆破
                            selectCmd = "SELECT  * FROM [PPT_Burst] WHERE [ManufacturingNo] = '" + LotNumber + "' and FinalResult='PASS' order by AcceptanceNo desc";
                            cmd = new SqlCommand(selectCmd, conn);
                            using (reader = cmd.ExecuteReader())
                            {
                                if (reader.HasRows)
                                {
                                    ;
                                }
                                else
                                {
                                    Error += "Code：113 無爆破資料，請聯繫品保\n";
                                }
                            }
                        }
                    }
                    else if (ProductNo.Contains("3-A-") == true)
                    {
                        using (conn = new SqlConnection(myConnectionString30))
                        {
                            conn.Open();

                            //拉伸
                            selectCmd = "SELECT  * FROM [PPT_Tensile] WHERE [ManufacturingNo] = '" + LotNumber + "' and FinalResult='PASS' ";
                            cmd = new SqlCommand(selectCmd, conn);
                            using (reader = cmd.ExecuteReader())
                            {
                                if (reader.HasRows)
                                {
                                    ;
                                }
                                else
                                {
                                    Error += "Code：111 無拉伸資料，請聯繫品保\n";
                                }
                            }

                            //壓扁
                            selectCmd = "SELECT * FROM [PPT_FlatBend] WHERE [ManufacturingNo] = '" + LotNumber + "' and FinalResult='PASS' and Method='壓扁' ";
                            cmd = new SqlCommand(selectCmd, conn);
                            using (reader = cmd.ExecuteReader())
                            {
                                if (reader.HasRows)
                                {
                                    ;
                                }
                                else
                                {
                                    Error += "Code：110 無壓扁資料，請聯繫品保\n";
                                }
                            }
                        }
                    }
                    else if (ProductNo.Contains("5-A-") == true)
                    {
                        using (conn = new SqlConnection(myConnectionString30))
                        {
                            conn.Open();

                            //爆破
                            selectCmd = "SELECT  * FROM [PPT_Burst] WHERE [ManufacturingNo] = '" + LotNumber + "' and FinalResult='PASS' order by AcceptanceNo desc";
                            cmd = new SqlCommand(selectCmd, conn);
                            using (reader = cmd.ExecuteReader())
                            {
                                if (reader.HasRows)
                                {
                                    ;
                                }
                                else
                                {
                                    Error += "Code：113 無爆破資料，請聯繫品保\n";
                                }
                            }
                        }
                    }
                }
            }
            else if (ProductType.Contains("Composite"))
            {
                string ResinLotNo = "", CarbonLotNo = "", GlassLotNo = "";
                string CarbonSpec = "", GlassSpec = "";

                using (conn = new SqlConnection(myConnectionString30))
                {
                    conn.Open();

                    //判別是否有做出貨檢驗，無出貨檢驗資料不允許包裝
                    selectCmd = "SELECT  * FROM  CH_ShippingInspection where SerialNo='" + CylinderNO + "'";
                    cmd = new SqlCommand(selectCmd, conn);
                    using (reader = cmd.ExecuteReader())
                    {
                        if (reader.Read())
                        {
                            ;
                        }
                        else
                        {
                            Error += "Code：114 無出貨檢驗資料，請聯繫品保\n";
                        }
                    }

                    //爆破
                    selectCmd = "SELECT  * FROM [PPT_Burst] WHERE [ManufacturingNo] = '" + LotNumber + "' and FinalResult='PASS' order by AcceptanceNo desc";
                    cmd = new SqlCommand(selectCmd, conn);
                    using (reader = cmd.ExecuteReader())
                    {
                        if (reader.HasRows)
                        {
                            ;
                        }
                        else
                        {
                            Error += "Code：113 無爆破資料，請聯繫品保\n";
                        }
                    }

                    //循環
                    selectCmd = "SELECT  * FROM [PPT_Cycling] WHERE [LotNo] = '" + LotNumber + "' and FinalResult='PASS'";
                    cmd = new SqlCommand(selectCmd, conn);
                    using (reader = cmd.ExecuteReader())
                    {
                        if (reader.HasRows)
                        {
                            ;
                        }
                        else
                        {
                            if (!ProductNo.Contains("-L-"))
                            {
                                Error += "Code：117 無循環資料，請聯繫品保\n";
                            }
                        }
                    }

                    selectCmd = "SELECT  ResinLotNo, CarbonLotNo, GlassLotNo, CarbonSpec, GlassSpec FROM [FilamentWinding] WHERE [LotNo] = '" + LotNumber + "' order by id desc";
                    cmd = new SqlCommand(selectCmd, conn);
                    using (reader = cmd.ExecuteReader())
                    {
                        if (reader.Read())
                        {
                            ResinLotNo = reader.GetString(0);
                            CarbonLotNo = reader.GetString(1);
                            GlassLotNo = reader.GetString(2);
                            CarbonSpec = reader.GetString(3);
                            GlassSpec = reader.GetString(4);
                        }
                    }
                }

                if (ResinLotNo != "" && CarbonLotNo != "" && GlassLotNo != "")
                {
                    using (conn = new SqlConnection(myConnectionString30))
                    {
                        conn.Open();

                        //碳纖
                        selectCmd = "SELECT * FROM [IQC] A, [Esign2] B WHERE A.[AcceptanceNo]=B.[AcceptanceNo] AND A.[Type] = '碳纖' AND A.[LotNo] LIKE '%" + CarbonLotNo + "%'";
                        cmd = new SqlCommand(selectCmd, conn);
                        using (reader = cmd.ExecuteReader())
                        {
                            if (reader.Read())
                            {
                                ;
                            }
                            else
                            {
                                Error += "Code：118 無碳纖檢驗資料，請聯繫品保\n";
                            }
                        }

                        //玻纖
                        selectCmd = "SELECT * FROM [PPT] A, [Esign2] B WHERE A.[AcceptanceNo]=B.[AcceptanceNo] AND A.[Type] = '玻纖' AND A.[LotNo] LIKE '%" + GlassLotNo + "%'";
                        cmd = new SqlCommand(selectCmd, conn);
                        using (reader = cmd.ExecuteReader())
                        {
                            if (reader.Read())
                            {
                                ;
                            }
                            else
                            {
                                Error += "Code：119 無玻纖檢驗資料，請聯繫品保\n";
                            }
                        }

                        //樹酯
                        selectCmd = "SELECT * FROM [PPT] A, [Esign2] B WHERE A.[AcceptanceNo]=B.[AcceptanceNo] AND A.[Type] = '樹脂' AND A.[LotNo] LIKE '%" + ResinLotNo + "%' and FiberType ='玻' and (FiberLotNo like '%" + GlassLotNo + "%' or FiberSpec like '%" + GlassSpec + "%')";//20180912品保系統檢驗組組長 說只要規格一樣沒有對應批號也可以。當初為CE0086有問題
                        cmd = new SqlCommand(selectCmd, conn);
                        using (reader = cmd.ExecuteReader())
                        {
                            if (reader.Read())
                            {
                                ;
                            }
                            else
                            {
                                Error += "Code：120 無樹脂檢驗資料，請聯繫品保\n";
                            }
                        }

                        selectCmd = "SELECT * FROM [PPT] A, [Esign2] B WHERE A.[AcceptanceNo]=B.[AcceptanceNo] AND A.[Type] = '樹脂' AND A.[LotNo] LIKE '%" + ResinLotNo + "%' and FiberType ='碳' and (FiberLotNo like '%" + CarbonLotNo + "%' or FiberSpec like '%" + CarbonSpec + "%')";//20180912品保系統檢驗組組長 說只要規格一樣沒有對應批號也可以。當初為CE0086有問題
                        cmd = new SqlCommand(selectCmd, conn);
                        using (reader = cmd.ExecuteReader())
                        {
                            if (reader.Read())
                            {
                                ;
                            }
                            else
                            {
                                Error += "Code：120 無樹脂檢驗資料，請聯繫品保\n";
                            }
                        }
                    }
                }

                //對應內膽  拉伸、爆破
                //找出對應內膽批號
                string BuildUp = "";

                using (conn = new SqlConnection(AMS21_ConnectionString))
                {
                    conn.Open();

                    selectCmd = "SELECT [LinerLotNo] FROM [AMS_DATA].[dbo].[ComCylinderNo]" +
                        " WHERE [CylinderNo] = @CylinderNo";
                    cmd = new SqlCommand(selectCmd, conn);
                    cmd.Parameters.AddWithValue("@CylinderNo", CylinderNO);
                    using (reader = cmd.ExecuteReader())
                    {
                        if (reader.Read())
                        {
                            BuildUp = reader.GetString(reader.GetOrdinal("LinerLotNo"));
                        }
                    }
                }

                if (BuildUp != "")
                {
                    using (conn = new SqlConnection(myConnectionString30))
                    {
                        conn.Open();

                        selectCmd = "SELECT  * FROM [PPT_Tensile]" +
                            " WHERE [ManufacturingNo] = @LotNo" +
                            " AND FinalResult = 'PASS' ";
                        cmd = new SqlCommand(selectCmd, conn);
                        cmd.Parameters.AddWithValue("@LotNo", BuildUp);
                        using (reader = cmd.ExecuteReader())
                        {
                            if (reader.HasRows)
                            {
                                ;
                            }
                            else
                            {
                                Error += "Code：115 無對應內膽(" + BuildUp + ")拉伸資料，請聯繫品保\n";
                            }
                        }

                        selectCmd = "SELECT  * FROM [PPT_Burst]" +
                            " WHERE [ManufacturingNo] = @LotNo" +
                            " AND [FinalResult] ='PASS' order by AcceptanceNo desc";
                        cmd = new SqlCommand(selectCmd, conn);
                        cmd.Parameters.AddWithValue("@LotNo", BuildUp);
                        using (reader = cmd.ExecuteReader())
                        {
                            if (reader.HasRows)
                            {
                                ;
                            }
                            else
                            {
                                Error += "Code：116 無對應內膽(" + BuildUp + ")爆破資料，請聯繫品保\n";
                            }
                        }
                    }
                }
            }

            //20200702 客戶序號檢查
            string CustomerCylinderNo = string.Empty;

            using (conn = new SqlConnection(myConnectionString))
            {
                conn.Open();

                selectCmd = "select isnull([CustomerCylinderNo],'N') CustomerCylinderNo from [MSNBody] where [CylinderNo] = @CylinderNo";
                cmd = new SqlCommand(selectCmd, conn);
                cmd.Parameters.Add("CylinderNo", SqlDbType.VarChar).Value = CylinderNO;
                using (reader = cmd.ExecuteReader())
                {
                    if (reader.Read())
                    {
                        CustomerCylinderNo = reader.GetString(reader.GetOrdinal("CustomerCylinderNo"));
                    }
                }

                if (CustomerCylinderNo != "N" && CustomerCylinderNo != "")
                {
                    selectCmd = "select count(ID) count from MSNBody where CustomerCylinderNo = @CustomerCylinderNo ";
                    cmd = new SqlCommand(selectCmd, conn);
                    cmd.Parameters.Add("@CustomerCylinderNo", SqlDbType.VarChar).Value = CustomerCylinderNo;
                    using (reader = cmd.ExecuteReader())
                    {
                        if (reader.Read())
                        {
                            if (reader.GetInt32(reader.GetOrdinal("count")) > 1)
                            {
                                Error += "Code：121 此客戶序號以重複";
                            }
                        }
                    }

                    using (conn1 = new SqlConnection(myConnectionString21_AMS_check))
                    {
                        conn1.Open();

                        selectCmd1 = "select ID from [CylinderNoCheck_Q] where CylinderNo = @CylinderNo ";
                        cmd1 = new SqlCommand(selectCmd1, conn1);
                        cmd1.Parameters.Add("@CylinderNo", SqlDbType.VarChar).Value = CylinderNO;
                        using (reader1 = cmd1.ExecuteReader())
                        {
                            if (reader1.HasRows)
                            {
                                ;
                            }
                            else
                            {
                                //Error += "Code：122 品保未確認客戶序號";
                            }
                        }

                        selectCmd1 = "select ID from [CylinderNoCheck_P] where CylinderNo = @CylinderNo ";
                        cmd1 = new SqlCommand(selectCmd1, conn1);
                        cmd1.Parameters.Add("@CylinderNo", SqlDbType.VarChar).Value = CylinderNO;
                        using (reader1 = cmd1.ExecuteReader())
                        {
                            if (reader1.HasRows)
                            {
                                ;
                            }
                            else
                            {
                                //Error += "Code：123 生產未確認客戶序號";
                            }
                        }
                    }
                }
            }

            if (Error.Any())
            {
                BottomTextBox.Text = "";
                MessageBox.Show(Error, "警告", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                return;
            }


            //20200617 新增客戶序號確認
            using (conn = new SqlConnection(myConnectionString))
            {
                conn.Open();

                selectCmd = "select isnull([CustomerCylinderNo],'N') CustomerCylinderNo from [MSNBody] where [CylinderNo] = @CylinderNo";
                cmd = new SqlCommand(selectCmd, conn);
                cmd.Parameters.Add("CylinderNo", SqlDbType.VarChar).Value = NoLMCylinderNOTextBox.Text;
                using (reader = cmd.ExecuteReader())
                {
                    if (reader.Read())
                    {
                        if (reader.GetString(reader.GetOrdinal("CustomerCylinderNo")) != "N" && reader.GetString(reader.GetOrdinal("CustomerCylinderNo")) != "")
                        {
                            DialogResult result = MessageBox.Show("請確認客戶序號：" + reader.GetString(reader.GetOrdinal("CustomerCylinderNo")), "確認", MessageBoxButtons.OKCancel, MessageBoxIcon.Warning);

                            if (result == DialogResult.Cancel)
                            {
                                return;
                            }
                        }
                    }
                }
            }

            using (conn = new SqlConnection(myConnectionString))
            {
                conn.Open();

                selectCmd = "select vchManufacturingNo from MSNBody where CylinderNo = @CylinderNo ";
                cmd = new SqlCommand(selectCmd, conn);
                cmd.Parameters.Add("@CylinderNo", SqlDbType.VarChar).Value = CylinderNO;
                using (reader = cmd.ExecuteReader())
                {
                    if (reader.Read())
                    {
                        if (LotNumber != reader.GetString(reader.GetOrdinal("vchManufacturingNo")))
                        {
                            MessageBox.Show("請聯繫MIS", "警告", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                            return;
                        }
                    }
                }
            }

            int InsertSB = 0, UpdateLP = 0, UpdateMsn = 0;

            using (TransactionScope scope = new TransactionScope())
            {
                using (conn = new SqlConnection(myConnectionString))
                {
                    conn.Open();

                    //雷刻掃描完確認瓶身瓶底相同後載入資料
                    selectCmd = "INSERT INTO [ShippingBody] ( ListDate, ProductName, CylinderNumbers, WhereBox, WhereSeat, vchUser, Time, LotNumber ) " +
                        "VALUES ( @ListDate, @ProductName, @CylinderNumbers, @WhereBox, @WhereSeat, @vchUser, @Time, @LotNumber )";
                    cmd = new SqlCommand(selectCmd, conn);

                    cmd.Parameters.Add("@ListDate", SqlDbType.VarChar).Value = ListDate_LB.SelectedItem;
                    cmd.Parameters.Add("@ProductName", SqlDbType.VarChar).Value = ProductName_CB.SelectedItem;
                    cmd.Parameters.Add("@CylinderNumbers", SqlDbType.VarChar).Value = NoLMCylinderNOTextBox.Text;
                    cmd.Parameters.Add("@WhereBox", SqlDbType.VarChar).Value = WhereBox_LB.SelectedItem;
                    cmd.Parameters.Add("@WhereSeat", SqlDbType.VarChar).Value = Convert.ToInt32(NowSeat) + 1;
                    cmd.Parameters.Add("@vchUser", SqlDbType.VarChar).Value = User_LB.Text.Remove(0, 7);
                    cmd.Parameters.Add("@Time", SqlDbType.VarChar).Value = DateTime.Now.ToLocalTime().ToString();
                    cmd.Parameters.Add("@LotNumber", SqlDbType.VarChar).Value = LotNumber;

                    InsertSB = cmd.ExecuteNonQuery();

                    //更新登出時間
                    selectCmd = "UPDATE [LoginPackage] SET  [LogoutTime]= '" + DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss") + "' , [IsUpdate]='0' WHERE [ID] = '" + toolStripStatusLabel1.Text + "'";
                    cmd = new SqlCommand(selectCmd, conn);
                    UpdateLP = InsertSB = cmd.ExecuteNonQuery();

                    selectCmd = "INSERT INTO [WorkTimePackage] ( CylinderNo, Operator, OperatorId, AddTime, Date, WorkType, ProcessNO ) " +
                        "VALUES ( @CylinderNo, @Operator, @OperatorId, @AddTime, @Date, @WorkType, @ProcessNO )";
                    cmd = new SqlCommand(selectCmd, conn);

                    cmd.Parameters.Add("@CylinderNo", SqlDbType.VarChar).Value = NoLMCylinderNOTextBox.Text;
                    cmd.Parameters.Add("@Operator", SqlDbType.VarChar).Value = User;
                    cmd.Parameters.Add("@OperatorId", SqlDbType.VarChar).Value = ID;
                    cmd.Parameters.Add("@AddTime", SqlDbType.VarChar).Value = DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss");
                    cmd.Parameters.Add("@Date", SqlDbType.VarChar).Value = DateTime.Now.ToString("yyyy-MM-dd");
                    cmd.Parameters.Add("@WorkType", SqlDbType.VarChar).Value = worktype;
                    cmd.Parameters.Add("@ProcessNO", SqlDbType.VarChar).Value = ProcessNo;

                    //cmd.ExecuteNonQuery();

                    selectCmd = "update [MSNBody] set [Package]= '1' where [CylinderNo]='" + NoLMCylinderNOTextBox.Text + "'";
                    cmd = new SqlCommand(selectCmd, conn);
                    UpdateMsn = cmd.ExecuteNonQuery();
                }

                if (InsertSB != 0 && UpdateLP != 0 && UpdateMsn != 0)
                {
                    scope.Complete();
                }
                else
                {
                    MessageBox.Show("新增失敗，請重新新增", "警告", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                }
            }

            time = 420;

            if (CustomerBarCodeCheckBox.Checked == true)
            {
                CustomerBarCode CBC = new CustomerBarCode();
                CBC.ProductName = ProductName_CB.SelectedItem.ToString();
                CBC.ListDate = ListDate_LB.SelectedItem.ToString();
                CBC.Boxs = WhereBox_LB.SelectedItem.ToString();
                CBC.Location = (Convert.ToInt32(NowSeat) + 1).ToString();
                CBC.ShowDialog();
            }

            if (WeightCheckBox.Checked == true && ComPortcomboBox.SelectedIndex != -1)
            {
                //20200903 扣底做重
                bool HaveBase = false;

                var v = (from p in SDT.AsEnumerable()
                         where p.Field<string>("vchBoxs").Contains(WhereBox_LB.SelectedItem.ToString())
                         select p.Field<string>("vchBoxs")).FirstOrDefault();

                if (v == null)
                {
                    HaveBase = false;
                }
                else
                {
                    HaveBase = true;
                }

                CylinderNoWeight CNW = new CylinderNoWeight();
                CNW.ComPort = ComPortcomboBox.SelectedItem.ToString();
                CNW.ListDate = ListDate_LB.SelectedItem.ToString();
                CNW.Boxs = WhereBox_LB.SelectedItem.ToString();
                CNW.CylinderNo = NoLMCylinderNOTextBox.Text.ToString();
                CNW.check = checkBox1.Checked.ToString();
                CNW.HaveBase = HaveBase;

                CNW.ShowDialog();
            }

            if (SecondPrintCheckBox.CheckState == CheckState.Checked)
            {
                //列印標籤貼紙
                MarkSecondPrintBarCode(NoLMCylinderNOTextBox.Text.ToString());
                OutputSecondPrintExcel();
                GC.Collect();
                SetProfileString(FirstPrinterComboBox.SelectedItem.ToString());
            }

            //自動跳下一箱 
            NextBoxs();

            //載入目前箱號
            WhereBox_LB.SelectedItem = GetNowBoxNo();

            //載入入箱狀況的圖片
            LoadPictrue();

            //載入dataGridView資料
            LoadSQLDate();
        }

        private void MarkSecondPrintBarCode(string SerialNo)
        {
            CreateBarCode(SerialNo);

            Bitmap bmp = new Bitmap(BarCodePictureBox.Width, BarCodePictureBox.Height);
            BarCodePictureBox.DrawToBitmap(bmp, new Rectangle(0, 0, BarCodePictureBox.Width, BarCodePictureBox.Height));
            bmp.Save(@"C:\SerialNoCode\" + SerialNo + ".png", ImageFormat.Png);
        }

        private void CreateBarCode(string BarCodeData)
        {
            Code128_Label MyCode = new Code128_Label();

            //條碼高度
            MyCode.Height = Convert.ToUInt16("112");
            MyCode.Width = Convert.ToUInt16("318");

            //可見號碼
            MyCode.ValueFont = new Font("細明體", 32, FontStyle.Bold);

            //產生條碼
            Image img = MyCode.GetCodeImage(BarCodeData, Code128_Label.Encode.Code128A);
            BarCodePictureBox.Width = img.Width;
            BarCodePictureBox.Image = img;

        }

        private void Create_DataMatrix(string SerialNo)
        {
            //如果資料匣不在自動新增
            if (!Directory.Exists(@"C:\SerialNoCode"))
            {
                Directory.CreateDirectory(@"C:\SerialNoCode");
            }

            string saveQRcode = @"C:\SerialNoCode\";

            string fileName = saveQRcode + SerialNo + ".png";
            DmtxImageEncoder encoder = new DmtxImageEncoder();
            DmtxImageEncoderOptions options = new DmtxImageEncoderOptions();

            options.ModuleSize = 8;//8
            options.MarginSize = 4;//4
            options.BackColor = Color.White;
            options.ForeColor = Color.Black;

            Bitmap encodedBitmap = encoder.EncodeImage(SerialNo);
            encodedBitmap.Save(fileName, ImageFormat.Png);
        }

        //EXCEL輸出
        private void OutputSecondPrintExcel()
        {
            //設定標籤貼紙之印表機
            SetProfileString(this.SecondPrinterComboBox.SelectedItem.ToString());

            Excel.Application oXL = new Excel.Application();
            Excel.Workbook oWB;
            Excel.Worksheet oSheet;

            string srcFileName = "";

            srcFileName = Application.StartupPath + @".\BarCode.xlsx";//EXCEL檔案路徑

            try
            {
                //產生一個Workbook物件，並加入Application//改成.open以及在()中輸入開啟位子
                oWB = oXL.Workbooks.Open(srcFileName, Type.Missing, Type.Missing, Type.Missing, Type.Missing,
                                        Type.Missing, Type.Missing, Type.Missing, Type.Missing,
                                        Type.Missing, Type.Missing, Type.Missing, Type.Missing,
                                        Type.Missing, Type.Missing);
            }
            catch
            {
                MessageBox.Show(@"找不到EXCEL檔案！", "Warning");
                return;
            }
            //設定工作表
            oSheet = (Excel.Worksheet)oWB.ActiveSheet;
            float PicLeft, PicTop, PicWidth, PicHeight;
            string PicturePath, PicLocation;

            //PicLocation = "A2";
            PicLocation = ((char)65).ToString() + 2.ToString();
            PicturePath = @"C:\SerialNoCode\" + NoLMCylinderNOTextBox.Text.ToString() + ".png";

            Excel.Worksheet xSheet = (Excel.Worksheet)oWB.Sheets[1];
            //xSheet.Cells[ (1 + 2 * (i / 8)),(i % 8) + 1] = SelectCylinderNoListBox.Items[i].ToString();
            Excel.Range m_objRange = xSheet.get_Range(PicLocation, Type.Missing);
            m_objRange.Select();

            PicLeft = Convert.ToSingle(m_objRange.Left);
            PicTop = Convert.ToSingle(m_objRange.Top);
            PicWidth = 63;
            PicHeight = 35;
            xSheet.Shapes.AddPicture(PicturePath, Microsoft.Office.Core.MsoTriState.msoFalse, Microsoft.Office.Core.MsoTriState.msoTrue, PicLeft + 3, PicTop + 7, PicWidth, PicHeight);

            //顯示EXCEL
            oXL.Visible = false;

            //列印EXCEL
            oWB.PrintOutEx(Missing.Value, Missing.Value, Missing.Value, Missing.Value, Missing.Value, Missing.Value, Missing.Value, Missing.Value, Missing.Value);

            oXL.DisplayAlerts = true;

            //關閉EXCEL
            oWB.Close(Type.Missing, Type.Missing, Type.Missing);

            //釋放EXCEL資源
            oXL = null;
            oWB = null;
            oSheet = null;
        }

        private void NextNumber()
        {
            char[] b = new char[12];
            StringReader sr = new StringReader(NoLMCylinderNOTextBox.Text);
            sr.Read(b, 0, 12);
            sr.Close();

            string Nbb = "";
            int AddNbb = 0;
            for (int i = 0; i <= NoLMCylinderNOTextBox.Text.Length; i++)
            {
                if (Convert.ToChar(b[i]) >= 48 && Convert.ToChar(b[i]) <= 57)
                {
                    Nbb += b[i];
                }
                else if (Convert.ToChar(b[i]) >= 65 && Convert.ToChar(b[i]) <= 90)
                {
                    Ebb += b[i];
                }
            }

            AddNbb = Convert.ToInt32(Nbb);
            AddNbb += 1;
            NoLMCylinderNOTextBox.Text = NoLMCylinderNOTextBox.Text.Replace(Nbb, Convert.ToString(AddNbb).PadLeft(Nbb.Length, '0'));
            // NoLMCylinderNOTextBox.Text = Ebb + TrialCarry(AddNbb);
        }

        private void NextBoxs()
        {
            //用來自動跳下一箱     

            string BoxsListBoxIndex = "";
            string NowSeat2 = "";

            //此處插入一個跳出式的視窗，詢問是否要列印

            using (conn = new SqlConnection(myConnectionString))
            {
                conn.Open();

                selectCmd = "SELECT WhereSeat FROM [ShippingBody] where [ListDate]='" + ListDate_LB.SelectedItem + "' and [ProductName]='" + ProductName_CB.SelectedItem + "' and [WhereBox]='" + WhereBox_LB.SelectedItem + "' order by Convert(INT,[WhereSeat]) DESC ";
                cmd = new SqlCommand(selectCmd, conn);
                using (reader = cmd.ExecuteReader())
                {
                    if (reader.Read())
                    {
                        NowSeat2 = reader.GetString(reader.GetOrdinal("WhereSeat"));
                        BoxsListBoxIndex = WhereBox_LB.SelectedIndex.ToString();

                        //如果箱號已經超過最大箱數則不自動跳箱
                        if ((Convert.ToInt32(BoxsListBoxIndex) >= (WhereBox_LB.Items.Count - 1)) && WhereBox_LB.Items.Count != 1 && NowSeat2 == Aboxof())
                        {
                            MessageBox.Show("此日期嘜頭已經完全結束", "提示");
                            return;
                        }

                        if (NowSeat2 == Aboxof())
                        {
                            if (PrintCheckBox.Checked == true)
                            {
                                PrintButton.PerformClick();
                            }
                            else
                            {
                                WhereBox_LB.SelectedIndex = (Convert.ToInt32(BoxsListBoxIndex) + 1);
                            }
                            WhereSeatLabel.Text = "1";
                        }
                    }
                }
            }
        }

        public string GetNowBoxNo()
        {
            return WhereBox_LB.SelectedItem.ToString();
        }

        private void MakeQRCode()
        {
            string QRcodDetail1 = "";

            QRcodDetail1 = QRcodDetailData();
            QRCodeWriter writer = new QRCodeWriter();
            ByteMatrix matrix1; // 這是放 2D code 的結果
            int size = 400;   // 指定最後產生的 2D code 影像大小 (pixel)
            // 進行 QRCode 的編碼工作
            //<關鍵片段>            
            if (QRcodDetail1.Trim() == "")
            {
                MessageBox.Show("無QR Code 資訊，請確認是否有包裝氣瓶或聯繫MIS建立該產品型號之打字資訊");
                return;
            }
            matrix1 = writer.encode(QRcodDetail1, BarcodeFormat.QR_CODE, size, size, null);

            // 把 2d code 畫出來
            Bitmap img1 = new Bitmap(size, size); // 建立 Bitmap 圖形物件

            Color Color1 = Color.FromArgb(0, 0, 0); // 設定 Bitmap 物件內每一個點的顏色格式為 RGB

            for (int y = 0; y < matrix1.Height; ++y)
            {
                for (int x = 0; x < matrix1.Width; ++x)
                {
                    Color pixelColor = img1.GetPixel(x, y);

                    if (matrix1.get_Renamed(x, y) == -1)
                    {
                        //不設定為白色，則將以透明呈現。某些圖片顯示軟體會以白色表示
                        ;//img1.SetPixel(x, y, Color.White);
                    }
                    else
                    {
                        img1.SetPixel(x, y, Color.Black);
                    } // end of update 2d barcode image
                }
            } // end of for-loop

            pictureBox1.Image = img1;

            //如果資料匣不在自動新增
            if (!Directory.Exists(@"C:\QRCode"))
            {
                Directory.CreateDirectory(@"C:\QRCode");
            }

            string saveQRcode = @"C:\QRCode\";

            pictureBox1.Image.Save(saveQRcode + QRcodeName() + ".png", System.Drawing.Imaging.ImageFormat.Png);//(路徑,內設定相關訊息)儲存圖片
        }

        private string QRcodeName()
        {
            string QRcodeName1 = "";

            using (conn = new SqlConnection(myConnectionString))
            {
                conn.Open();

                selectCmd = "SELECT ListDate, ProductName, vchBoxs FROM [ShippingHead] where [ListDate]='" + ListDate_LB.SelectedItem + "' and [ProductName]='" + ProductName_CB.Text + "' and [vchBoxs]='" + WhereBox_LB.SelectedItem + "' ";
                cmd = new SqlCommand(selectCmd, conn);
                using (reader = cmd.ExecuteReader())
                {
                    if (reader.Read())
                    {
                        QRcodeName1 = reader.GetString(reader.GetOrdinal("ListDate")) + reader.GetString(reader.GetOrdinal("ProductName")) + reader.GetString(reader.GetOrdinal("vchBoxs"));
                    }
                }
            }

            return QRcodeName1;
        }

        private string QRcodDetailData()
        {
            string QRcodDetail1 = ""; string Aboxof = "";
            string QRClient = "", QRProductName = "", PackingMarks = "";
            string DemandNo = string.Empty;
            // int section = 0;

            //找出客戶資訊
            using (conn = new SqlConnection(myConnectionString))
            {
                conn.Open();

                selectCmd = "SELECT isnull(Client,'') Client, isnull(ProductName,'') ProductName, isnull(PackingMarks,'') PackingMarks" +
                    ", vchAboxof, isnull(DemandNo,'') DemandNo FROM [ShippingHead] where [ListDate] = @ListDate AND [ProductName]= @ProductName" +
                    " AND [vchBoxs]= @vchBoxs";
                cmd = new SqlCommand(selectCmd, conn);
                cmd.Parameters.AddWithValue("@ListDate", ListDate_LB.SelectedItem);
                cmd.Parameters.AddWithValue("@ProductName", ProductName_CB.Text);
                cmd.Parameters.AddWithValue("@vchBoxs", WhereBox_LB.SelectedItem);
                using (reader = cmd.ExecuteReader())
                {
                    if (reader.Read())
                    {
                        QRClient = reader.GetValue(reader.GetOrdinal("Client")).ToString();
                        QRProductName = reader.GetValue(reader.GetOrdinal("ProductName")).ToString();
                        //找出外箱嘜頭貼紙是否有客製化需求PackingMarks
                        PackingMarks = reader.GetValue(reader.GetOrdinal("PackingMarks")).ToString();
                        Aboxof = reader.GetValue(reader.GetOrdinal("vchAboxof")).ToString();
                        DemandNo = reader.GetString(reader.GetOrdinal("DemandNo")).ToString();

                    }
                }
            }

            GetThisBoxMaxCount();
            bool HasSpecial = false;

            if ((QRClient.ToUpper().Trim().StartsWith("SGA") || QRClient.ToUpper().Trim().StartsWith("AIRTANKS")
                || PackingMarks.Trim().StartsWith("Scientific Gas Australia Pty Ltd")
                || PackingMarks.Trim().StartsWith("Airtanks Limited")) && Aboxof == "1")
            {
                string CylinderNO = "";
                //find SGA Marking //CustomerQRCode
                //2020/04/10 EMMY已經與客戶確認  只要單支裝 右手邊最大QRCode改為只顯示序號
                //找出序號再找出產品型號，找出Marking
                using (conn = new SqlConnection(myConnectionString))
                {
                    conn.Open();

                    selectCmd = "SELECT ISNULL([CylinderNumbers],'') FROM [ShippingBody]" +
                        " WHERE [ListDate] = @ListDate AND [ProductName] = @ProductName " +
                        "AND [WhereBox] = @WhereBox ORDER BY Convert(int,WhereSeat)";
                    cmd = new SqlCommand(selectCmd, conn);
                    cmd.Parameters.AddWithValue("@ListDate", ListDate_LB.SelectedItem);
                    cmd.Parameters.AddWithValue("@ProductName", ProductName_CB.Text);
                    cmd.Parameters.AddWithValue("@WhereBox", WhereBox_LB.SelectedItem);
                    using (reader = cmd.ExecuteReader())
                    {
                        if (reader.Read())
                        {
                            CylinderNO = reader.GetValue(0).ToString();
                        }
                    }
                    QRcodDetail1 = CylinderNO;

                    HasSpecial = true;
                }
            }

            if (!HasSpecial)
            {
                List<string> SerialNoArray = new List<string>();
                SerialNoArray.Clear();
                int Cumulative = 0;

                using (conn = new SqlConnection(myConnectionString))
                {
                    conn.Open();

                    selectCmd = "SELECT CylinderNumbers,[LotNumber] FROM [ShippingBody] where [ListDate]='" + ListDate_LB.SelectedItem + "' and [ProductName]='" + ProductName_CB.Text + "' and [WhereBox]='" + WhereBox_LB.SelectedItem + "' ORDER BY convert(int,[WhereSeat]) asc ";
                    cmd = new SqlCommand(selectCmd, conn);
                    using (reader = cmd.ExecuteReader())
                    {
                        while (reader.Read())
                        {
                            if (QRClient.Contains("Praxair") == true)
                            {//Praxair ->only CylinderNo
                                SerialNoArray.Add(reader.GetString(reader.GetOrdinal("CylinderNumbers")));
                            }
                            //1.20200821 AMS CC 試單 暫依此規則
                            else if (DemandNo == "2201-20200820001")
                            {
                                SerialNoArray.Add(reader.GetString(reader.GetOrdinal("LotNumber"))
                                    + " - " + reader.GetString(reader.GetOrdinal("CylinderNumbers")));
                            }
                            else
                            {//AMS Default data
                                SerialNoArray.Add((Cumulative + 1) + " " + reader.GetString(reader.GetOrdinal("CylinderNumbers")));
                            }
                            MarkSerialNoDataMatrix(reader.GetString(reader.GetOrdinal("CylinderNumbers")));
                            //MarkSerialNoBarCode(reader.GetString(3));

                            Cumulative++;
                        }
                    }

                    if (QRClient.Contains("Praxair") == false)
                    {
                        //AMS Default data
                        selectCmd = "SELECT isnull( CustomerProductName ,'') CustomerProductName,isnull([CustomerProductNo],'') CustomerProductNo FROM [ShippingHead] where [ListDate]='" + ListDate_LB.SelectedItem + "' and [ProductName]='" + ProductName_CB.Text + "' and [vchBoxs]='" + WhereBox_LB.SelectedItem + "' ";
                        cmd = new SqlCommand(selectCmd, conn);
                        using (reader = cmd.ExecuteReader())
                        {
                            if (reader.Read())
                            {
                                //1.20200821 AMS CC 試單 暫依此規則
                                if (DemandNo == "2201-20200820001")
                                {
                                    QRcodDetail1 = "Part Description:" + reader.GetString(reader.GetOrdinal("CustomerProductName"))
                                        + "\r\nPart No. " + reader.GetString(reader.GetOrdinal("CustomerProductNo"))
                                        + "\r\nQuantity: " + Getcount + " pieces\r\nC/NO. " + WhereBox_LB.SelectedItem
                                        + "\r\nBatch No./Serial No.\r\n";
                                }
                                else
                                {

                                    QRcodDetail1 = "Part Description:" + reader.GetString(reader.GetOrdinal("CustomerProductName")) + "\r\nPart No. " + reader.GetString(reader.GetOrdinal("CustomerProductNo")) + "\r\nQuantity: " + Getcount + " pieces\r\nC/NO. " + WhereBox_LB.SelectedItem + "\r\nSerial No.\r\n";
                                }

                            }
                        }
                    }
                }

                for (int i = 0; i < SerialNoArray.Count; i++)
                {
                    if (SerialNoArray[i] != null)
                    {
                        QRcodDetail1 = QRcodDetail1 + SerialNoArray[i];
                        if (QRClient.Contains("Praxair") == true)
                        {
                            QRcodDetail1 += "\r\n";
                        }
                        else
                        {
                            QRcodDetail1 += "\r\n";
                        }
                    }
                }
            }

            return QRcodDetail1;
        }

        //檢查跳箱的
        private void Match()
        {
            for (int i = 1; i < BoxsArray.Length; i++)
            {
                if (BoxsArray[i] == null)
                {
                    break;
                }

                if (Convert.ToInt32(BoxsCountArray[i + 1]) > Convert.ToInt32(BoxsCountArray[i]))
                {
                    JumpBoxLabel.Text = "跳箱箱號：" + BoxsArray[i];
                }
            }
        }

        private void GetThisBoxMaxCount()
        {
            using (conn = new SqlConnection(myConnectionString))
            {
                conn.Open();

                selectCmd = "SELECT count([WhereSeat]) FROM [ShippingBody] where [ListDate]='" + ListDate_LB.SelectedItem + "' and [ProductName]='" + ProductName_CB.Text + "' and [WhereBox]='" + WhereBox_LB.SelectedItem + "'";
                cmd = new SqlCommand(selectCmd, conn);
                using (reader = cmd.ExecuteReader())
                {
                    if (reader.Read())
                    {
                        Getcount = reader.GetInt32(0);
                    }
                }
            }
        }

        //一維條碼//
        public class Code128
        {
            private DataTable m_Code128 = new DataTable();

            private uint m_Height = 40;
            /// <summary>
            /// 高度
            /// </summary>
            public uint Height { get { return m_Height; } set { m_Height = value; } }

            private Font m_ValueFont = null;
            /// <summary>
            /// 是否顯示可見號碼 如果為Null不顯示號碼
            /// </summary>
            public Font ValueFont { get { return m_ValueFont; } set { m_ValueFont = value; } }

            private byte m_Magnify = 0;
            /// <summary>
            /// 放大倍數
            /// </summary>
            public byte Magnify { get { return m_Magnify; } set { m_Magnify = value; } }
            /// <summary>
            /// 條碼類別
            /// </summary>
            public enum Encode
            {
                Code128A,
                Code128B,
                Code128C,
                EAN128
            }

            public Code128()
            {
                m_Code128.Columns.Add("ID");
                m_Code128.Columns.Add("Code128A");
                m_Code128.Columns.Add("Code128B");
                m_Code128.Columns.Add("Code128C");
                m_Code128.Columns.Add("BandCode");

                m_Code128.CaseSensitive = true;

                #region 資料表
                m_Code128.Rows.Add("0", " ", " ", "00", "212222");
                m_Code128.Rows.Add("1", "!", "!", "01", "222122");
                m_Code128.Rows.Add("2", "/", "/", "02", "222221");
                m_Code128.Rows.Add("3", "#", "#", "03", "121223");
                m_Code128.Rows.Add("4", "$", "$", "04", "121322");
                m_Code128.Rows.Add("5", "%", "%", "05", "131222");
                m_Code128.Rows.Add("6", "&", "&", "06", "122213");
                m_Code128.Rows.Add("7", "'", "'", "07", "122312");
                m_Code128.Rows.Add("8", "(", "(", "08", "132212");
                m_Code128.Rows.Add("9", ")", ")", "09", "221213");
                m_Code128.Rows.Add("10", "*", "*", "10", "221312");
                m_Code128.Rows.Add("11", "+", "+", "11", "231212");
                m_Code128.Rows.Add("12", ",", ",", "12", "112232");
                m_Code128.Rows.Add("13", "-", "-", "13", "122132");
                m_Code128.Rows.Add("14", ".", ".", "14", "122231");
                m_Code128.Rows.Add("15", "/", "/", "15", "113222");
                m_Code128.Rows.Add("16", "0", "0", "16", "123122");
                m_Code128.Rows.Add("17", "1", "1", "17", "123221");
                m_Code128.Rows.Add("18", "2", "2", "18", "223211");
                m_Code128.Rows.Add("19", "3", "3", "19", "221132");
                m_Code128.Rows.Add("20", "4", "4", "20", "221231");
                m_Code128.Rows.Add("21", "5", "5", "21", "213212");
                m_Code128.Rows.Add("22", "6", "6", "22", "223112");
                m_Code128.Rows.Add("23", "7", "7", "23", "312131");
                m_Code128.Rows.Add("24", "8", "8", "24", "311222");
                m_Code128.Rows.Add("25", "9", "9", "25", "321122");
                m_Code128.Rows.Add("26", ":", ":", "26", "321221");
                m_Code128.Rows.Add("27", ";", ";", "27", "312212");
                m_Code128.Rows.Add("28", "<", "<", "28", "322112");
                m_Code128.Rows.Add("29", "=", "=", "29", "322211");
                m_Code128.Rows.Add("30", ">", ">", "30", "212123");
                m_Code128.Rows.Add("31", "?", "?", "31", "212321");
                m_Code128.Rows.Add("32", "@", "@", "32", "232121");
                m_Code128.Rows.Add("33", "A", "A", "33", "111323");
                m_Code128.Rows.Add("34", "B", "B", "34", "131123");
                m_Code128.Rows.Add("35", "C", "C", "35", "131321");
                m_Code128.Rows.Add("36", "D", "D", "36", "112313");
                m_Code128.Rows.Add("37", "E", "E", "37", "132113");
                m_Code128.Rows.Add("38", "F", "F", "38", "132311");
                m_Code128.Rows.Add("39", "G", "G", "39", "211313");
                m_Code128.Rows.Add("40", "H", "H", "40", "231113");
                m_Code128.Rows.Add("41", "I", "I", "41", "231311");
                m_Code128.Rows.Add("42", "J", "J", "42", "112133");
                m_Code128.Rows.Add("43", "K", "K", "43", "112331");
                m_Code128.Rows.Add("44", "L", "L", "44", "132131");
                m_Code128.Rows.Add("45", "M", "M", "45", "113123");
                m_Code128.Rows.Add("46", "N", "N", "46", "113321");
                m_Code128.Rows.Add("47", "O", "O", "47", "133121");
                m_Code128.Rows.Add("48", "P", "P", "48", "313121");
                m_Code128.Rows.Add("49", "Q", "Q", "49", "211331");
                m_Code128.Rows.Add("50", "R", "R", "50", "231131");
                m_Code128.Rows.Add("51", "S", "S", "51", "213113");
                m_Code128.Rows.Add("52", "T", "T", "52", "213311");
                m_Code128.Rows.Add("53", "U", "U", "53", "213131");
                m_Code128.Rows.Add("54", "V", "V", "54", "311123");
                m_Code128.Rows.Add("55", "W", "W", "55", "311321");
                m_Code128.Rows.Add("56", "X", "X", "56", "331121");
                m_Code128.Rows.Add("57", "Y", "Y", "57", "312113");
                m_Code128.Rows.Add("58", "Z", "Z", "58", "312311");
                m_Code128.Rows.Add("59", "[", "[", "59", "332111");
                m_Code128.Rows.Add("60", "//", "//", "60", "314111");
                m_Code128.Rows.Add("61", "]", "]", "61", "221411");
                m_Code128.Rows.Add("62", "^", "^", "62", "431111");
                m_Code128.Rows.Add("63", "_", "_", "63", "111224");
                m_Code128.Rows.Add("64", "NUL", "`", "64", "111422");
                m_Code128.Rows.Add("65", "SOH", "a", "65", "121124");
                m_Code128.Rows.Add("66", "STX", "b", "66", "121421");
                m_Code128.Rows.Add("67", "ETX", "c", "67", "141122");
                m_Code128.Rows.Add("68", "EOT", "d", "68", "141221");
                m_Code128.Rows.Add("69", "ENQ", "e", "69", "112214");
                m_Code128.Rows.Add("70", "ACK", "f", "70", "112412");
                m_Code128.Rows.Add("71", "BEL", "g", "71", "122114");
                m_Code128.Rows.Add("72", "BS", "h", "72", "122411");
                m_Code128.Rows.Add("73", "HT", "i", "73", "142112");
                m_Code128.Rows.Add("74", "LF", "j", "74", "142211");
                m_Code128.Rows.Add("75", "VT", "k", "75", "241211");
                m_Code128.Rows.Add("76", "FF", "I", "76", "221114");
                m_Code128.Rows.Add("77", "CR", "m", "77", "413111");
                m_Code128.Rows.Add("78", "SO", "n", "78", "241112");
                m_Code128.Rows.Add("79", "SI", "o", "79", "134111");
                m_Code128.Rows.Add("80", "DLE", "p", "80", "111242");
                m_Code128.Rows.Add("81", "DC1", "q", "81", "121142");
                m_Code128.Rows.Add("82", "DC2", "r", "82", "121241");
                m_Code128.Rows.Add("83", "DC3", "s", "83", "114212");
                m_Code128.Rows.Add("84", "DC4", "t", "84", "124112");
                m_Code128.Rows.Add("85", "NAK", "u", "85", "124211");
                m_Code128.Rows.Add("86", "SYN", "v", "86", "411212");
                m_Code128.Rows.Add("87", "ETB", "w", "87", "421112");
                m_Code128.Rows.Add("88", "CAN", "x", "88", "421211");
                m_Code128.Rows.Add("89", "EM", "y", "89", "212141");
                m_Code128.Rows.Add("90", "SUB", "z", "90", "214121");
                m_Code128.Rows.Add("91", "ESC", "{", "91", "412121");
                m_Code128.Rows.Add("92", "FS", "|", "92", "111143");
                m_Code128.Rows.Add("93", "GS", "}", "93", "111341");
                m_Code128.Rows.Add("94", "RS", "~", "94", "131141");
                m_Code128.Rows.Add("95", "US", "DEL", "95", "114113");
                m_Code128.Rows.Add("96", "FNC3", "FNC3", "96", "114311");
                m_Code128.Rows.Add("97", "FNC2", "FNC2", "97", "411113");
                m_Code128.Rows.Add("98", "SHIFT", "SHIFT", "98", "411311");
                m_Code128.Rows.Add("99", "CODEC", "CODEC", "99", "113141");
                m_Code128.Rows.Add("100", "CODEB", "FNC4", "CODEB", "114131");
                m_Code128.Rows.Add("101", "FNC4", "CODEA", "CODEA", "311141");
                m_Code128.Rows.Add("102", "FNC1", "FNC1", "FNC1", "411131");
                m_Code128.Rows.Add("103", "StartA", "StartA", "StartA", "211412");
                m_Code128.Rows.Add("104", "StartB", "StartB", "StartB", "211214");
                m_Code128.Rows.Add("105", "StartC", "StartC", "StartC", "211232");
                m_Code128.Rows.Add("106", "Stop", "Stop", "Stop", "2331112");
                #endregion
            }
            /// <summary>
            /// 獲取128圖形
            /// </summary>
            /// <param name="p_Text">文字</param>
            /// <param name="p_Code">編碼</param>
            /// <returns>圖形</returns>

            public Bitmap GetCodeImage(string p_Text, Encode p_Code)
            {
                string _ViewText = p_Text;
                string _Text = "";
                IList<int> _TextNumb = new List<int>();
                int _Examine = 0; //首位
                switch (p_Code)
                {
                    case Encode.Code128C:
                        _Examine = 105;
                        if (!((p_Text.Length & 1) == 0))
                        {
                            throw new Exception("128C長度必須是偶數");
                        }

                        while (p_Text.Length != 0)
                        {
                            int _Temp = 0;
                            try
                            {
                                int _CodeNumb128 = Int32.Parse(p_Text.Substring(0, 2));
                            }
                            catch
                            {
                                throw new Exception("128C必須是數位！");
                            }
                            _Text += GetValue(p_Code, p_Text.Substring(0, 2), ref _Temp);
                            _TextNumb.Add(_Temp);
                            p_Text = p_Text.Remove(0, 2);
                        }
                        break;
                    case Encode.EAN128:
                        _Examine = 105;
                        if (!((p_Text.Length & 1) == 0))
                        {
                            throw new Exception("EAN128長度必須是偶數");
                        }

                        _TextNumb.Add(102);
                        _Text += "411131";
                        while (p_Text.Length != 0)
                        {
                            int _Temp = 0;
                            try
                            {
                                int _CodeNumb128 = Int32.Parse(p_Text.Substring(0, 2));
                            }
                            catch
                            {
                                throw new Exception("128C必須是數位！");
                            }
                            _Text += GetValue(Encode.Code128C, p_Text.Substring(0, 2), ref _Temp);
                            _TextNumb.Add(_Temp);
                            p_Text = p_Text.Remove(0, 2);
                        }
                        break;
                    default:
                        if (p_Code == Encode.Code128A)
                        {
                            _Examine = 103;
                        }
                        else
                        {
                            _Examine = 104;
                        }

                        while (p_Text.Length != 0)
                        {
                            int _Temp = 0;
                            string _ValueCode = GetValue(p_Code, p_Text.Substring(0, 1), ref _Temp);
                            if (_ValueCode.Length == 0)
                            {
                                throw new Exception("不正確字元集!" + p_Text.Substring(0, 1).ToString());
                            }

                            _Text += _ValueCode;
                            _TextNumb.Add(_Temp);
                            p_Text = p_Text.Remove(0, 1);
                        }
                        break;
                }

                if (_TextNumb.Count == 0)
                {
                    throw new Exception("錯誤的編碼,無資料");
                }

                _Text = _Text.Insert(0, GetValue(_Examine)); //獲取開始位

                for (int i = 0; i != _TextNumb.Count; i++)
                {
                    _Examine += _TextNumb[i] * (i + 1);
                }
                _Examine = _Examine % 103; //獲得嚴效位
                _Text += GetValue(_Examine); //獲取嚴效位

                _Text += "2331112"; //結束位

                Bitmap _CodeImage = GetImage(_Text);
                //GetViewText(_CodeImage, _ViewText);
                return _CodeImage;
            }

            /// <summary>
            /// 獲取目標對應的資料
            /// </summary>
            /// <param name="p_Code">編碼</param>
            /// <param name="p_Value">數值 A b 30</param>
            /// <param name="p_SetID">返回編號</param>
            /// <returns>編碼</returns>
            private string GetValue(Encode p_Code, string p_Value, ref int p_SetID)
            {
                if (m_Code128 == null)
                {
                    return "";
                }

                DataRow[] _Row = m_Code128.Select(p_Code.ToString() + "='" + p_Value + "'");
                if (_Row.Length != 1)
                {
                    throw new Exception("錯誤的編碼" + p_Value.ToString());
                }

                p_SetID = Int32.Parse(_Row[0]["ID"].ToString());
                return _Row[0]["BandCode"].ToString();
            }
            /// <summary>
            /// 根據編號獲得條紋
            /// </summary>
            /// <param name="p_CodeId"></param>
            /// <returns></returns>

            private string GetValue(int p_CodeId)
            {
                DataRow[] _Row = m_Code128.Select("ID='" + p_CodeId.ToString() + "'");
                if (_Row.Length != 1)
                {
                    throw new Exception("驗效位的編碼錯誤" + p_CodeId.ToString());
                }

                return _Row[0]["BandCode"].ToString();
            }

            /// <summary>
            /// 獲得條碼圖形
            /// </summary>
            /// <param name="p_Text">文字</param>
            /// <returns>圖形</returns>

            private Bitmap GetImage(string p_Text)
            {
                char[] _Value = p_Text.ToCharArray();
                int _Width = 0;
                for (int i = 0; i != _Value.Length; i++)
                {
                    _Width += Int32.Parse(_Value[i].ToString()) * (m_Magnify + 1);
                }

                Bitmap _CodeImage = new Bitmap(_Width, (int)m_Height);
                Graphics _Garphics = Graphics.FromImage(_CodeImage);
                //Pen _Pen;
                int _LenEx = 0;
                for (int i = 0; i != _Value.Length; i++)
                {
                    int _ValueNumb = Int32.Parse(_Value[i].ToString()) * (m_Magnify + 1); //獲取寬和放大係數

                    if (!((i & 1) == 0))
                    {
                        //_Pen = new Pen(Brushes.White, _ValueNumb);
                        _Garphics.FillRectangle(Brushes.White, new Rectangle(_LenEx, 0, _ValueNumb, (int)m_Height));
                    }
                    else
                    {
                        //_Pen = new Pen(Brushes.Black, _ValueNumb);
                        _Garphics.FillRectangle(Brushes.Black, new Rectangle(_LenEx, 0, _ValueNumb, (int)m_Height));
                    }
                    //_Garphics.(_Pen, new Point(_LenEx, 0), new Point(_LenEx, m_Height));
                    _LenEx += _ValueNumb;
                }
                _Garphics.Dispose();
                return _CodeImage;
            }
        }

        public class Code128_Label
        {
            private DataTable m_Code128 = new DataTable();

            private uint m_Height = 100;
            /// <summary>
            /// 高度
            /// </summary>
            public uint Height { get { return m_Height; } set { m_Height = value; } }

            private uint m_Width = 274;
            /// <summary>
            /// 高度
            /// </summary>
            public uint Width { get { return m_Width; } set { m_Width = value; } }

            private Font m_ValueFont = null;
            /// <summary>
            /// 是否顯示可見號碼 如果為Null不顯示號碼
            /// </summary>
            public Font ValueFont { get { return m_ValueFont; } set { m_ValueFont = value; } }

            private byte m_Magnify = 0;
            private float f_Magnify = 0;
            /// <summary>
            /// 放大倍數
            /// </summary>
            public byte Magnify { get { return m_Magnify; } set { m_Magnify = value; } }
            /// <summary>
            /// 條碼類別
            /// </summary>
            public enum Encode
            {
                Code128A,
                Code128B,
                Code128C,
                EAN128
            }

            public Code128_Label()
            {
                m_Code128.Columns.Add("ID");
                m_Code128.Columns.Add("Code128A");
                m_Code128.Columns.Add("Code128B");
                m_Code128.Columns.Add("Code128C");
                m_Code128.Columns.Add("BandCode");

                m_Code128.CaseSensitive = true;

                #region 資料表
                m_Code128.Rows.Add("0", " ", " ", "00", "212222");
                m_Code128.Rows.Add("1", "!", "!", "01", "222122");
                m_Code128.Rows.Add("2", "/", "/", "02", "222221");
                m_Code128.Rows.Add("3", "#", "#", "03", "121223");
                m_Code128.Rows.Add("4", "$", "$", "04", "121322");
                m_Code128.Rows.Add("5", "%", "%", "05", "131222");
                m_Code128.Rows.Add("6", "&", "&", "06", "122213");
                m_Code128.Rows.Add("7", "'", "'", "07", "122312");
                m_Code128.Rows.Add("8", "(", "(", "08", "132212");
                m_Code128.Rows.Add("9", ")", ")", "09", "221213");
                m_Code128.Rows.Add("10", "*", "*", "10", "221312");
                m_Code128.Rows.Add("11", "+", "+", "11", "231212");
                m_Code128.Rows.Add("12", ",", ",", "12", "112232");
                m_Code128.Rows.Add("13", "-", "-", "13", "122132");
                m_Code128.Rows.Add("14", ".", ".", "14", "122231");
                m_Code128.Rows.Add("15", "/", "/", "15", "113222");
                m_Code128.Rows.Add("16", "0", "0", "16", "123122");
                m_Code128.Rows.Add("17", "1", "1", "17", "123221");
                m_Code128.Rows.Add("18", "2", "2", "18", "223211");
                m_Code128.Rows.Add("19", "3", "3", "19", "221132");
                m_Code128.Rows.Add("20", "4", "4", "20", "221231");
                m_Code128.Rows.Add("21", "5", "5", "21", "213212");
                m_Code128.Rows.Add("22", "6", "6", "22", "223112");
                m_Code128.Rows.Add("23", "7", "7", "23", "312131");
                m_Code128.Rows.Add("24", "8", "8", "24", "311222");
                m_Code128.Rows.Add("25", "9", "9", "25", "321122");
                m_Code128.Rows.Add("26", ":", ":", "26", "321221");
                m_Code128.Rows.Add("27", ";", ";", "27", "312212");
                m_Code128.Rows.Add("28", "<", "<", "28", "322112");
                m_Code128.Rows.Add("29", "=", "=", "29", "322211");
                m_Code128.Rows.Add("30", ">", ">", "30", "212123");
                m_Code128.Rows.Add("31", "?", "?", "31", "212321");
                m_Code128.Rows.Add("32", "@", "@", "32", "232121");
                m_Code128.Rows.Add("33", "A", "A", "33", "111323");
                m_Code128.Rows.Add("34", "B", "B", "34", "131123");
                m_Code128.Rows.Add("35", "C", "C", "35", "131321");
                m_Code128.Rows.Add("36", "D", "D", "36", "112313");
                m_Code128.Rows.Add("37", "E", "E", "37", "132113");
                m_Code128.Rows.Add("38", "F", "F", "38", "132311");
                m_Code128.Rows.Add("39", "G", "G", "39", "211313");
                m_Code128.Rows.Add("40", "H", "H", "40", "231113");
                m_Code128.Rows.Add("41", "I", "I", "41", "231311");
                m_Code128.Rows.Add("42", "J", "J", "42", "112133");
                m_Code128.Rows.Add("43", "K", "K", "43", "112331");
                m_Code128.Rows.Add("44", "L", "L", "44", "132131");
                m_Code128.Rows.Add("45", "M", "M", "45", "113123");
                m_Code128.Rows.Add("46", "N", "N", "46", "113321");
                m_Code128.Rows.Add("47", "O", "O", "47", "133121");
                m_Code128.Rows.Add("48", "P", "P", "48", "313121");
                m_Code128.Rows.Add("49", "Q", "Q", "49", "211331");
                m_Code128.Rows.Add("50", "R", "R", "50", "231131");
                m_Code128.Rows.Add("51", "S", "S", "51", "213113");
                m_Code128.Rows.Add("52", "T", "T", "52", "213311");
                m_Code128.Rows.Add("53", "U", "U", "53", "213131");
                m_Code128.Rows.Add("54", "V", "V", "54", "311123");
                m_Code128.Rows.Add("55", "W", "W", "55", "311321");
                m_Code128.Rows.Add("56", "X", "X", "56", "331121");
                m_Code128.Rows.Add("57", "Y", "Y", "57", "312113");
                m_Code128.Rows.Add("58", "Z", "Z", "58", "312311");
                m_Code128.Rows.Add("59", "[", "[", "59", "332111");
                m_Code128.Rows.Add("60", "//", "//", "60", "314111");
                m_Code128.Rows.Add("61", "]", "]", "61", "221411");
                m_Code128.Rows.Add("62", "^", "^", "62", "431111");
                m_Code128.Rows.Add("63", "_", "_", "63", "111224");
                m_Code128.Rows.Add("64", "NUL", "`", "64", "111422");
                m_Code128.Rows.Add("65", "SOH", "a", "65", "121124");
                m_Code128.Rows.Add("66", "STX", "b", "66", "121421");
                m_Code128.Rows.Add("67", "ETX", "c", "67", "141122");
                m_Code128.Rows.Add("68", "EOT", "d", "68", "141221");
                m_Code128.Rows.Add("69", "ENQ", "e", "69", "112214");
                m_Code128.Rows.Add("70", "ACK", "f", "70", "112412");
                m_Code128.Rows.Add("71", "BEL", "g", "71", "122114");
                m_Code128.Rows.Add("72", "BS", "h", "72", "122411");
                m_Code128.Rows.Add("73", "HT", "i", "73", "142112");
                m_Code128.Rows.Add("74", "LF", "j", "74", "142211");
                m_Code128.Rows.Add("75", "VT", "k", "75", "241211");
                m_Code128.Rows.Add("76", "FF", "I", "76", "221114");
                m_Code128.Rows.Add("77", "CR", "m", "77", "413111");
                m_Code128.Rows.Add("78", "SO", "n", "78", "241112");
                m_Code128.Rows.Add("79", "SI", "o", "79", "134111");
                m_Code128.Rows.Add("80", "DLE", "p", "80", "111242");
                m_Code128.Rows.Add("81", "DC1", "q", "81", "121142");
                m_Code128.Rows.Add("82", "DC2", "r", "82", "121241");
                m_Code128.Rows.Add("83", "DC3", "s", "83", "114212");
                m_Code128.Rows.Add("84", "DC4", "t", "84", "124112");
                m_Code128.Rows.Add("85", "NAK", "u", "85", "124211");
                m_Code128.Rows.Add("86", "SYN", "v", "86", "411212");
                m_Code128.Rows.Add("87", "ETB", "w", "87", "421112");
                m_Code128.Rows.Add("88", "CAN", "x", "88", "421211");
                m_Code128.Rows.Add("89", "EM", "y", "89", "212141");
                m_Code128.Rows.Add("90", "SUB", "z", "90", "214121");
                m_Code128.Rows.Add("91", "ESC", "{", "91", "412121");
                m_Code128.Rows.Add("92", "FS", "|", "92", "111143");
                m_Code128.Rows.Add("93", "GS", "}", "93", "111341");
                m_Code128.Rows.Add("94", "RS", "~", "94", "131141");
                m_Code128.Rows.Add("95", "US", "DEL", "95", "114113");
                m_Code128.Rows.Add("96", "FNC3", "FNC3", "96", "114311");
                m_Code128.Rows.Add("97", "FNC2", "FNC2", "97", "411113");
                m_Code128.Rows.Add("98", "SHIFT", "SHIFT", "98", "411311");
                m_Code128.Rows.Add("99", "CODEC", "CODEC", "99", "113141");
                m_Code128.Rows.Add("100", "CODEB", "FNC4", "CODEB", "114131");
                m_Code128.Rows.Add("101", "FNC4", "CODEA", "CODEA", "311141");
                m_Code128.Rows.Add("102", "FNC1", "FNC1", "FNC1", "411131");
                m_Code128.Rows.Add("103", "StartA", "StartA", "StartA", "211412");
                m_Code128.Rows.Add("104", "StartB", "StartB", "StartB", "211214");
                m_Code128.Rows.Add("105", "StartC", "StartC", "StartC", "211232");
                m_Code128.Rows.Add("106", "Stop", "Stop", "Stop", "2331112");
                #endregion
            }

            public Bitmap GetCodeImage(string p_Text, Encode p_Code)
            {
                string _ViewText = p_Text;
                string _Text = "";
                IList<int> _TextNumb = new List<int>();
                int _Examine = 0; //首位
                switch (p_Code)
                {
                    case Encode.Code128C:
                        _Examine = 105;
                        if (!((p_Text.Length & 1) == 0))
                        {
                            throw new Exception("128C長度必須是偶數");
                        }

                        while (p_Text.Length != 0)
                        {
                            int _Temp = 0;
                            try
                            {
                                int _CodeNumb128 = Int32.Parse(p_Text.Substring(0, 2));
                            }
                            catch
                            {
                                throw new Exception("128C必須是數位！");
                            }
                            _Text += GetValue(p_Code, p_Text.Substring(0, 2), ref _Temp);
                            _TextNumb.Add(_Temp);
                            p_Text = p_Text.Remove(0, 2);
                        }
                        break;
                    case Encode.EAN128:
                        _Examine = 105;
                        if (!((p_Text.Length & 1) == 0))
                        {
                            throw new Exception("EAN128長度必須是偶數");
                        }

                        _TextNumb.Add(102);
                        _Text += "411131";
                        while (p_Text.Length != 0)
                        {
                            int _Temp = 0;
                            try
                            {
                                int _CodeNumb128 = Int32.Parse(p_Text.Substring(0, 2));
                            }
                            catch
                            {
                                throw new Exception("128C必須是數位！");
                            }
                            _Text += GetValue(Encode.Code128C, p_Text.Substring(0, 2), ref _Temp);
                            _TextNumb.Add(_Temp);
                            p_Text = p_Text.Remove(0, 2);
                        }
                        break;
                    default:
                        if (p_Code == Encode.Code128A)
                        {
                            _Examine = 103;
                        }
                        else
                        {
                            _Examine = 104;
                        }

                        while (p_Text.Length != 0)
                        {
                            int _Temp = 0;
                            string _ValueCode = GetValue(p_Code, p_Text.Substring(0, 1), ref _Temp);
                            if (_ValueCode.Length == 0)
                            {
                                throw new Exception("不正確字元集!" + p_Text.Substring(0, 1).ToString());
                            }

                            _Text += _ValueCode;
                            _TextNumb.Add(_Temp);
                            p_Text = p_Text.Remove(0, 1);
                        }
                        break;
                }

                if (_TextNumb.Count == 0)
                {
                    throw new Exception("錯誤的編碼,無資料");
                }

                _Text = _Text.Insert(0, GetValue(_Examine)); //獲取開始位

                for (int i = 0; i != _TextNumb.Count; i++)
                {
                    _Examine += _TextNumb[i] * (i + 1);
                }
                _Examine = _Examine % 103; //獲得嚴效位
                _Text += GetValue(_Examine); //獲取嚴效位

                _Text += "2331112"; //結束位

                Bitmap _CodeImage = GetImage(_Text);
                GetViewText(_CodeImage, _ViewText);
                return _CodeImage;
            }

            private string GetValue(Encode p_Code, string p_Value, ref int p_SetID)
            {
                if (m_Code128 == null)
                {
                    return "";
                }

                DataRow[] _Row = m_Code128.Select(p_Code.ToString() + "='" + p_Value + "'");
                if (_Row.Length != 1)
                {
                    throw new Exception("錯誤的編碼" + p_Value.ToString());
                }

                p_SetID = Int32.Parse(_Row[0]["ID"].ToString());
                return _Row[0]["BandCode"].ToString();
            }

            private string GetValue(int p_CodeId)
            {
                DataRow[] _Row = m_Code128.Select("ID='" + p_CodeId.ToString() + "'");
                if (_Row.Length != 1)
                {
                    throw new Exception("驗效位的編碼錯誤" + p_CodeId.ToString());
                }

                return _Row[0]["BandCode"].ToString();
            }

            private Bitmap GetImage(string p_Text)
            {
                char[] _Value = p_Text.ToCharArray();
                int _LenEx = 0;
                int Magnify = 0;
                for (int i = 0; i != _Value.Length; i++)
                {
                    Magnify += Int32.Parse(_Value[i].ToString());
                }
                m_Magnify = 3;// (byte)(m_Width / Magnify);
                //f_Magnify = (float)((float)m_Width / (float)Magnify);
                m_Width = (uint)(Magnify * 3);// (uint)(m_Magnify * Magnify);
                //int _Width = 0;
                //for (int i = 0; i != _Value.Length; i++)
                //{
                //    _Width += Int32.Parse(_Value[i].ToString()) * (m_Magnify + 1);
                //}
                //Bitmap _CodeImage = new Bitmap(_Width, (int)m_Height);
                Bitmap _CodeImage = new Bitmap((int)m_Width, (int)m_Height);
                Graphics _Garphics = Graphics.FromImage(_CodeImage);
                //Pen _Pen;
                //int _LenEx = 0;
                //int Magnify = 0;
                //for (int i = 0; i != _Value.Length; i++)
                //{
                //    Magnify += Int32.Parse(_Value[i].ToString());
                //}
                //m_Magnify =(byte)( m_Width / Magnify);
                //f_Magnify = (float)((float)m_Width / (float)Magnify);
                //m_Width = (uint)(m_Magnify * Magnify);
                for (int i = 0; i != _Value.Length; i++)
                {
                    int _ValueNumb = (Int32.Parse(_Value[i].ToString()) * m_Magnify); //獲取寬和放大係數
                    // int _ValueNumb = (int)(Int32.Parse(_Value[i].ToString()) * f_Magnify); //獲取寬和放大係數
                    // int _ValueNumb = Int32.Parse(m_Width.ToString());

                    if (!((i & 1) == 0))
                    {
                        //_Pen = new Pen(Brushes.White, _ValueNumb);
                        _Garphics.FillRectangle(Brushes.White, new Rectangle(_LenEx, 0, _ValueNumb, (int)m_Height));
                    }
                    else
                    {
                        //_Pen = new Pen(Brushes.Black, _ValueNumb);
                        _Garphics.FillRectangle(Brushes.Black, new Rectangle(_LenEx, 0, _ValueNumb, (int)m_Height));
                    }
                    //_Garphics.(_Pen, new Point(_LenEx, 0), new Point(_LenEx, m_Height));
                    _LenEx += _ValueNumb;
                }

                _Garphics.Dispose();
                return _CodeImage;
            }

            private void GetViewText(Bitmap p_Bitmap, string p_ViewText)
            {
                if (m_ValueFont == null)
                {
                    return;
                }

                Graphics _Graphics = Graphics.FromImage(p_Bitmap);
                SizeF _DrawSize = _Graphics.MeasureString(p_ViewText, m_ValueFont);
                if (_DrawSize.Height > p_Bitmap.Height - 10 || _DrawSize.Width > p_Bitmap.Width)
                {
                    _Graphics.Dispose();
                    return;
                }

                int _StarY = p_Bitmap.Height - (int)_DrawSize.Height;

                _Graphics.FillRectangle(Brushes.White, new Rectangle(0, _StarY, p_Bitmap.Width, (int)_DrawSize.Height));
                StringFormat sf = new StringFormat();
                sf.Alignment = StringAlignment.Center;
                sf.LineAlignment = StringAlignment.Center;
                _Graphics.DrawString(p_ViewText, m_ValueFont, Brushes.Black, new RectangleF(0, _StarY, p_Bitmap.Width, (int)_DrawSize.Height), sf);
                //_Graphics.DrawString(p_ViewText, m_ValueFont, Brushes.Black, new RectangleF(0, _StarY, 274, (int)_DrawSize.Height), sf);
                // _Graphics.DrawString(p_ViewText, m_ValueFont, Brushes.Black, 0, _StarY,sf);
            }
        }

        private void SelectListBoxTextBox_KeyDown(object sender, KeyEventArgs e)
        {
            if (e.KeyCode == Keys.Enter)
            {
                //找出箱號
                if (WhereBox_LB.Items.Count != 0)
                {
                    for (int BoxListIndex = 0; BoxListIndex < WhereBox_LB.Items.Count; BoxListIndex++)
                    {
                        if (WhereBox_LB.Items[BoxListIndex].ToString().CompareTo(SelectListBoxTextBox.Text.ToString()) == 0)
                        {
                            WhereBox_LB.SelectedIndex = BoxListIndex;
                            break;
                        }
                    }
                }
            }
        }

        private void MarkSerialNoBarCode(string SerialNo)
        {
            Code128 MyCode = new Code128();

            //條碼高度
            MyCode.Height = 50;

            //可見號碼
            MyCode.ValueFont = new Font("細明體", 12, FontStyle.Regular);

            //產生條碼
            Image img = MyCode.GetCodeImage(SerialNo, Code128.Encode.Code128A);

            pictureBox1.Image = img;

            //如果資料匣不在自動新增
            if (!Directory.Exists(@"C:\SerialNoCode"))
            {
                Directory.CreateDirectory(@"C:\SerialNoCode");
            }

            string saveQRcode = @"C:\SerialNoCode\";

            pictureBox1.Image.Save(saveQRcode + SerialNo + ".png");
        }

        private void MarkSerialNoDataMatrix(string SerialNo)
        {
            Create_DataMatrix(SerialNo);
        }

        private void TodayDataButton_Click(object sender, EventArgs e)
        {
            OutputTodayDaaExcel();
            GC.Collect();
        }

        private void OutputTodayDaaExcel()
        {
            List<string> CylinderNumbersList = new List<string>();
            List<int> WhereBoxList = new List<int>();
            List<string> WhereSeatList = new List<string>();
            List<string> CustomerBarCodeList = new List<string>();
            List<string> ManufacturingNoList = new List<string>();
            List<string> CylinderWeightList = new List<string>();
            CylinderNumbersList.Clear();
            WhereBoxList.Clear();
            WhereSeatList.Clear();
            CustomerBarCodeList.Clear();
            ManufacturingNoList.Clear();
            CylinderWeightList.Clear();

            //載入[ShippingHead]的ListDate
            using (conn = new SqlConnection(myConnectionString))
            {
                conn.Open();

                selectCmd = "SELECT  CylinderNumbers, WhereBox, WhereSeat,ISNULL(CustomerBarCode,'') CustomerBarCode, ISNULL(CylinderWeight,'0') CylinderWeight FROM [ShippingBody]  where  [ListDate]='" + ListDate_LB.SelectedItem.ToString() + "' and [ProductName]='" + ProductName_CB.SelectedItem.ToString() + "' and CONVERT(datetime, SUBSTRING(Time, 0, 11), 111)>='" + DateTime.Now.ToLocalTime().ToString().Split(' ')[0].ToString() + "' and CONVERT(datetime, SUBSTRING(Time, 0, 11), 111)<='" + DateTime.Now.AddDays(1).ToLocalTime().ToString().Split(' ')[0].ToString() + "' ORDER BY RIGHT(REPLICATE('0', 8) + CAST(SUBSTRING(CylinderNumbers, 3, Len(CylinderNumbers)-2) AS NVARCHAR), 8)";
                cmd = new SqlCommand(selectCmd, conn);
                using (reader = cmd.ExecuteReader())
                {
                    while (reader.Read())
                    {
                        CylinderNumbersList.Add(reader.GetString(reader.GetOrdinal("CylinderNumbers")));
                        WhereBoxList.Add(Convert.ToInt32(reader.GetString(reader.GetOrdinal("WhereBox"))));
                        WhereSeatList.Add(reader.GetString(reader.GetOrdinal("WhereSeat")));
                        CustomerBarCodeList.Add(reader.GetString(reader.GetOrdinal("CustomerBarCode")));
                        CylinderWeightList.Add(reader.GetValue(reader.GetOrdinal("CylinderWeight")).ToString());
                    }
                }
            }

            if (CylinderNumbersList.Count == 0)
            {
                MessageBox.Show("無產品名稱:" + ProductName_CB.SelectedItem.ToString() + "、嘜頭日期:" + ListDate_LB.SelectedItem.ToString() + "於今天包裝之資料。");
                return;
            }

            //show excel
            Excel.Application oXL = new Excel.Application();
            Excel.Workbook oWB;
            Excel.Worksheet oSheet, oSheet2;
            string srcFileName = Application.StartupPath + @".\TodayPackageData.xlsx";//EXCEL檔案路徑

            try
            {
                //產生一個Workbook物件，並加入Application//改成.open以及在()中輸入開啟位子
                oWB = oXL.Workbooks.Open(srcFileName, Type.Missing, Type.Missing, Type.Missing, Type.Missing,
                                        Type.Missing, Type.Missing, Type.Missing, Type.Missing,
                                        Type.Missing, Type.Missing, Type.Missing, Type.Missing,
                                        Type.Missing, Type.Missing);
            }
            catch
            {
                MessageBox.Show(@"找不到TodayPackageData.xlsx檔案！", "Warning");
                return;
            }
            oXL.Visible = false;

            //設定工作表
            oSheet = (Excel.Worksheet)oWB.Worksheets.get_Item("批號資訊");

            oSheet2 = (Excel.Worksheet)oWB.Worksheets.get_Item("詳細資訊");
            oSheet.Cells[1, 2] = ProductName_CB.SelectedItem.ToString();
            oSheet.Cells[2, 2] = ListDate_LB.SelectedItem.ToString();
            oSheet.Cells[3, 2] = DateTime.Now.ToString("yyyy/MM/dd");

            oSheet2.Cells[1, 2] = ProductName_CB.SelectedItem.ToString();
            oSheet2.Cells[2, 2] = ListDate_LB.SelectedItem.ToString();
            oSheet2.Cells[3, 2] = DateTime.Now.ToString("yyyy/MM/dd");

            for (int i = 0; i < CylinderNumbersList.Count; i++)
            {
                oSheet2.Cells[7 + i, 1] = CylinderNumbersList[i].ToString();
                // oSheet2.Cells[7 + i, 2] = ManufacturingNoList[i].ToString();
                oSheet2.Cells[7 + i, 3] = WhereBoxList[i].ToString();
                oSheet2.Cells[7 + i, 4] = WhereSeatList[i].ToString();
                oSheet2.Cells[7 + i, 5] = CustomerBarCodeList[i].ToString();
                oSheet2.Cells[7 + i, 6] = CylinderWeightList[i].ToString();
            }

            WhereBoxList.Sort();
            oSheet.Cells[4, 2] = WhereBoxList[0].ToString() + "~" + WhereBoxList[CylinderNumbersList.Count - 1].ToString();
            oSheet2.Cells[4, 2] = WhereBoxList[0].ToString() + "~" + WhereBoxList[CylinderNumbersList.Count - 1].ToString();

            Excel.Sheets excelSheets = oWB.Worksheets;

            oXL.Visible = true;

            oXL = null;
            oWB = null;
            oSheet = null;

            GC.Collect();
        }

        private void WeightCheckBox_CheckedChanged(object sender, EventArgs e)
        {
            if (WeightCheckBox.Checked == true)
            {
                string[] ports = System.IO.Ports.SerialPort.GetPortNames();
                List<string> listPorts = new List<string>(ports);
                Comparison<string> comparer = delegate (string name1, string name2)
                {
                    int port1 = Convert.ToInt32(name1.Remove(0, 3));
                    int port2 = Convert.ToInt32(name2.Remove(0, 3));
                    return (port1 - port2);
                };

                listPorts.Sort(comparer);
                this.ComPortcomboBox.Items.AddRange(listPorts.ToArray());
                this.ComPortcomboBox.SelectedIndex = this.ComPortcomboBox.Items.Count - 1;
                ComPortcomboBox.Enabled = true;
            }
            else
            {
                ComPortcomboBox.Items.Clear();
                ComPortcomboBox.Enabled = false;
            }
        }

        private void ComplexQRCodeCheckBox_CheckedChanged(object sender, EventArgs e)
        {
            if (ComplexQRCodeCheckBox.CheckState == CheckState.Checked)
            {
                NoLMCylinderNOTextBox.MaxLength = 200;
                //NoLMCylinderNOTextBox.Multiline = true; //一般序號都在第一行故不用此
                //NoLMCylinderNOTextBox.Size =new System.Drawing.Size(301, 55);
            }
            else
            {
                NoLMCylinderNOTextBox.MaxLength = 10;
                //NoLMCylinderNOTextBox.Multiline = false;
            }
        }

        private void SecondPrintCheckBox_CheckedChanged(object sender, EventArgs e)
        {
            if (SecondPrintCheckBox.CheckState == CheckState.Checked)
            {
                SecondPrinterComboBox.Enabled = true;
            }
            else
            {
                SecondPrinterComboBox.Enabled = false;
            }
        }

        private void ListDate_LB_SelectedIndexChanged(object sender, EventArgs e)
        {
            BoxRangeLabel.Text = "";
            WhereBox_LB.Items.Clear();

            int BoxMax = 0, BoxMin = 0;

            //查詢箱號最小值
            //20190212
            using (conn = new SqlConnection(myConnectionString))
            {
                conn.Open();

                selectCmd = "SELECT [vchBoxs] FROM [ShippingHead] where [ListDate]='" + ListDate_LB.SelectedItem + "' and [ProductName]='" + ProductName_CB.SelectedItem + "' and vchPrint='" + ColorListBox.SelectedItem + "'  order by convert(int,[vchBoxs]) asc ";
                cmd = new SqlCommand(selectCmd, conn);
                using (reader = cmd.ExecuteReader())
                {
                    if (reader.Read())
                    {
                        BoxMin = Convert.ToInt32(reader.GetString(reader.GetOrdinal("vchBoxs")));
                    }
                }

                //20190212
                selectCmd = "SELECT [vchBoxs] FROM [ShippingHead] where [ListDate]='" + ListDate_LB.SelectedItem + "' and [ProductName]='" + ProductName_CB.SelectedItem + "' and vchPrint='" + ColorListBox.SelectedItem + "' order by convert(int,[vchBoxs]) desc ";
                cmd = new SqlCommand(selectCmd, conn);
                using (reader = cmd.ExecuteReader())
                {
                    if (reader.Read())
                    {
                        BoxMax = Convert.ToInt32(reader.GetString(reader.GetOrdinal("vchBoxs")));
                    }
                }

                BoxRangeLabel.Text = BoxMin + "~" + BoxMax;

                //20190212
                selectCmd = "SELECT [vchBoxs] FROM [ShippingHead] where [ListDate]='" + ListDate_LB.SelectedItem + "' and [ProductName]='" + ProductName_CB.SelectedItem + "' and vchPrint='" + ColorListBox.SelectedItem + "'  order by convert(int,[vchBoxs]) asc ";
                cmd = new SqlCommand(selectCmd, conn);
                using (reader = cmd.ExecuteReader())
                {
                    while (reader.Read())
                    {
                        WhereBox_LB.Items.Add(reader.GetString(reader.GetOrdinal("vchBoxs")));
                    }
                }
            }

            Product_L.Text = "產品名稱：" + ProductName_CB.Text;

            ListDateLabel.Text = "嘜頭日期：" + ListDate_LB.SelectedItem;

            if (this.ListDate_LB.SelectedIndex != -1)
            {
                TodayDataButton.Enabled = true;
            }
            else
            {
                TodayDataButton.Enabled = false;
            }

            //ProductComboBox.Text="";
            //BoxsListBox.Items.Clear();
        }

        private void ProductName_CB_SelectedIndexChanged(object sender, EventArgs e)
        {
            //判斷做哪種瓶子
            if (ProductName_CB.SelectedItem.ToString().Contains("Composite") == true)
            {
                ProcessNo = "P56";
            }
            else
            {
                ProcessNo = "P26";
            }

            //載入產品Color  20190212
            BoxRangeLabel.Text = "";
            WhereBox_LB.Items.Clear();

            ListDate_LB.SelectedIndex = -1;
            ListDate_LB.Items.Clear();

            ColorListBox.SelectedIndex = -1;
            ColorListBox.Items.Clear();

            //載入[ShippingHead]的vchPrint
            using (conn = new SqlConnection(myConnectionString))
            {
                conn.Open();

                selectCmd = "SELECT  DISTINCT [vchPrint] FROM [ShippingHead]  where [ProductName]='" + ProductName_CB.SelectedItem.ToString() + "' order by [vchPrint] desc";
                cmd = new SqlCommand(selectCmd, conn);
                using (reader = cmd.ExecuteReader())
                {
                    while (reader.Read())
                    {
                        ColorListBox.Items.Add(reader.GetString(reader.GetOrdinal("vchPrint")));
                    }
                }
            }


            //載入賣頭的DATE
            //LoadListDate();

            //清除箱號Range Label
            //BoxRangeLabel.Text = "";

            if (ProductName_CB.SelectedIndex != -1)
            {
                if (ProductName_CB.Text.Contains("Composite") == true)
                {
                    NoLMCheckBox.CheckState = CheckState.Checked;
                    WeightCheckBox.CheckState = CheckState.Checked;
                    ComplexLabel.Visible = true;
                }
                else
                {
                    WeightCheckBox.CheckState = CheckState.Unchecked;
                    ComplexLabel.Visible = true;
                }
            }
        }

        private void WhereBox_LB_SelectedIndexChanged(object sender, EventArgs e)
        {
            if (WhereBox_LB.SelectedItem == null)
            {
                return;
            }

            //載入入箱狀況的圖片
            LoadPictrue();

            //載入入箱狀況資訊
            LoadSQLDate();

            if (ListDate_LB.SelectedIndex != -1 && ProductName_CB.Text != "")
            {
                using (conn = new SqlConnection(myConnectionString))
                {
                    conn.Open();

                    selectCmd = "SELECT isnull([CustomerPO],''),[vchPrint],[vchAssembly],isnull(PackingMarks,'') FROM [ShippingHead] where [ListDate]='" + ListDate_LB.SelectedItem + "' and [ProductName]='" + ProductName_CB.Text + "' and [vchBoxs]='" + WhereBox_LB.SelectedItem + "'";
                    cmd = new SqlCommand(selectCmd, conn);
                    using (reader = cmd.ExecuteReader())
                    {
                        if (reader.Read())
                        {
                            if (reader.GetString(0) != "")
                            {
                                CustomerPO_L.Text = "PO：" + reader.GetString(0);
                            }
                            else
                            {
                                //CustomerPO_L.Text = "PO：查無PO資料";
                            }
                            if (reader.GetString(1) != "")
                            {
                                PrintLabel.Text = "塗裝漆別：" + reader.GetString(1);
                                AssemblyLabel.Text = "氣瓶配件：" + reader.GetString(2);
                                ComplexLabel.Text = "嘜頭標籤：" + reader.GetString(3);
                            }
                            else
                            {

                                PrintLabel.Text = "塗裝漆別：";
                                AssemblyLabel.Text = "氣瓶配件：";
                                ComplexLabel.Text = "嘜頭標籤：" + reader.GetString(3);
                            }
                        }
                    }
                }
            }
            else
            {
                CustomerPO_L.Text = "PO：";
            }

            using (conn = new SqlConnection(myConnectionString))
            {
                conn.Open();

                selectCmd = "SELECT isnull([Storage],'') Storage FROM [ShippingHead] where [ListDate]='" + ListDate_LB.SelectedItem + "' and [ProductName]='" + ProductName_CB.Text + "' and [vchBoxs]='" + WhereBox_LB.SelectedItem + "'";
                cmd = new SqlCommand(selectCmd, conn);
                using (reader = cmd.ExecuteReader())
                {
                    if (reader.Read())
                    {
                        if (reader.GetString(reader.GetOrdinal("Storage")) == "Y")
                        {
                            StorageLabel.Text = "嘜頭狀態：入庫嘜頭";
                        }
                        else
                        {
                            StorageLabel.Text = "嘜頭狀態：出貨嘜頭";
                        }
                    }
                }
            }

            //檢查跳箱的
            Match();

            if (PalletNoLabel.Text.ToString().CompareTo("棧板號：") == 0)
            {
            }
            else if (PalletNoLabel.Text.ToString().Split('：')[1].Trim().CompareTo(APalletof()) != 0)
            {
                MessageBox.Show("請注意棧板編號變更為 " + APalletof() + "\nThe Pallet No. is change.");
            }

            NowBoxsLabel.Text = "目前箱號：" + WhereBox_LB.SelectedItem;
            ABoxofLabel.Text = "一箱幾隻：" + Aboxof();
            PalletNoLabel.Text = "棧板號：" + APalletof();
        }

        private void User_LB_SelectedIndexChanged(object sender, EventArgs e)
        {
            ID = User_LB.SelectedItem.ToString().Remove(6);
            User = User_LB.SelectedItem.ToString().Remove(0, 7);

            //身分確認
            DialogResult result = MessageBox.Show("工號   " + ID + "\n" + "操作員 " + User, "操作員確認", MessageBoxButtons.OKCancel);
            if (result == DialogResult.OK)
            {
                ProductName_CB.Enabled = true;
                User_LB.Enabled = false;

                UserLabel.Text = "操作人員：" + User_LB.SelectedItem;

                try
                {
                    //抓班表
                    using (conn = new SqlConnection(myConnectionString21))
                    {
                        conn.Open();

                        selectCmd = "SELECT C.WorkBeginTime,C.WorkEndTime FROM [HRMDB].[dbo].[AttendanceEmpRank] AS A LEFT JOIN [HRMDB].[dbo].[Employee] AS B ON A.EmployeeId=B.EmployeeId LEFT JOIN [HRMDB].[dbo].[AttendanceRank] AS C ON A.AttendanceRankId=C.AttendanceRankId WHERE A.Date = '" + DateTime.Now.ToString("yyyy-MM-dd 00:00:00.000") + "' and B.Code = '" + ID + "'";
                        cmd = new SqlCommand(selectCmd, conn);
                        using (reader = cmd.ExecuteReader())
                        {
                            if (reader.Read())
                            {
                                if (int.Parse(DateTime.Now.ToString("HHmm").ToString()) >= int.Parse(reader.GetString(0).Replace(":", "")) && int.Parse(DateTime.Now.ToString("HHmm")) <= int.Parse(reader.GetString(1).Replace(":", "")))
                                {
                                    worktype = "生產";
                                }
                                else
                                {
                                    worktype = "加班";
                                }
                            }
                            else
                            {
                                worktype = DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss") + "當日查無班表";
                            }
                        }
                    }

                    //初始化登錄登出時間
                    using (conn = new SqlConnection(myConnectionString))
                    {
                        conn.Open();

                        selectCmd = "INSERT INTO [LoginPackage] ([OperatorId],[Operator],[LoginTime],[LogoutTime],[Date]) VALUES('" + ID + "','" + User + "','" + DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss") + "','" + DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss") + "','" + DateTime.Now.ToString("yyyyMMdd") + "')";
                        cmd = new SqlCommand(selectCmd, conn);
                        cmd.ExecuteNonQuery();

                        selectCmd = "SELECT TOP(1) [ID] FROM [LoginPackage] WHERE [OperatorId] = '" + ID + "' ORDER BY [ID] desc";
                        cmd = new SqlCommand(selectCmd, conn);
                        using (reader = cmd.ExecuteReader())
                        {
                            while (reader.Read())
                            {
                                toolStripStatusLabel1.Text = reader.GetInt64(0).ToString();
                            }
                        }
                    }
                }
                catch (Exception ex)
                {
                    MessageBox.Show("" + ex);
                }

                return;
            }
            else if (result == DialogResult.Cancel)
            {
                ProductName_CB.Enabled = false;
                return;
            }
        }

        private void FirstPrinterComboBox_SelectedIndexChanged(object sender, EventArgs e)
        {
            IsChangePrinter = true;
        }

        private void PrinterRefreshButton_Click(object sender, EventArgs e)
        {
            FirstPrinterComboBox.Items.Clear();
            SecondPrinterComboBox.Items.Clear();

            List<string> PrinterList = new List<string>();
            PrinterList.Clear();

            PrintDocument printDoc = new PrintDocument();
            string sDefaultPrinter = printDoc.PrinterSettings.PrinterName; // 取得預設的印表機名稱

            // 取得安裝於電腦上的所有印表機名稱，加入 ListBox (Name : lbInstalledPrinters) 中
            foreach (string strPrinter in PrinterSettings.InstalledPrinters)
            {
                PrinterList.Add(strPrinter);
            }
            PrinterList.Sort();

            this.FirstPrinterComboBox.Items.AddRange(PrinterList.ToArray());
            SecondPrinterComboBox.Items.AddRange(PrinterList.ToArray());
            // ListBox (Name : lbInstalledPrinters) 選擇在預設印表機
            this.FirstPrinterComboBox.SelectedIndex = this.FirstPrinterComboBox.FindString(sDefaultPrinter);
            this.SecondPrinterComboBox.SelectedIndex = this.SecondPrinterComboBox.FindString(sDefaultPrinter);
        }

        private void Main_FormClosing(object sender, FormClosingEventArgs e)
        {
            if (time > 0)
            {
                DialogResult dr = MessageBox.Show("是否確定要關閉程式? \n Do you really want to exit?", "關閉程式  Exit", MessageBoxButtons.YesNo);

                if (dr == DialogResult.Yes)
                {
                    try
                    {
                        //更新登出時間
                        using (conn = new SqlConnection(myConnectionString))
                        {
                            conn.Open();

                            selectCmd = "UPDATE [LoginPackage] SET  [LogoutTime]= '" + DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss") + "' WHERE [ID] = '" + toolStripStatusLabel1.Text + "'";
                            cmd = new SqlCommand(selectCmd, conn);
                            cmd.ExecuteNonQuery();
                        }
                    }
                    catch (Exception ex)
                    {
                        MessageBox.Show(ex.ToString());
                    }
                    e.Cancel = false;
                }
                else if (dr == DialogResult.No)
                {
                    e.Cancel = true;
                }
            }
            else if (time <= 0)
            {
                e.Cancel = false;
            }
        }

        private void PrinterButton_Click(object sender, EventArgs e)
        {
            if (IsChangePrinter == true)
            {
                SetProfileString(FirstPrinterComboBox.SelectedItem.ToString());
                IsChangePrinter = false;
            }
        }

        public void SetProfileString(string sPrintName)
        {
            string DeviceLine = sPrintName + ",,";

            // 使用 WriteProfileString 設定預設印表機
            WriteProfileString("windows", "Device", DeviceLine);

            // 使用 SendMessage 傳送正確的通知給所有最上層的層級視窗。
            // WIN.INI 要在意的應用程式接聽此訊息，並且視需要重新讀取 WIN.ini
            //SendMessage(HWND_BROADCAST, WM_WININICHANGE, 0, "windows");//目前注解起來，不然會沒有回應
        }

        private void tabControl1_SelectedIndexChanged(object sender, EventArgs e)
        {
            PrintDocument printDoc = new PrintDocument();
            String sDefaultPrinter = printDoc.PrinterSettings.PrinterName; // 取得預設的印表機名稱


            // ListBox (Name : lbInstalledPrinters) 選擇在預設印表機
            this.FirstPrinterComboBox.SelectedIndex = this.FirstPrinterComboBox.FindString(sDefaultPrinter);
        }

        private void SecondPrintButton_Click(object sender, EventArgs e)
        {
            if (NoLMCylinderNOTextBox.Text != "")
            {
                //列印標籤貼紙

                MarkSecondPrintBarCode(NoLMCylinderNOTextBox.Text.ToString());
                OutputSecondPrintExcel();
                GC.Collect();
                SetProfileString(FirstPrinterComboBox.SelectedItem.ToString());
            }
            else
            {
                MessageBox.Show("請輸入氣瓶序號");
            }
        }

        private void RegulatorPrintButton_Click(object sender, EventArgs e)
        {
            Customer_Estratego_Form("", "Regulator 3000psi");
        }

        private void ColorListBox_SelectedIndexChanged(object sender, EventArgs e)
        {
            if (ColorListBox.SelectedIndex == -1)
            {
                return;
            }
            //載入賣頭的DATE
            LoadListDate();

            //清除箱號Range Label
            BoxRangeLabel.Text = "";

            if (this.ListDate_LB.SelectedIndex != -1)
            {
                TodayDataButton.Enabled = true;
            }
            else
            {
                TodayDataButton.Enabled = false;
            }
        }

        private void timer1_Tick_1(object sender, EventArgs e)
        {
            if (time > 0)
            {
                time = time - 1;
            }

            if (time == 0)
            {
                try
                {
                    //更新登出時間
                    using (conn = new SqlConnection(myConnectionString))
                    {
                        conn.Open();

                        selectCmd = "UPDATE [LoginPackage] SET  [LogoutTime]= '" + DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss") + "' WHERE [ID] = '" + toolStripStatusLabel1.Text + "'";
                        cmd = new SqlCommand(selectCmd, conn);
                        cmd.ExecuteNonQuery();
                    }
                }
                catch (Exception ex)
                {
                    MessageBox.Show(ex.ToString());
                }
                this.Close();
            }
        }
    }
}