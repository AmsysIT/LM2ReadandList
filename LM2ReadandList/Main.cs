//--for QRcode
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
using System.Reflection;
using System.Runtime.InteropServices;
using System.Threading;
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
        //////////////////////////

        string PalletNoString = "-";
        string Ebb = "";
        public int time=420;

        //資料庫宣告
        string myConnectionString, myConnectionString21;
        SqlConnection myConnection;
        string selectCmd, selectCmd1, selectCmd2;
        SqlConnection conn, conn1, conn2;
        SqlCommand cmd, cmd1, cmd2;
        SqlDataReader reader, reader1;
        SqlDataAdapter sqlAdapter;

        public string ESIGNmyConnectionString;
        string selectCmdP;
        SqlConnection connP;
        SqlCommand cmdP;
        SqlDataReader readerP;

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
        public string ID=null;
        public string User=null;
        public string worktype;

        public Main()
        {
            //資料庫路徑與位子
            myConnectionString = "Server=192.168.0.15;database=amsys;uid=sa;pwd=ams.sql;";
            ESIGNmyConnectionString = "Server=192.168.0.30;database=AMS2;uid=sa;pwd=Ams.sql;";
            myConnectionString21 = "Server=192.168.0.21;database=HRMDB;uid=sa;pwd=dsc;";
            InitializeComponent();
        }

        private void Form1_Load(object sender, EventArgs e)
        {
            IsChangePrinter = false;
            LoadUser();
            //LoadPrinter();
            LoadSQL_ShippingHead_ProductName();
        }

        private void LoadUser()
        {
            UserListComboBox.Items.Clear();

            selectCmd = "SELECT  vchTestersNo,vchTestersName FROM [LaserMarkTesters]  ORDER BY vchTestersNo";
            conn = new SqlConnection(myConnectionString);
            conn.Open();
            cmd = new SqlCommand(selectCmd, conn);
            reader = cmd.ExecuteReader();
            while (reader.Read())
            {
                UserListComboBox.Items.Add(reader.GetString(0) +" "+ reader.GetString(1));
            }
            reader.Close();
            conn.Close();
        }

        private void LoadPrinter()
        {
            FirstPrinterComboBox.Items.Clear();
            SecondPrinterComboBox.Items.Clear();

            List<string> PrinterList = new List<string>();
            PrinterList.Clear();

            PrintDocument printDoc = new PrintDocument();
            String sDefaultPrinter = printDoc.PrinterSettings.PrinterName; // 取得預設的印表機名稱

            // 取得安裝於電腦上的所有印表機名稱，加入 ListBox (Name : lbInstalledPrinters) 中
            foreach (String strPrinter in PrinterSettings.InstalledPrinters)
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

        private void LoadColor()
        {
            BoxRangeLabel.Text = "";
            BoxsListBox.Items.Clear();

            ListDateListBox.SelectedIndex = -1;
            ListDateListBox.Items.Clear();

            ColorListBox.SelectedIndex = -1;
            ColorListBox.Items.Clear();
            //載入[ShippingHead]的vchPrint
            selectCmd = "SELECT  DISTINCT [vchPrint] FROM [ShippingHead]  where [ProductName]='" + ProductComboBox.SelectedItem.ToString() + "' order by [vchPrint] desc";
            conn = new SqlConnection(myConnectionString);
            conn.Open();
            cmd = new SqlCommand(selectCmd, conn);
            reader = cmd.ExecuteReader();
            while (reader.Read())
            {
                ColorListBox.Items.Add(reader.GetString(0));
            }
            reader.Close();
            conn.Close();
        }

        private void LoadListDate()
        {
            ListDateListBox.SelectedIndex = -1;
            ListDateListBox.Items.Clear();

            //載入[ShippingHead]的ListDate
            //selectCmd = "SELECT  DISTINCT [ListDate] FROM [ShippingHead]  where [ProductName]='" + ProductComboBox.SelectedItem.ToString() + "' order by [ListDate] desc";
            //加入vchPrint之條件 20190212
            selectCmd = "SELECT  DISTINCT [ListDate] FROM [ShippingHead]  where [ProductName]='" + ProductComboBox.SelectedItem.ToString() + "' and vchPrint='" + ColorListBox.SelectedItem.ToString() + "' order by [ListDate] desc";
            conn = new SqlConnection(myConnectionString);
            conn.Open();
            cmd = new SqlCommand(selectCmd, conn);
            reader = cmd.ExecuteReader();
            while (reader.Read())
            {
                ListDateListBox.Items.Add(reader.GetString(0));
            }

            reader.Close();
            conn.Close();
        }

        private void LoadSQL_ShippingHead_ProductName()
        {
            ProductComboBox.Items.Clear();

            //載入[ShippingHead]的ListDate
            myConnection = new SqlConnection(myConnectionString);
            selectCmd = "SELECT DISTINCT [ProductName] FROM [amsys].[dbo].[ShippingHead]  order by [ProductName] asc";
            conn = new SqlConnection(myConnectionString);
            conn.Open();
            cmd = new SqlCommand(selectCmd, conn);
            reader = cmd.ExecuteReader();
            while (reader.Read())
            {
                ProductComboBox.Items.Add(reader.GetString(0));
            }
            reader.Close();
            conn.Close();
        }

        private void ListDateListBox_SelectedIndexChanged(object sender, EventArgs e)
        {
            BoxRangeLabel.Text = "";
            BoxsListBox.Items.Clear();

            int BoxMax = 0, BoxMin = 0;

            //查詢箱號最小值

            //selectCmd = "SELECT  * FROM [ShippingHead] where [ListDate]='" + ListDateListBox.SelectedItem + "' and [ProductName]='" + ProductComboBox.SelectedItem + "'  order by convert(int,[vchBoxs]) asc ";
            //20190212
            selectCmd = "SELECT  * FROM [ShippingHead] where [ListDate]='" + ListDateListBox.SelectedItem + "' and [ProductName]='" + ProductComboBox.SelectedItem + "' and vchPrint='" + ColorListBox.SelectedItem + "'  order by convert(int,[vchBoxs]) asc ";
            conn = new SqlConnection(myConnectionString);
            conn.Open();
            cmd = new SqlCommand(selectCmd, conn);
            reader = cmd.ExecuteReader();
            if (reader.Read())
            {
                BoxMin = Convert.ToInt32(reader.GetString(3));
            }
            reader.Close();
            conn.Close();

            //selectCmd = "SELECT  * FROM [ShippingHead] where [ListDate]='" + ListDateListBox.SelectedItem + "' and [ProductName]='" + ProductComboBox.SelectedItem + "' order by convert(int,[vchBoxs]) desc ";
            //20190212
            selectCmd = "SELECT  * FROM [ShippingHead] where [ListDate]='" + ListDateListBox.SelectedItem + "' and [ProductName]='" + ProductComboBox.SelectedItem + "' and vchPrint='" + ColorListBox.SelectedItem + "' order by convert(int,[vchBoxs]) desc ";
            conn = new SqlConnection(myConnectionString);
            conn.Open();
            cmd = new SqlCommand(selectCmd, conn);
            reader = cmd.ExecuteReader();
            if (reader.Read())
            {
                BoxMax = Convert.ToInt32(reader.GetString(3));
            }
            reader.Close();
            conn.Close();

            BoxRangeLabel.Text = BoxMin + "~" + BoxMax;

            //selectCmd = "SELECT  * FROM [ShippingHead] where [ListDate]='" + ListDateListBox.SelectedItem + "' and [ProductName]='" + ProductComboBox.SelectedItem + "'  order by convert(int,[vchBoxs]) asc ";
            //20190212
            selectCmd = "SELECT  * FROM [ShippingHead] where [ListDate]='" + ListDateListBox.SelectedItem + "' and [ProductName]='" + ProductComboBox.SelectedItem + "' and vchPrint='" + ColorListBox.SelectedItem + "'  order by convert(int,[vchBoxs]) asc ";
            conn = new SqlConnection(myConnectionString);
            conn.Open();
            cmd = new SqlCommand(selectCmd, conn);
            reader = cmd.ExecuteReader();
            while (reader.Read())
            {
                BoxsListBox.Items.Add(reader.GetString(3));
            }
            reader.Close();
            conn.Close();

            ProductLabel2.Text = "產品名稱：" + ProductComboBox.Text;

            ListDateLabel.Text = "嘜頭日期：" + ListDateListBox.SelectedItem;

            if (this.ListDateListBox.SelectedIndex != -1)
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

        private void ListDateComboBox_SelectedIndexChanged(object sender, EventArgs e)
        {
            //載入產品Color  20190212
            LoadColor();
            //載入賣頭的DATE
            //LoadListDate();

            ////清除箱號Range Label
            //BoxRangeLabel.Text = "";

            //if (this.ListDateListBox.SelectedIndex != -1)
            //{
            //    TodayDataButton.Enabled = true;
            //}
            //else
            //{
            //    TodayDataButton.Enabled = false;
            //}
            //若選擇複合瓶，則將雷刻相關設定取消
            if (ProductComboBox.SelectedIndex != -1)
            {
                if (ProductComboBox.Text.Contains("Composite") == true)
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

        private void BoxsListBox_SelectedIndexChanged(object sender, EventArgs e)
        {
            //載入入箱狀況的圖片
            LoadPictrue();

            //載入入箱狀況資訊
            LoadSQLDate();
            GetCustomerPO();

            GetStorage();
            //檢查跳箱的
            //LoadBoxsNo();
            //LoadBoxsCount();
            Match();

            if (PalletNoLabel.Text.ToString().CompareTo("棧板號：") == 0)
            {
            }
            else if (PalletNoLabel.Text.ToString().Split('：')[1].Trim().CompareTo(APalletof()) != 0)
            {
                MessageBox.Show("請注意棧板編號變更為 " + APalletof() + "\nThe Pallet No. is change.");
            }
            NowBoxsLabel.Text = "目前箱號：" + BoxsListBox.SelectedItem;
            ABoxofLabel.Text = "一箱幾隻：" + Aboxof();
            PalletNoLabel.Text = "棧板號：" + APalletof();
        }

        private void GetStorage()
        {
            myConnection = new SqlConnection(myConnectionString);
            selectCmd = "SELECT isnull([Storage],'') FROM [ShippingHead] where [ListDate]='" + ListDateListBox.SelectedItem + "' and [ProductName]='" + ProductComboBox.Text + "' and [vchBoxs]='" + BoxsListBox.SelectedItem + "'";
            conn = new SqlConnection(myConnectionString);
            conn.Open();
            cmd = new SqlCommand(selectCmd, conn);
            reader = cmd.ExecuteReader();
            if (reader.Read())
            {
                if (reader.GetString(0) == "Y")
                {
                    StorageLabel.Text = "嘜頭狀態：入庫嘜頭";
                }
                else
                {
                    StorageLabel.Text = "嘜頭狀態：出貨嘜頭";
                }
            }
            reader.Close();
            conn.Close();
        }

        private void GetCustomerPO()
        {
            if (ListDateListBox.SelectedIndex != -1 && ProductComboBox.Text != "")
            {
                myConnection = new SqlConnection(myConnectionString);
                selectCmd = "SELECT isnull([CustomerPO],''),[vchPrint],[vchAssembly],isnull(PackingMarks,'') FROM [ShippingHead] where [ListDate]='" + ListDateListBox.SelectedItem + "' and [ProductName]='" + ProductComboBox.Text + "' and [vchBoxs]='" + BoxsListBox.SelectedItem + "'";
                conn = new SqlConnection(myConnectionString);
                conn.Open();
                cmd = new SqlCommand(selectCmd, conn);
                reader = cmd.ExecuteReader();
                if (reader.Read())
                {
                    if (reader.GetString(0) != "")
                    {
                        CustomerPOLabel.Text = "PO：" + reader.GetString(0);
                    }
                    else
                    {
                        CustomerPOLabel.Text = "PO：查無PO資料";
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

                reader.Close();
                conn.Close();
            }
            else
            {
                CustomerPOLabel.Text = "PO：";
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
                myConnection = new SqlConnection(myConnectionString);
                selectCmd = "SELECT  * FROM [ShippingHead] where [ListDate]='" + ListDateListBox.SelectedItem + "' and [ProductName]='" + ProductComboBox.SelectedItem + "'and [vchBoxs]='" + BoxsListBox.SelectedItem + "'";
                conn = new SqlConnection(myConnectionString);
                conn.Open();
                cmd = new SqlCommand(selectCmd, conn);
                reader = cmd.ExecuteReader();
                if (reader.Read())
                {
                    bAboxof = reader.GetString(4);
                }
                reader.Close();
                conn.Close();

                if (bAboxof == "20" || bAboxof == "40")
                {
                    //載入[ShippingHead]的ListDate
                    myConnection = new SqlConnection(myConnectionString);
                    selectCmd = "SELECT  * FROM [ShippingBody] where [ListDate]='" + ListDateListBox.SelectedItem + "' and [ProductName]='" + ProductComboBox.SelectedItem + "'and [WhereBox]='" + BoxsListBox.SelectedItem + "'  order by Convert(INT,[WhereSeat]) DESC ";
                    conn = new SqlConnection(myConnectionString);
                    conn.Open();
                    cmd = new SqlCommand(selectCmd, conn);
                    reader = cmd.ExecuteReader();
                    if (reader.Read())
                    {
                        SeatNo = reader.GetString(5);

                        if (reader.IsDBNull(5) == false && (Convert.ToInt32(reader.GetString(5)) >= 1 && Convert.ToInt32(reader.GetString(5)) <= 20))
                        {
                            pictureBox1.Image = Image.FromFile(Application.StartupPath + @".\" + reader.GetString(5) + ".jpg");
                        }
                    }
                    else
                    {
                        pictureBox1.Image = Image.FromFile(Application.StartupPath + @".\0.jpg");
                    }
                    reader.Close();
                    conn.Close();
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
                    myConnection = new SqlConnection(myConnectionString);
                    selectCmd = "SELECT  * FROM [ShippingBody] where [ListDate]='" + ListDateListBox.SelectedItem + "' and [ProductName]='" + ProductComboBox.SelectedItem + "'and [WhereBox]='" + BoxsListBox.SelectedItem + "'  order by Convert(INT,[WhereSeat]) DESC ";
                    conn = new SqlConnection(myConnectionString);
                    conn.Open();
                    cmd = new SqlCommand(selectCmd, conn);
                    reader = cmd.ExecuteReader();
                    if (reader.Read())
                    {
                        switch (reader.GetString(5))
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
                    reader.Close();
                    conn.Close();
                }
                else if (bAboxof == "25")
                {
                    //載入[ShippingHead]的ListDate
                    myConnection = new SqlConnection(myConnectionString);
                    selectCmd = "SELECT  * FROM [ShippingBody] where [ListDate]='" + ListDateListBox.SelectedItem + "' and [ProductName]='" + ProductComboBox.SelectedItem + "'and [WhereBox]='" + BoxsListBox.SelectedItem + "'  order by Convert(INT,[WhereSeat]) DESC ";
                    conn = new SqlConnection(myConnectionString);
                    conn.Open();
                    cmd = new SqlCommand(selectCmd, conn);
                    reader = cmd.ExecuteReader();
                    if (reader.Read())
                    {
                        switch (reader.GetString(5))
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
                    reader.Close();
                    conn.Close();
                }

                else if (bAboxof == "8")
                {
                    
                    //載入[ShippingHead]的ListDate
                    myConnection = new SqlConnection(myConnectionString);
                    selectCmd = "SELECT  * FROM [ShippingBody] where [ListDate]='" + ListDateListBox.SelectedItem + "' and [ProductName]='" + ProductComboBox.SelectedItem + "'and [WhereBox]='" + BoxsListBox.SelectedItem + "'  order by Convert(INT,[WhereSeat]) DESC ";
                    conn = new SqlConnection(myConnectionString);
                    conn.Open();
                    cmd = new SqlCommand(selectCmd, conn);
                    reader = cmd.ExecuteReader();
                    if (reader.Read())
                    {
                        
                        switch (reader.GetString(5))
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
                    reader.Close();
                    conn.Close();
                }
                else if (bAboxof == "12")
                {
                    
                    //載入[ShippingHead]的ListDate
                    myConnection = new SqlConnection(myConnectionString);
                    selectCmd = "SELECT  * FROM [ShippingBody] where [ListDate]='" + ListDateListBox.SelectedItem + "' and [ProductName]='" + ProductComboBox.SelectedItem + "'and [WhereBox]='" + BoxsListBox.SelectedItem + "'  order by Convert(INT,[WhereSeat]) DESC ";
                    conn = new SqlConnection(myConnectionString);
                    conn.Open();
                    cmd = new SqlCommand(selectCmd, conn);
                    reader = cmd.ExecuteReader();
                    if (reader.Read())
                    {
                        switch (reader.GetString(5))
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
                    reader.Close();
                    conn.Close();
                }
                else if (bAboxof == "36")
                {
                    //載入[ShippingHead]的ListDate
                    myConnection = new SqlConnection(myConnectionString);
                    selectCmd = "SELECT  * FROM [ShippingBody] where [ListDate]='" + ListDateListBox.SelectedItem + "' and [ProductName]='" + ProductComboBox.SelectedItem + "'and [WhereBox]='" + BoxsListBox.SelectedItem + "'  order by Convert(INT,[WhereSeat]) DESC ";
                    conn = new SqlConnection(myConnectionString);
                    conn.Open();
                    cmd = new SqlCommand(selectCmd, conn);
                    reader = cmd.ExecuteReader();
                    if (reader.Read())
                    {
                        if (reader.IsDBNull(5) == false && (Convert.ToInt32(reader.GetString(5)) >= 1 && Convert.ToInt32(reader.GetString(5)) <= 117))
                        {
                            pictureBox1.Image = Image.FromFile(Application.StartupPath + @".\36\36-" + reader.GetString(5) + ".jpg");
                        }
                    }
                    else
                    {
                        pictureBox1.Image = Image.FromFile(Application.StartupPath + @".\36\36-0.jpg");
                    }
                    reader.Close();
                    conn.Close();
                }
                else if (bAboxof == "117")
                {
                    //載入[ShippingHead]的ListDate
                    myConnection = new SqlConnection(myConnectionString);
                    selectCmd = "SELECT  * FROM [ShippingBody] where [ListDate]='" + ListDateListBox.SelectedItem + "' and [ProductName]='" + ProductComboBox.SelectedItem + "'and [WhereBox]='" + BoxsListBox.SelectedItem + "'  order by Convert(INT,[WhereSeat]) DESC ";
                    conn = new SqlConnection(myConnectionString);
                    conn.Open();
                    cmd = new SqlCommand(selectCmd, conn);
                    reader = cmd.ExecuteReader();
                    if (reader.Read())
                    {
                        if (reader.IsDBNull(5) == false && (Convert.ToInt32(reader.GetString(5)) >= 1 && Convert.ToInt32(reader.GetString(5)) <= 117))
                        {
                            pictureBox1.Image = Image.FromFile(Application.StartupPath + @".\117\117-" + reader.GetString(5) + ".jpg");
                        }
                    }
                    else
                    {
                        pictureBox1.Image = Image.FromFile(Application.StartupPath + @".\117\117-0.jpg");
                    }
                    reader.Close();
                    conn.Close();
                }
                else if (bAboxof == "30")
                {
                    //載入[ShippingHead]的ListDate
                    myConnection = new SqlConnection(myConnectionString);
                    selectCmd = "SELECT  * FROM [ShippingBody] where [ListDate]='" + ListDateListBox.SelectedItem + "' and [ProductName]='" + ProductComboBox.SelectedItem + "'and [WhereBox]='" + BoxsListBox.SelectedItem + "'  order by Convert(INT,[WhereSeat]) DESC ";
                    conn = new SqlConnection(myConnectionString);
                    conn.Open();
                    cmd = new SqlCommand(selectCmd, conn);
                    reader = cmd.ExecuteReader();
                    if (reader.Read())
                    {
                        if (reader.IsDBNull(5) == false && (Convert.ToInt32(reader.GetString(5)) >= 1 && Convert.ToInt32(reader.GetString(5)) <= 30))
                        {
                            pictureBox1.Image = Image.FromFile(Application.StartupPath + @".\30\30-" + reader.GetString(5) + ".jpg");
                        }
                    }
                    else
                    {
                        pictureBox1.Image = Image.FromFile(Application.StartupPath + @".\30\30-0.jpg");
                    }
                    reader.Close();
                    conn.Close();
                }
                else if (bAboxof == "111")
                {
                    //載入[ShippingHead]的ListDate
                    myConnection = new SqlConnection(myConnectionString);
                    selectCmd = "SELECT  * FROM [ShippingBody] where [ListDate]='" + ListDateListBox.SelectedItem + "' and [ProductName]='" + ProductComboBox.SelectedItem + "'and [WhereBox]='" + BoxsListBox.SelectedItem + "'  order by Convert(INT,[WhereSeat]) DESC ";
                    conn = new SqlConnection(myConnectionString);
                    conn.Open();
                    cmd = new SqlCommand(selectCmd, conn);
                    reader = cmd.ExecuteReader();
                    if (reader.Read())
                    {
                        if (reader.IsDBNull(5) == false && (Convert.ToInt32(reader.GetString(5)) >= 1 && Convert.ToInt32(reader.GetString(5)) <= 111))
                        {
                            pictureBox1.Image = Image.FromFile(Application.StartupPath + @".\111\111-" + reader.GetString(5) + ".jpg");
                        }
                    }
                    else
                    {
                        pictureBox1.Image = Image.FromFile(Application.StartupPath + @".\111\111-0.jpg");
                    }
                    reader.Close();
                    conn.Close();
                }
                else if (bAboxof == "4" || bAboxof == "3")
                {
                    //載入[ShippingHead]的ListDate
                    selectCmd = "SELECT  * FROM [ShippingBody] where [ListDate]='" + ListDateListBox.SelectedItem + "' and [ProductName]='" + ProductComboBox.SelectedItem + "'and [WhereBox]='" + BoxsListBox.SelectedItem + "'  order by Convert(INT,[WhereSeat]) DESC ";
                    conn = new SqlConnection(myConnectionString);
                    conn.Open();
                    cmd = new SqlCommand(selectCmd, conn);
                    reader = cmd.ExecuteReader();
                    if (reader.Read())
                    {
                        switch (reader.GetString(5))
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
                    reader.Close();
                    conn.Close();
                }
                else if (bAboxof == "2")
                {
                    //載入[ShippingHead]的ListDate
                    selectCmd = "SELECT  * FROM [ShippingBody] where [ListDate]='" + ListDateListBox.SelectedItem + "' and [ProductName]='" + ProductComboBox.SelectedItem + "'and [WhereBox]='" + BoxsListBox.SelectedItem + "'  order by Convert(INT,[WhereSeat]) DESC ";
                    conn = new SqlConnection(myConnectionString);
                    conn.Open();
                    cmd = new SqlCommand(selectCmd, conn);
                    reader = cmd.ExecuteReader();
                    if (reader.Read())
                    {
                        switch (reader.GetString(5))
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
                    reader.Close();
                    conn.Close();
                }
            }
            catch 
            {
                Thread.Sleep(500);
                LoadPictrue();
            }
        }

        private void UserListComboBox_SelectedIndexChanged(object sender, EventArgs e)
        {
            ID = UserListComboBox.SelectedItem.ToString().Remove(6);
            User= UserListComboBox.SelectedItem.ToString().Remove(0,7);

            //身分確認
            DialogResult result = MessageBox.Show("工號   " + ID + "\n" + "操作員 " + User, "操作員確認", MessageBoxButtons.OKCancel);
            if (result == DialogResult.OK)
            {
                ProductComboBox.Enabled = true;
                UserListComboBox.Enabled = false;

                UserLabel.Text = "操作人員：" + UserListComboBox.SelectedItem;

                try
                {
                    //抓班表
                    selectCmd = "SELECT C.WorkBeginTime,C.WorkEndTime FROM [HRMDB].[dbo].[AttendanceEmpRank] AS A LEFT JOIN [HRMDB].[dbo].[Employee] AS B ON A.EmployeeId=B.EmployeeId LEFT JOIN [HRMDB].[dbo].[AttendanceRank] AS C ON A.AttendanceRankId=C.AttendanceRankId WHERE A.Date = '" + DateTime.Now.ToString("yyyy-MM-dd 00:00:00.000") + "' and B.Code = '" + ID + "'";
                    conn = new SqlConnection(myConnectionString21);
                    conn.Open();
                    cmd = new SqlCommand(selectCmd, conn);
                    reader = cmd.ExecuteReader();
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
                    reader.Close();
                    conn.Close();

                    //初始化登錄登出時間
                    selectCmd = "INSERT INTO [LoginPackage] ([OperatorId],[Operator],[LoginTime],[LogoutTime],[Date]) VALUES('" + ID + "','" + User + "','" + DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss") + "','" + DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss") + "','" + DateTime.Now.ToString("yyyyMMdd") + "')";
                    conn = new SqlConnection(myConnectionString);
                    conn.Open();
                    cmd = new SqlCommand(selectCmd, conn);
                    cmd.ExecuteNonQuery();
                    conn.Close();
                
                    selectCmd = "SELECT TOP(1) [ID] FROM [LoginPackage] WHERE [OperatorId] = '" + ID + "' ORDER BY [ID] desc";
                    conn = new SqlConnection(myConnectionString);
                    conn.Open();
                    cmd = new SqlCommand(selectCmd, conn);
                    reader = cmd.ExecuteReader();
                    while (reader.Read())
                    {
                        toolStripStatusLabel1.Text = reader.GetInt64(0).ToString();
                    }
                    reader.Close();
                    conn.Close();
                }
                catch (Exception ex)
                {
                    MessageBox.Show("" + ex);
                }

                return;
            }
            else if (result == DialogResult.Cancel)
            {
                ProductComboBox.Enabled = false;
                return;
            }
        }

        private void StepTimer_Tick(object sender, EventArgs e)
        {
            if (UserListComboBox.Text == "")
            {
                StepLabel1.BackColor = Color.Red;
            }
            else
            {
                StepLabel1.BackColor = Color.MediumTurquoise;
            }

            if (ProductComboBox.Text == "")
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

            if (ListDateListBox.SelectedIndex == -1)
            {
                StepLabel4.BackColor = Color.Red;
            }
            else
            {
                StepLabel4.BackColor = Color.MediumTurquoise;
            }

            if (BoxsListBox.SelectedIndex == -1)
            {
                StepLabel5.BackColor = Color.Red;
            }
            else
            {
                StepLabel5.BackColor = Color.MediumTurquoise;
            }
            
            if (ProductComboBox.Text == "")
            {
                ProductLabel2.Text = "產品名稱：";
            }

            if (BoxsListBox.SelectedIndex == -1)
            {
                NowBoxsLabel.Text = "目前箱號：";
                ABoxofLabel.Text = "一箱幾隻：";
                PrintLabel.Text = "塗裝漆別";
                AssemblyLabel.Text = "氣瓶配件";
                StorageLabel.Text = "嘜頭狀態：";
                CustomerPOLabel.Text = "PO：";
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
            myConnection = new SqlConnection(myConnectionString);
            selectCmd = "SELECT  * FROM [ShippingHead] where [ListDate]='" + ListDateListBox.SelectedItem + "' and [ProductName]='" + ProductComboBox.SelectedItem + "'and [vchBoxs]='" + BoxsListBox.SelectedItem + "' ";
            conn = new SqlConnection(myConnectionString);
            conn.Open();
            cmd = new SqlCommand(selectCmd, conn);
            reader = cmd.ExecuteReader();
            while (reader.Read())
            {
                temp = reader.GetString(4);
            }
            reader.Close();
            conn.Close();

            return temp;
        }

        public string APalletof()
        {
            string temp = "";

            //載入[ShippingHead]的棧板編號
            myConnection = new SqlConnection(myConnectionString);
            selectCmd = "SELECT  isnull(PalletNo,'') FROM [ShippingHead] where [ListDate]='" + ListDateListBox.SelectedItem + "' and [ProductName]='" + ProductComboBox.SelectedItem + "'and [vchBoxs]='" + BoxsListBox.SelectedItem + "' ";
            conn = new SqlConnection(myConnectionString);
            conn.Open();
            cmd = new SqlCommand(selectCmd, conn);
            reader = cmd.ExecuteReader();
            if (reader.Read())
            {
                temp = reader.GetString(0);
            }
            reader.Close();
            conn.Close();

            return temp;
        }

        private void LuckButton_Click(object sender, EventArgs e)
        {
            if (UserListComboBox.Text == "")
            {
                MessageBox.Show("尚未選擇測試人員", "警告");
                return;
            }
            else if (ListDateListBox.SelectedIndex == -1)
            {
                MessageBox.Show("尚未選擇嘜頭日期", "警告");
                return;
            }
            else if (ProductComboBox.Text == "")
            {
                MessageBox.Show("尚未選擇嘜頭名稱", "警告");
                return;
            }
            else if (BoxsListBox.SelectedIndex == -1)
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
            
            myConnection = new SqlConnection(myConnectionString);
            selectCmd = "SELECT  * FROM [LaserMarkDirection] ";
            conn = new SqlConnection(myConnectionString);
            conn.Open();
            cmd = new SqlCommand(selectCmd, conn);
            reader = cmd.ExecuteReader();
            if (reader.Read())
            {
                where = reader.GetString(0);
            }
            reader.Close();
            conn.Close();
            
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

            myConnection = new SqlConnection(myConnectionString);
            selectCmd = "SELECT  * FROM [LaserMarkDirection] ";
            conn = new SqlConnection(myConnectionString);
            conn.Open();
            cmd = new SqlCommand(selectCmd, conn);
            reader = cmd.ExecuteReader();
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
            reader.Close();
            conn.Close();
        }
        
        public void LoadSQLDate()
        {
            DataTable DT = new DataTable();
            //載入已放入的氣瓶內容
            dataGridView1.AutoGenerateColumns = false;

            selectCmd = "SELECT [WhereBox] 嘜頭箱號,[WhereSeat] 嘜頭位置,[CylinderNumbers] 氣瓶序號,[CustomerBarCode] 客戶BARCODE,[CylinderWeight] 氣瓶重量 FROM [ShippingBody] " +
                "Where [ListDate] = @ListDate and [ProductName]= @ProductName and [WhereBox] = @WhereBox  order by Convert(INT,[WhereSeat]) asc ";
            sqlAdapter = new SqlDataAdapter(selectCmd,myConnectionString);
            sqlAdapter.SelectCommand.Parameters.AddWithValue("@ListDate", ListDateListBox.SelectedItem);
            sqlAdapter.SelectCommand.Parameters.AddWithValue("@ProductName", ProductComboBox.SelectedItem);
            sqlAdapter.SelectCommand.Parameters.AddWithValue("@WhereBox", BoxsListBox.SelectedItem);
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
            System.Drawing.Image img = MyCode.GetCodeImage(BoxNo, Code128.Encode.Code128A);

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
            if (BoxsListBox.SelectedIndex == -1)
            {
                MessageBox.Show("請選擇箱號.Select the box number.");
                return;
            }
            
            MakeQRCode();
            MarkBarCode(BoxsListBox.SelectedItem.ToString());

            OutputExcel();
            GC.Collect();
        }

        //EXCEL輸出
        private void OutputExcel()
        {
            //判斷一箱幾隻
            string Aboxof = "", PackingMarks = "", Client = "";
            PalletNoString = "-";
            //判斷一箱幾隻
            selectCmd = "SELECT vchAboxof ,isnull(PackingMarks,''),isnull(PalletNo,'-'),Client FROM [ShippingHead] where [ListDate]='" + ListDateListBox.SelectedItem + "' and [ProductName]='" + ProductComboBox.SelectedItem + "'and [vchBoxs]='" + BoxsListBox.SelectedItem + "'";
            conn = new SqlConnection(myConnectionString);
            conn.Open();
            cmd = new SqlCommand(selectCmd, conn);
            reader = cmd.ExecuteReader();
            if (reader.Read())
            {
                Aboxof = reader.GetString(0);
                PackingMarks = reader.GetValue(1).ToString();
                PalletNoString = reader.GetValue(2).ToString();
                Client = reader.GetValue(3).ToString();
            }
            reader.Close();
            conn.Close();
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
                String BoxsListBoxIndex = BoxsListBox.SelectedIndex.ToString();
                BoxsListBox.SelectedIndex = (Convert.ToInt32(BoxsListBoxIndex) + 1);
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
            selectCmd = "SELECT  * FROM [ShippingHead] where [ListDate]='" + ListDateListBox.SelectedItem + "' and [ProductName]='" + ProductComboBox.Text + "' and [vchBoxs]='" + BoxsListBox.SelectedItem + "'";
            conn = new SqlConnection(myConnectionString);
            conn.Open();
            cmd = new SqlCommand(selectCmd, conn);
            reader = cmd.ExecuteReader();
            while (reader.Read())
            {
                Excel.Worksheet xSheet = (Excel.Worksheet)oWB.Sheets[1];
                if (Aboxof == "8")
                {
                    oSheet.Shapes.AddPicture(oneadd + BoxsListBox.SelectedItem + ".png", Microsoft.Office.Core.MsoTriState.msoFalse,
                                        Microsoft.Office.Core.MsoTriState.msoTrue, 252, oneY, 200, 30);
                    oSheet.Shapes.AddPicture(oneadd + PalletNoString + ".png", Microsoft.Office.Core.MsoTriState.msoFalse,
                    Microsoft.Office.Core.MsoTriState.msoTrue, 704, oneY, 200, 30);
                }
                else if (Aboxof == "16")
                {
                    oSheet.Shapes.AddPicture(oneadd + BoxsListBox.SelectedItem + ".png", Microsoft.Office.Core.MsoTriState.msoFalse,
                                        Microsoft.Office.Core.MsoTriState.msoTrue, 252, oneY, 200, 30);
                    oSheet.Shapes.AddPicture(oneadd + PalletNoString + ".png", Microsoft.Office.Core.MsoTriState.msoFalse,
                    Microsoft.Office.Core.MsoTriState.msoTrue, 704, oneY, 200, 30);
                }
                else if (Aboxof == "10")
                {
                    oSheet.Shapes.AddPicture(oneadd + BoxsListBox.SelectedItem + ".png", Microsoft.Office.Core.MsoTriState.msoFalse,
                                        Microsoft.Office.Core.MsoTriState.msoTrue, 222, oneY, 200, 30);
                    oSheet.Shapes.AddPicture(oneadd + PalletNoString + ".png", Microsoft.Office.Core.MsoTriState.msoFalse,
                    Microsoft.Office.Core.MsoTriState.msoTrue, 704, oneY, 200, 30);
                }
                else if (Aboxof == "20")
                {
                    oSheet.Shapes.AddPicture(oneadd + BoxsListBox.SelectedItem + ".png", Microsoft.Office.Core.MsoTriState.msoFalse,
                    Microsoft.Office.Core.MsoTriState.msoTrue, 265, oneY, 200, 30);
                    //PalletNoString
                    oSheet.Shapes.AddPicture(oneadd + PalletNoString + ".png", Microsoft.Office.Core.MsoTriState.msoFalse,
                    Microsoft.Office.Core.MsoTriState.msoTrue, 754, oneY, 200, 30);
                }
                else if (Aboxof == "40")
                {
                    oSheet.Shapes.AddPicture(oneadd + BoxsListBox.SelectedItem + ".png", Microsoft.Office.Core.MsoTriState.msoFalse,
                    Microsoft.Office.Core.MsoTriState.msoTrue, 220, oneY, 200, 30);
                    oSheet.Shapes.AddPicture(oneadd + PalletNoString + ".png", Microsoft.Office.Core.MsoTriState.msoFalse,
                    Microsoft.Office.Core.MsoTriState.msoTrue, 704, oneY, 200, 30);
                }
                else if (Aboxof == "36")
                {
                    oSheet.Shapes.AddPicture(oneadd + BoxsListBox.SelectedItem + ".png", Microsoft.Office.Core.MsoTriState.msoFalse,
                    Microsoft.Office.Core.MsoTriState.msoTrue, 240, oneY, 200, 30);
                    oSheet.Shapes.AddPicture(oneadd + PalletNoString + ".png", Microsoft.Office.Core.MsoTriState.msoFalse,
                    Microsoft.Office.Core.MsoTriState.msoTrue, 774, oneY, 200, 30);
                }
                else if (Aboxof == "25")
                {
                    oSheet.Shapes.AddPicture(oneadd + BoxsListBox.SelectedItem + ".png", Microsoft.Office.Core.MsoTriState.msoFalse,
                    Microsoft.Office.Core.MsoTriState.msoTrue, 268, oneY, 200, 30);
                    oSheet.Shapes.AddPicture(oneadd + PalletNoString + ".png", Microsoft.Office.Core.MsoTriState.msoFalse,
                    Microsoft.Office.Core.MsoTriState.msoTrue, 754, oneY, 200, 30);
                }
                else if (Aboxof == "30")
                {
                    oSheet.Shapes.AddPicture(oneadd + BoxsListBox.SelectedItem + ".png", Microsoft.Office.Core.MsoTriState.msoFalse,
                    Microsoft.Office.Core.MsoTriState.msoTrue, 268, 431, 200, 30);
                    oSheet.Shapes.AddPicture(oneadd + PalletNoString + ".png", Microsoft.Office.Core.MsoTriState.msoFalse,
                    Microsoft.Office.Core.MsoTriState.msoTrue, 754, 430, 200, 30);
                }
                else if (Aboxof == "15")
                {
                    oSheet.Shapes.AddPicture(oneadd + BoxsListBox.SelectedItem + ".png", Microsoft.Office.Core.MsoTriState.msoFalse,
                    Microsoft.Office.Core.MsoTriState.msoTrue, 260, oneY, 200, 30);
                    oSheet.Shapes.AddPicture(oneadd + PalletNoString + ".png", Microsoft.Office.Core.MsoTriState.msoFalse,
                    Microsoft.Office.Core.MsoTriState.msoTrue, 754, oneY, 200, 30);
                }
                else if (Aboxof == "12")
                {
                    oSheet.Shapes.AddPicture(oneadd + BoxsListBox.SelectedItem + ".png", Microsoft.Office.Core.MsoTriState.msoFalse,
                    Microsoft.Office.Core.MsoTriState.msoTrue, 260, oneY, 200, 30);
                    oSheet.Shapes.AddPicture(oneadd + PalletNoString + ".png", Microsoft.Office.Core.MsoTriState.msoFalse,
                    Microsoft.Office.Core.MsoTriState.msoTrue, 704, oneY, 200, 30);
                }
                else if (Aboxof == "117")
                {
                    oSheet.Shapes.AddPicture(oneadd + BoxsListBox.SelectedItem + ".png", Microsoft.Office.Core.MsoTriState.msoFalse,
                    Microsoft.Office.Core.MsoTriState.msoTrue, 200, 587, 200, 30);
                    oSheet.Shapes.AddPicture(oneadd + PalletNoString + ".png", Microsoft.Office.Core.MsoTriState.msoFalse,
                    Microsoft.Office.Core.MsoTriState.msoTrue, 620, 587, 130, 30);
                }
                else if (Aboxof == "111")
                {
                    oSheet.Shapes.AddPicture(oneadd + BoxsListBox.SelectedItem + ".png", Microsoft.Office.Core.MsoTriState.msoFalse,
                    Microsoft.Office.Core.MsoTriState.msoTrue, 200, 587, 200, 30);
                    oSheet.Shapes.AddPicture(oneadd + PalletNoString + ".png", Microsoft.Office.Core.MsoTriState.msoFalse,
                    Microsoft.Office.Core.MsoTriState.msoTrue, 620, 587, 130, 30);
                }
                else if (Aboxof == "4"|| Aboxof == "3")
                {
                    oSheet.Shapes.AddPicture(oneadd + BoxsListBox.SelectedItem + ".png", Microsoft.Office.Core.MsoTriState.msoFalse,
                                        Microsoft.Office.Core.MsoTriState.msoTrue, 256, oneY, 200, 30);
                    oSheet.Shapes.AddPicture(oneadd + PalletNoString + ".png", Microsoft.Office.Core.MsoTriState.msoFalse,
                   Microsoft.Office.Core.MsoTriState.msoTrue, 704, oneY, 200, 30);
                }
                else if (Aboxof == "2")
                {
                    oSheet.Shapes.AddPicture(oneadd + BoxsListBox.SelectedItem + ".png", Microsoft.Office.Core.MsoTriState.msoFalse,
                                        Microsoft.Office.Core.MsoTriState.msoTrue, 256, oneY, 200, 30);
                    oSheet.Shapes.AddPicture(oneadd + PalletNoString + ".png", Microsoft.Office.Core.MsoTriState.msoFalse,
                  Microsoft.Office.Core.MsoTriState.msoTrue, 704, oneY, 200, 30);
                }
                else if (Aboxof == "1")
                {
                    oSheet.Shapes.AddPicture(oneadd + BoxsListBox.SelectedItem + ".png", Microsoft.Office.Core.MsoTriState.msoFalse,
                                        Microsoft.Office.Core.MsoTriState.msoTrue, 256, oneY, 200, 30);
                }
            }

            string Client = "";
            if (Aboxof == "20")
            {
                string HowMuch = "";
                int Cumulative = 0;
                int Total = 0;

                //載入嘜頭資料
                myConnection = new SqlConnection(myConnectionString);
                selectCmd = "SELECT  * FROM [ShippingHead] where [ListDate]='" + ListDateListBox.SelectedItem + "' and [ProductName]='" + ProductComboBox.SelectedItem + "'";
                conn = new SqlConnection(myConnectionString);
                conn.Open();
                cmd = new SqlCommand(selectCmd, conn);
                reader = cmd.ExecuteReader();
                while (reader.Read())
                {
                    HowMuch = reader.GetString(4);
                    Cumulative++;
                }
                reader.Close();
                conn.Close();

                Total = Convert.ToInt32(HowMuch) * Cumulative;

                //載入嘜頭資料
                myConnection = new SqlConnection(myConnectionString);
                selectCmd = "SELECT  [Client],isnull([CustomerPO],''),isnull([CustomerProductName],''),isnull([CustomerProductNo],''),[vchBoxs],isnull(PalletNo,'') FROM [ShippingHead] where [ListDate]='" + ListDateListBox.SelectedItem + "' and [ProductName]='" + ProductComboBox.SelectedItem + "'and [vchBoxs]='" + BoxsListBox.SelectedItem + "'";
                conn = new SqlConnection(myConnectionString);
                conn.Open();
                cmd = new SqlCommand(selectCmd, conn);
                reader = cmd.ExecuteReader();
                if (reader.Read())
                {
                    Client = reader.GetString(0).Trim();
                    //載入客戶產品名稱
                    oSheet.Cells[1, 7] = reader.GetString(2);

                    //載入客戶產品型號
                    oSheet.Cells[2, 7] = reader.GetString(3);

                    //載入一箱幾隻
                    oSheet.Cells[4, 7] = Getcount;

                    //載入箱號
                    oSheet.Cells[10, 2] = reader.GetString(4);

                    //載入客戶名稱
                    oSheet.Cells[3, 7] = reader.GetString(0);
                    //載入箱號
                    oSheet.Cells[10, 10] = reader.GetString(5);

                    //該客戶要其自己的logo
                    if (reader.GetString(0).Trim().CompareTo("Wicked Sportz") == 0)
                    {
                        oSheet.Shapes.AddPicture(Application.StartupPath + @".\LOGO-ENAIRGY_Wicked Sportz.jpg", Microsoft.Office.Core.MsoTriState.msoFalse,
                                        Microsoft.Office.Core.MsoTriState.msoTrue, 2, 17, 212, 125);
                    }
                    if (reader.GetString(0).Trim().CompareTo("達成數位") == 0)
                    {
                        oSheet.Shapes.AddPicture(Application.StartupPath + @".\LOGO_DCT.jpg", Microsoft.Office.Core.MsoTriState.msoFalse,
                                        Microsoft.Office.Core.MsoTriState.msoTrue, 2, 17, 212, 125);
                    }
                    else if (Client.ToUpper().StartsWith("EMB"))
                    {
                        oSheet.Shapes.AddPicture(Application.StartupPath + @".\LOGO_EMB.png", Microsoft.Office.Core.MsoTriState.msoFalse,
                                            Microsoft.Office.Core.MsoTriState.msoTrue, 2, 17, 212, 125);
                    }
                }
                reader.Close();
                conn.Close();
                //////////
                int serialnooneX = 7, serialnooneY = 205;
                string serialnooneadd = @"C:\SerialNoCode\";
                //////
                string FirstCNO = "";

                //載入嘜頭氣瓶序號位子
                myConnection = new SqlConnection(myConnectionString);
                selectCmd = "SELECT  * FROM [ShippingBody] where [ListDate]='" + ListDateListBox.SelectedItem + "' and [ProductName]='" + ProductComboBox.SelectedItem + "'and [WhereBox]='" + BoxsListBox.SelectedItem + "'";
                conn = new SqlConnection(myConnectionString);
                conn.Open();
                cmd = new SqlCommand(selectCmd, conn);
                reader = cmd.ExecuteReader();
                while (reader.Read())
                {
                    serialnooneX = 3; serialnooneY = 203;
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
                    }

                    /////////
                    serialnooneX = serialnooneX + ((Convert.ToInt32(reader.GetString(5)) + 4) % 5) * 145;
                    serialnooneY = serialnooneY + ((Convert.ToInt32(reader.GetString(5)) - 1) / 5) * 56;
                    oSheet.Shapes.AddPicture(serialnooneadd + reader.GetString(3) + ".png", Microsoft.Office.Core.MsoTriState.msoFalse,
                    Microsoft.Office.Core.MsoTriState.msoTrue, serialnooneX, serialnooneY, 44, 44);// 130, 22);
                    /////////
                }
                reader.Close();

                if ((Client == "SGA" || Client == "Airtanks Limited") && PackingMarks.Trim().CompareTo("SGA-GLADIATAIR") == 0)
                {
                    string ProductNO = "";
                    //該客戶要其自己的logo  PartNo   Part Description
                    selectCmd = "SELECT  Product_NO FROM MSNBody,Manufacturing where [CylinderNo]='" + FirstCNO + "' and vchManufacturingNo=  Manufacturing_NO";
                    cmd = new SqlCommand(selectCmd, conn);
                    reader = cmd.ExecuteReader();
                    if (reader.Read())
                    {
                        ProductNO = reader.GetValue(0).ToString();
                    }
                    reader.Close();

                    selectCmd = "SELECT  ProductCode, ProductDescription FROM CustomerPackingMark where ProductNo='" + ProductNO + "' and LogoType='" + (PackingMarks.Trim().Contains("-") == true ? PackingMarks.Trim().Split('-')[1].Trim().ToUpper() : PackingMarks.Trim()) + "'";
                    cmd = new SqlCommand(selectCmd, conn);
                    reader = cmd.ExecuteReader();
                    if (reader.Read())
                    {
                        //載入客戶產品名稱
                        oSheet.Cells[1, 7] = reader.GetString(1);

                        //載入客戶產品型號
                        oSheet.Cells[2, 7] = reader.GetString(0);
                    }
                    reader.Close();

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
                else if (Client.ToUpper().StartsWith("AIR TEC") == true)
                {//20190314 AIR TEC 1.55L 增加Country of Origin : Taiwan 字樣
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

                conn.Close();

                //if (StorageStatus == "N")//20190212
                {
                    //預設位子在X:680,Y:155
                    //預設QRCODE圖片大小250*250

                    //插入圖片
                    int picX = 730, picY = 185;
                    string picadd = @"C:\QRCode\";
                    myConnection = new SqlConnection(myConnectionString);
                    selectCmd = "SELECT  * FROM [ShippingHead] where [ListDate]='" + ListDateListBox.SelectedItem + "' and [ProductName]='" + ProductComboBox.Text + "' and [vchBoxs]='" + BoxsListBox.SelectedItem + "'";
                    conn = new SqlConnection(myConnectionString);
                    conn.Open();
                    cmd = new SqlCommand(selectCmd, conn);
                    reader = cmd.ExecuteReader();
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
            else if (Aboxof == "36")
            {
                //載入嘜頭資料
                myConnection = new SqlConnection(myConnectionString);
                selectCmd = "SELECT  [Client],isnull([CustomerPO],''),isnull([CustomerProductName],''),isnull([CustomerProductNo],''),[vchBoxs],isnull(PalletNo,'') FROM [ShippingHead] where [ListDate]='" + ListDateListBox.SelectedItem + "' and [ProductName]='" + ProductComboBox.SelectedItem + "'and [vchBoxs]='" + BoxsListBox.SelectedItem + "'";
                conn = new SqlConnection(myConnectionString);
                conn.Open();
                cmd = new SqlCommand(selectCmd, conn);
                reader = cmd.ExecuteReader();
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
                    else if (Client.ToUpper().StartsWith("EMB"))
                    {
                        oSheet.Shapes.AddPicture(Application.StartupPath + @".\LOGO_EMB.png", Microsoft.Office.Core.MsoTriState.msoFalse,
                                            Microsoft.Office.Core.MsoTriState.msoTrue, 2, 17, 212, 125);
                    }
                }
                reader.Close();
                conn.Close();
                string FirstCNO = "";
                //載入嘜頭氣瓶序號位子
                myConnection = new SqlConnection(myConnectionString);
                selectCmd = "SELECT  * FROM [ShippingBody] where [ListDate]='" + ListDateListBox.SelectedItem + "' and [ProductName]='" + ProductComboBox.SelectedItem + "'and [WhereBox]='" + BoxsListBox.SelectedItem + "'";
                conn = new SqlConnection(myConnectionString);
                conn.Open();
                cmd = new SqlCommand(selectCmd, conn);
                reader = cmd.ExecuteReader();
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
                reader.Close();

                if ((Client == "SGA" || Client == "Airtanks Limited") && PackingMarks.Trim().CompareTo("SGA-GLADIATAIR") == 0)
                {
                    string ProductNO = "";
                    //該客戶要其自己的logo  PartNo   Part Description
                    selectCmd = "SELECT  Product_NO FROM MSNBody,Manufacturing where [CylinderNo]='" + FirstCNO + "' and vchManufacturingNo=  Manufacturing_NO";
                    cmd = new SqlCommand(selectCmd, conn);
                    reader = cmd.ExecuteReader();
                    if (reader.Read())
                    {
                        ProductNO = reader.GetValue(0).ToString();
                    }
                    reader.Close();

                    selectCmd = "SELECT  ProductCode, ProductDescription FROM CustomerPackingMark where ProductNo='" + ProductNO + "' and LogoType='" + (PackingMarks.Trim().Contains("-") == true ? PackingMarks.Trim().Split('-')[1].Trim().ToUpper() : PackingMarks.Trim()) + "'";
                    cmd = new SqlCommand(selectCmd, conn);
                    reader = cmd.ExecuteReader();
                    if (reader.Read())
                    {
                        //載入客戶產品名稱
                        oSheet.Cells[1, 8] = reader.GetString(1);

                        //載入客戶產品型號
                        oSheet.Cells[2, 8] = reader.GetString(0);
                    }
                    reader.Close();

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

                conn.Close();

                //if (StorageStatus == "N")//20190212
                {
                    //預設位子在X:680,Y:155
                    //預設QRCODE圖片大小250*250

                    //插入二維條碼
                    int picX = 750, picY = 179;
                    string picadd = @"C:\QRCode\";
                    myConnection = new SqlConnection(myConnectionString);
                    selectCmd = "SELECT  * FROM [ShippingHead] where [ListDate]='" + ListDateListBox.SelectedItem + "' and [ProductName]='" + ProductComboBox.Text + "' and [vchBoxs]='" + BoxsListBox.SelectedItem + "'";
                    conn = new SqlConnection(myConnectionString);
                    conn.Open();
                    cmd = new SqlCommand(selectCmd, conn);
                    reader = cmd.ExecuteReader();
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
            else if (Aboxof == "40")
            {   //載入嘜頭資料
                myConnection = new SqlConnection(myConnectionString);
                selectCmd = "SELECT  [Client],isnull([CustomerPO],''),isnull([CustomerProductName],''),isnull([CustomerProductNo],''),[vchBoxs],isnull(PalletNo,'') FROM [ShippingHead] where [ListDate]='" + ListDateListBox.SelectedItem + "' and [ProductName]='" + ProductComboBox.SelectedItem + "'and [vchBoxs]='" + BoxsListBox.SelectedItem + "'";
                conn = new SqlConnection(myConnectionString);
                conn.Open();
                cmd = new SqlCommand(selectCmd, conn);
                reader = cmd.ExecuteReader();
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
                    else if (Client.ToUpper().StartsWith("EMB"))
                    {
                        oSheet.Shapes.AddPicture(Application.StartupPath + @".\LOGO_EMB.png", Microsoft.Office.Core.MsoTriState.msoFalse,
                                            Microsoft.Office.Core.MsoTriState.msoTrue, 2, 17, 212, 125);
                    }
                }
                reader.Close();
                conn.Close();

                string FirstCNO = "";
                //載入嘜頭氣瓶序號位子
                myConnection = new SqlConnection(myConnectionString);
                selectCmd = "SELECT  * FROM [ShippingBody] where [ListDate]='" + ListDateListBox.SelectedItem + "' and [ProductName]='" + ProductComboBox.SelectedItem + "'and [WhereBox]='" + BoxsListBox.SelectedItem + "'";
                conn = new SqlConnection(myConnectionString);
                conn.Open();
                cmd = new SqlCommand(selectCmd, conn);
                reader = cmd.ExecuteReader();
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
                reader.Close();

                if ((Client == "SGA" || Client == "Airtanks Limited") && PackingMarks.Trim().CompareTo("SGA-GLADIATAIR") == 0)
                {
                    string ProductNO = "";
                    //該客戶要其自己的logo  PartNo   Part Description
                    selectCmd = "SELECT  Product_NO FROM MSNBody,Manufacturing where [CylinderNo]='" + FirstCNO + "' and vchManufacturingNo=  Manufacturing_NO";
                    cmd = new SqlCommand(selectCmd, conn);
                    reader = cmd.ExecuteReader();
                    if (reader.Read())
                    {
                        ProductNO = reader.GetValue(0).ToString();
                    }
                    reader.Close();

                    selectCmd = "SELECT  ProductCode, ProductDescription FROM CustomerPackingMark where ProductNo='" + ProductNO + "' and LogoType='" + (PackingMarks.Trim().Contains("-") == true ? PackingMarks.Trim().Split('-')[1].Trim().ToUpper() : PackingMarks.Trim()) + "'";
                    cmd = new SqlCommand(selectCmd, conn);
                    reader = cmd.ExecuteReader();
                    if (reader.Read())
                    {
                        //載入客戶產品名稱
                        oSheet.Cells[1, 8] = reader.GetString(1);

                        //載入客戶產品型號
                        oSheet.Cells[2, 8] = reader.GetString(0);
                    }
                    reader.Close();

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

                conn.Close();

                //if (StorageStatus == "N")//20190212
                {
                    //預設位子在X:680,Y:155
                    //預設QRCODE圖片大小250*250

                    //插入二維條碼
                    int picX = 680, picY = 180;
                    string picadd = @"C:\QRCode\";
                    myConnection = new SqlConnection(myConnectionString);
                    selectCmd = "SELECT  * FROM [ShippingHead] where [ListDate]='" + ListDateListBox.SelectedItem + "' and [ProductName]='" + ProductComboBox.Text + "' and [vchBoxs]='" + BoxsListBox.SelectedItem + "'";
                    conn = new SqlConnection(myConnectionString);
                    conn.Open();
                    cmd = new SqlCommand(selectCmd, conn);
                    reader = cmd.ExecuteReader();
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
            else if (Aboxof == "15")
            {
                string HowMuch = "";
                int Cumulative = 0;
                int Total = 0;

                //載入嘜頭資料
                myConnection = new SqlConnection(myConnectionString);
                selectCmd = "SELECT  * FROM [ShippingHead] where [ListDate]='" + ListDateListBox.SelectedItem + "' and [ProductName]='" + ProductComboBox.SelectedItem + "'";
                conn = new SqlConnection(myConnectionString);
                conn.Open();
                cmd = new SqlCommand(selectCmd, conn);
                reader = cmd.ExecuteReader();
                while (reader.Read())
                {
                    HowMuch = reader.GetString(4);
                    Cumulative++;
                }
                reader.Close();
                conn.Close();

                Total = Convert.ToInt32(HowMuch) * Cumulative;

                //載入嘜頭資料
                myConnection = new SqlConnection(myConnectionString);
                selectCmd = "SELECT  [Client],isnull([CustomerPO],''),isnull([CustomerProductName],''),isnull([CustomerProductNo],''),[vchBoxs],isnull(PalletNo,'') FROM [ShippingHead] where [ListDate]='" + ListDateListBox.SelectedItem + "' and [ProductName]='" + ProductComboBox.SelectedItem + "'and [vchBoxs]='" + BoxsListBox.SelectedItem + "'";
                conn = new SqlConnection(myConnectionString);
                conn.Open();
                cmd = new SqlCommand(selectCmd, conn);
                reader = cmd.ExecuteReader();
                while (reader.Read())
                {
                    Client = reader.GetString(0).Trim();
                    //載入客戶產品名稱
                    oSheet.Cells[1, 7] = reader.GetString(2);

                    //載入客戶產品型號
                    oSheet.Cells[2, 7] = reader.GetString(3);

                    //載入一箱幾隻
                    oSheet.Cells[4, 7] = Getcount;

                    //載入箱號
                    oSheet.Cells[9, 2] = reader.GetString(4);
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
                    else if (Client.ToUpper().StartsWith("EMB"))
                    {
                        oSheet.Shapes.AddPicture(Application.StartupPath + @".\LOGO_EMB.png", Microsoft.Office.Core.MsoTriState.msoFalse,
                                            Microsoft.Office.Core.MsoTriState.msoTrue, 2, 17, 212, 125);
                    }
                }
                reader.Close();
                conn.Close();

                //////////
                int serialnooneX = 7, serialnooneY = 209;
                string serialnooneadd = @"C:\SerialNoCode\";
                //////

                string FirstCNO = "";
                //載入嘜頭氣瓶序號位子
                myConnection = new SqlConnection(myConnectionString);
                selectCmd = "SELECT  * FROM [ShippingBody] where [ListDate]='" + ListDateListBox.SelectedItem + "' and [ProductName]='" + ProductComboBox.SelectedItem + "'and [WhereBox]='" + BoxsListBox.SelectedItem + "'";
                conn = new SqlConnection(myConnectionString);
                conn.Open();
                cmd = new SqlCommand(selectCmd, conn);
                reader = cmd.ExecuteReader();
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
                reader.Close();

                if ((Client == "SGA" || Client == "Airtanks Limited") && PackingMarks.Trim().CompareTo("SGA-GLADIATAIR") == 0)
                {
                    string ProductNO = "";
                    //該客戶要其自己的logo  PartNo   Part Description
                    selectCmd = "SELECT  Product_NO FROM MSNBody,Manufacturing where [CylinderNo]='" + FirstCNO + "' and vchManufacturingNo=  Manufacturing_NO";
                    cmd = new SqlCommand(selectCmd, conn);
                    reader = cmd.ExecuteReader();
                    if (reader.Read())
                    {
                        ProductNO = reader.GetValue(0).ToString();
                    }
                    reader.Close();

                    selectCmd = "SELECT  ProductCode, ProductDescription FROM CustomerPackingMark where ProductNo='" + ProductNO + "' and LogoType='" + (PackingMarks.Trim().Contains("-") == true ? PackingMarks.Trim().Split('-')[1].Trim().ToUpper() : PackingMarks.Trim()) + "'";
                    cmd = new SqlCommand(selectCmd, conn);
                    reader = cmd.ExecuteReader();
                    if (reader.Read())
                    {
                        //載入客戶產品名稱
                        oSheet.Cells[1, 7] = reader.GetString(1);

                        //載入客戶產品型號
                        oSheet.Cells[2, 7] = reader.GetString(0);
                    }
                    reader.Close();

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

                conn.Close();

                //if (StorageStatus == "N")//20190212
                {
                    //預設位子在X:680,Y:155
                    //預設QRCODE圖片大小250*250

                    //插入圖片
                    int picX = 732, picY = 187;
                    string picadd = @"C:\QRCode\";
                    myConnection = new SqlConnection(myConnectionString);
                    selectCmd = "SELECT  * FROM [ShippingHead] where [ListDate]='" + ListDateListBox.SelectedItem + "' and [ProductName]='" + ProductComboBox.Text + "' and [vchBoxs]='" + BoxsListBox.SelectedItem + "'";
                    conn = new SqlConnection(myConnectionString);
                    conn.Open();
                    cmd = new SqlCommand(selectCmd, conn);
                    reader = cmd.ExecuteReader();
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
            else if (Aboxof == "12")
            {
                string HowMuch = "";
                int Cumulative = 0;
                int Total = 0;

                //載入嘜頭資料
                myConnection = new SqlConnection(myConnectionString);
                selectCmd = "SELECT  * FROM [ShippingHead] where [ListDate]='" + ListDateListBox.SelectedItem + "' and [ProductName]='" + ProductComboBox.SelectedItem + "'";
                conn = new SqlConnection(myConnectionString);
                conn.Open();
                cmd = new SqlCommand(selectCmd, conn);
                reader = cmd.ExecuteReader();
                while (reader.Read())
                {
                    HowMuch = reader.GetString(4);
                    Cumulative++;
                }
                reader.Close();
                conn.Close();

                Total = Convert.ToInt32(HowMuch) * Cumulative;

                //載入嘜頭資料
                myConnection = new SqlConnection(myConnectionString);
                selectCmd = "SELECT  [Client],isnull([CustomerPO],''),isnull([CustomerProductName],''),isnull([CustomerProductNo],''),[vchBoxs],isnull(PalletNo,'') FROM [ShippingHead] where [ListDate]='" + ListDateListBox.SelectedItem + "' and [ProductName]='" + ProductComboBox.SelectedItem + "'and [vchBoxs]='" + BoxsListBox.SelectedItem + "'";
                conn = new SqlConnection(myConnectionString);
                conn.Open();
                cmd = new SqlCommand(selectCmd, conn);
                reader = cmd.ExecuteReader();
                while (reader.Read())
                {
                    Client = reader.GetString(0).Trim();
                    //載入客戶產品名稱
                    oSheet.Cells[1, 7] = reader.GetString(2);

                    //載入客戶產品型號
                    oSheet.Cells[2, 7] = reader.GetString(3);

                    //載入一箱幾隻
                    oSheet.Cells[4, 7] = Getcount;

                    //載入箱號
                    oSheet.Cells[12, 2] = reader.GetString(4);

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
                    else if (Client.ToUpper().StartsWith("EMB"))
                    {
                        oSheet.Shapes.AddPicture(Application.StartupPath + @".\LOGO_EMB.png", Microsoft.Office.Core.MsoTriState.msoFalse,
                                            Microsoft.Office.Core.MsoTriState.msoTrue, 2, 17, 212, 125);
                    }
                }
                reader.Close();
                conn.Close();

                //////////
                int serialnooneX = 10, serialnooneY = 212;
                string serialnooneadd = @"C:\SerialNoCode\";
                //////

                string FirstCNO = "";
                //載入嘜頭氣瓶序號位子
                myConnection = new SqlConnection(myConnectionString);
                selectCmd = "SELECT  * FROM [ShippingBody] where [ListDate]='" + ListDateListBox.SelectedItem + "' and [ProductName]='" + ProductComboBox.SelectedItem + "'and [WhereBox]='" + BoxsListBox.SelectedItem + "'";
                conn = new SqlConnection(myConnectionString);
                conn.Open();
                cmd = new SqlCommand(selectCmd, conn);
                reader = cmd.ExecuteReader();
                while (reader.Read())
                {
                    serialnooneX = 3; serialnooneY = 212;
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

                        case "9":
                            oSheet.Cells[10, 1] = reader.GetString(3);
                            break;

                        case "10":
                            oSheet.Cells[10, 3] = reader.GetString(3);
                            break;

                        case "11":
                            oSheet.Cells[10, 5] = reader.GetString(3);
                            break;

                        case "12":
                            oSheet.Cells[10, 7] = reader.GetString(3);
                            break;
                    }
                    serialnooneX = serialnooneX + ((Convert.ToInt32(reader.GetString(5)) + 3) % 4) * 157;
                    serialnooneY = serialnooneY + ((Convert.ToInt32(reader.GetString(5)) - 1) / 4) * 75;
                    oSheet.Shapes.AddPicture(serialnooneadd + reader.GetString(3) + ".png", Microsoft.Office.Core.MsoTriState.msoFalse,
                    Microsoft.Office.Core.MsoTriState.msoTrue, serialnooneX, serialnooneY, 45, 45);//, 130, 25);
                }
                reader.Close();

                if ((Client == "SGA" || Client == "Airtanks Limited") && PackingMarks.Trim().CompareTo("SGA-GLADIATAIR") == 0)
                {
                    string ProductNO = "";
                    //該客戶要其自己的logo  PartNo   Part Description
                    selectCmd = "SELECT  Product_NO FROM MSNBody,Manufacturing where [CylinderNo]='" + FirstCNO + "' and vchManufacturingNo=  Manufacturing_NO";
                    cmd = new SqlCommand(selectCmd, conn);
                    reader = cmd.ExecuteReader();
                    if (reader.Read())
                    {
                        ProductNO = reader.GetValue(0).ToString();
                    }
                    reader.Close();

                    selectCmd = "SELECT  ProductCode, ProductDescription FROM CustomerPackingMark where ProductNo='" + ProductNO + "' and LogoType='" + (PackingMarks.Trim().Contains("-") == true ? PackingMarks.Trim().Split('-')[1].Trim().ToUpper() : PackingMarks.Trim()) + "'";
                    cmd = new SqlCommand(selectCmd, conn);
                    reader = cmd.ExecuteReader();
                    if (reader.Read())
                    {
                        //載入客戶產品名稱
                        oSheet.Cells[1, 7] = reader.GetString(1);

                        //載入客戶產品型號
                        oSheet.Cells[2, 7] = reader.GetString(0);
                    }
                    reader.Close();

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
                conn.Close();

                //if (StorageStatus == "N")//20190212
                {
                    //預設位子在X:680,Y:155
                    //預設QRCODE圖片大小250*250

                    //插入圖片
                    int picX = 680, picY = 185;
                    string picadd = @"C:\QRCode\";
                    myConnection = new SqlConnection(myConnectionString);
                    selectCmd = "SELECT  * FROM [ShippingHead] where [ListDate]='" + ListDateListBox.SelectedItem + "' and [ProductName]='" + ProductComboBox.Text + "' and [vchBoxs]='" + BoxsListBox.SelectedItem + "'";
                    conn = new SqlConnection(myConnectionString);
                    conn.Open();
                    cmd = new SqlCommand(selectCmd, conn);
                    reader = cmd.ExecuteReader();
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
            else if (Aboxof == "8")
            {
                string HowMuch = "";
                int Cumulative = 0;
                int Total = 0;

                //載入嘜頭資料
                myConnection = new SqlConnection(myConnectionString);
                selectCmd = "SELECT  * FROM [ShippingHead] where [ListDate]='" + ListDateListBox.SelectedItem + "' and [ProductName]='" + ProductComboBox.SelectedItem + "'";
                conn = new SqlConnection(myConnectionString);
                conn.Open();
                cmd = new SqlCommand(selectCmd, conn);
                reader = cmd.ExecuteReader();
                while (reader.Read())
                {
                    HowMuch = reader.GetString(4);
                    Cumulative++;
                }
                reader.Close();
                conn.Close();

                Total = Convert.ToInt32(HowMuch) * Cumulative;

                //載入嘜頭資料
                myConnection = new SqlConnection(myConnectionString);
                selectCmd = "SELECT  [Client],isnull([CustomerPO],''),isnull([CustomerProductName],''),isnull([CustomerProductNo],''),[vchBoxs],isnull(PalletNo,'') FROM [ShippingHead] where [ListDate]='" + ListDateListBox.SelectedItem + "' and [ProductName]='" + ProductComboBox.SelectedItem + "'and [vchBoxs]='" + BoxsListBox.SelectedItem + "'";
                conn = new SqlConnection(myConnectionString);
                conn.Open();
                cmd = new SqlCommand(selectCmd, conn);
                reader = cmd.ExecuteReader();
                while (reader.Read())
                {
                    Client = reader.GetString(0).Trim();
                    //載入客戶產品名稱
                    oSheet.Cells[1, 7] = reader.GetString(2);

                    //載入客戶產品型號
                    oSheet.Cells[2, 7] = reader.GetString(3);

                    //載入一箱幾隻
                    oSheet.Cells[4, 7] = Getcount;

                    //載入箱號
                    oSheet.Cells[10, 2] = reader.GetString(4);
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
                    else if (Client.ToUpper().StartsWith("EMB"))
                    {
                        oSheet.Shapes.AddPicture(Application.StartupPath + @".\LOGO_EMB.png", Microsoft.Office.Core.MsoTriState.msoFalse,
                                            Microsoft.Office.Core.MsoTriState.msoTrue, 2, 17, 212, 125);
                    }
                }
                reader.Close();
                conn.Close();

                //////////
                int serialnooneX = 10, serialnooneY = 239;
                string serialnooneadd = @"C:\SerialNoCode\";
                string FirstCNO = "";
                //////
                //載入嘜頭氣瓶序號位子
                myConnection = new SqlConnection(myConnectionString);
                selectCmd = "SELECT  * FROM [ShippingBody] where [ListDate]='" + ListDateListBox.SelectedItem + "' and [ProductName]='" + ProductComboBox.SelectedItem + "'and [WhereBox]='" + BoxsListBox.SelectedItem + "'";
                conn = new SqlConnection(myConnectionString);
                conn.Open();
                cmd = new SqlCommand(selectCmd, conn);
                reader = cmd.ExecuteReader();
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
                reader.Close();

                if ((Client == "SGA" || Client == "Airtanks Limited") && PackingMarks.Trim().CompareTo("SGA-GLADIATAIR") == 0)
                {
                    string ProductNO = "";
                    //該客戶要其自己的logo  PartNo   Part Description
                    selectCmd = "SELECT  Product_NO FROM MSNBody,Manufacturing where [CylinderNo]='" + FirstCNO + "' and vchManufacturingNo=  Manufacturing_NO";
                    cmd = new SqlCommand(selectCmd, conn);
                    reader = cmd.ExecuteReader();
                    if (reader.Read())
                    {
                        ProductNO = reader.GetValue(0).ToString();
                    }
                    reader.Close();

                    selectCmd = "SELECT  ProductCode, ProductDescription FROM CustomerPackingMark where ProductNo='" + ProductNO + "' and LogoType='" + (PackingMarks.Trim().Contains("-") == true ? PackingMarks.Trim().Split('-')[1].Trim().ToUpper() : PackingMarks.Trim()) + "'";
                    cmd = new SqlCommand(selectCmd, conn);
                    reader = cmd.ExecuteReader();
                    if (reader.Read())
                    {
                        //載入客戶產品名稱
                        oSheet.Cells[1, 7] = reader.GetString(1);

                        //載入客戶產品型號
                        oSheet.Cells[2, 7] = reader.GetString(0);
                    }
                    reader.Close();

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

                conn.Close();

                //if (StorageStatus == "N")//20190212
                {

                    //預設位子在X:680,Y:155
                    //預設QRCODE圖片大小250*250

                    //插入圖片
                    int picX = 680, picY = 182;
                    string picadd = @"C:\QRCode\";
                    myConnection = new SqlConnection(myConnectionString);
                    selectCmd = "SELECT  * FROM [ShippingHead] where [ListDate]='" + ListDateListBox.SelectedItem + "' and [ProductName]='" + ProductComboBox.Text + "' and [vchBoxs]='" + BoxsListBox.SelectedItem + "'";
                    conn = new SqlConnection(myConnectionString);
                    conn.Open();
                    cmd = new SqlCommand(selectCmd, conn);
                    reader = cmd.ExecuteReader();
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
            else if (Aboxof == "16")
            {
                string HowMuch = "";
                int Cumulative = 0;
                int Total = 0;

                //載入嘜頭資料
                myConnection = new SqlConnection(myConnectionString);
                selectCmd = "SELECT  * FROM [ShippingHead] where [ListDate]='" + ListDateListBox.SelectedItem + "' and [ProductName]='" + ProductComboBox.SelectedItem + "'";
                conn = new SqlConnection(myConnectionString);
                conn.Open();
                cmd = new SqlCommand(selectCmd, conn);
                reader = cmd.ExecuteReader();
                while (reader.Read())
                {
                    HowMuch = reader.GetString(4);
                    Cumulative++;
                }
                reader.Close();
                conn.Close();

                Total = Convert.ToInt32(HowMuch) * Cumulative;

                //載入嘜頭資料
                myConnection = new SqlConnection(myConnectionString);
                selectCmd = "SELECT  [Client],isnull([CustomerPO],''),isnull([CustomerProductName],''),isnull([CustomerProductNo],''),[vchBoxs] ,isnull(PalletNo,'')FROM [ShippingHead] where [ListDate]='" + ListDateListBox.SelectedItem + "' and [ProductName]='" + ProductComboBox.SelectedItem + "'and [vchBoxs]='" + BoxsListBox.SelectedItem + "'";
                conn = new SqlConnection(myConnectionString);
                conn.Open();
                cmd = new SqlCommand(selectCmd, conn);
                reader = cmd.ExecuteReader();
                while (reader.Read())
                {
                    Client = reader.GetString(0).Trim();
                    //載入客戶產品名稱
                    oSheet.Cells[1, 7] = reader.GetString(2);

                    //載入客戶產品型號
                    oSheet.Cells[2, 7] = reader.GetString(3);

                    //載入一箱幾隻
                    oSheet.Cells[4, 7] = Getcount;

                    //載入箱號
                    oSheet.Cells[10, 2] = reader.GetString(4);
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
                    else if (Client.ToUpper().StartsWith("EMB"))
                    {
                        oSheet.Shapes.AddPicture(Application.StartupPath + @".\LOGO_EMB.png", Microsoft.Office.Core.MsoTriState.msoFalse,
                                            Microsoft.Office.Core.MsoTriState.msoTrue, 2, 17, 212, 125);
                    }
                }
                reader.Close();
                conn.Close();

                //////////
                int serialnooneX = 10, serialnooneY = 239;
                string serialnooneadd = @"C:\SerialNoCode\";

                //////
                string FirstCNO = "";
                //載入嘜頭氣瓶序號位子
                myConnection = new SqlConnection(myConnectionString);
                selectCmd = "SELECT  * FROM [ShippingBody] where [ListDate]='" + ListDateListBox.SelectedItem + "' and [ProductName]='" + ProductComboBox.SelectedItem + "'and [WhereBox]='" + BoxsListBox.SelectedItem + "'";
                conn = new SqlConnection(myConnectionString);
                conn.Open();
                cmd = new SqlCommand(selectCmd, conn);
                reader = cmd.ExecuteReader();
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
                reader.Close();

                if ((Client == "SGA" || Client == "Airtanks Limited") && PackingMarks.Trim().CompareTo("SGA-GLADIATAIR") == 0)
                {
                    string ProductNO = "";
                    //該客戶要其自己的logo  PartNo   Part Description
                    selectCmd = "SELECT  Product_NO FROM MSNBody,Manufacturing where [CylinderNo]='" + FirstCNO + "' and vchManufacturingNo=  Manufacturing_NO";
                    cmd = new SqlCommand(selectCmd, conn);
                    reader = cmd.ExecuteReader();
                    if (reader.Read())
                    {
                        ProductNO = reader.GetValue(0).ToString();
                    }
                    reader.Close();

                    selectCmd = "SELECT  ProductCode, ProductDescription FROM CustomerPackingMark where ProductNo='" + ProductNO + "' and LogoType='" + (PackingMarks.Trim().Contains("-") == true ? PackingMarks.Trim().Split('-')[1].Trim().ToUpper() : PackingMarks.Trim()) + "'";
                    cmd = new SqlCommand(selectCmd, conn);
                    reader = cmd.ExecuteReader();
                    if (reader.Read())
                    {
                        //載入客戶產品名稱
                        oSheet.Cells[1, 7] = reader.GetString(1);

                        //載入客戶產品型號
                        oSheet.Cells[2, 7] = reader.GetString(0);
                    }
                    reader.Close();

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

                conn.Close();

                //if (StorageStatus == "N")//20190212
                {

                    //預設位子在X:680,Y:155
                    //預設QRCODE圖片大小250*250

                    //插入圖片
                    int picX = 680, picY = 185;
                    string picadd = @"C:\QRCode\";
                    myConnection = new SqlConnection(myConnectionString);
                    selectCmd = "SELECT  * FROM [ShippingHead] where [ListDate]='" + ListDateListBox.SelectedItem + "' and [ProductName]='" + ProductComboBox.Text + "' and [vchBoxs]='" + BoxsListBox.SelectedItem + "'";
                    conn = new SqlConnection(myConnectionString);
                    conn.Open();
                    cmd = new SqlCommand(selectCmd, conn);
                    reader = cmd.ExecuteReader();
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
            else if (Aboxof == "10")
            {
                string HowMuch = "";
                int Cumulative = 0;
                int Total = 0;

                //載入嘜頭資料
                myConnection = new SqlConnection(myConnectionString);
                selectCmd = "SELECT  * FROM [ShippingHead] where [ListDate]='" + ListDateListBox.SelectedItem + "' and [ProductName]='" + ProductComboBox.SelectedItem + "'";
                conn = new SqlConnection(myConnectionString);
                conn.Open();
                cmd = new SqlCommand(selectCmd, conn);
                reader = cmd.ExecuteReader();
                while (reader.Read())
                {
                    HowMuch = reader.GetString(4);
                    Cumulative++;
                }
                reader.Close();
                conn.Close();

                Total = Convert.ToInt32(HowMuch) * Cumulative;

                //載入嘜頭資料
                myConnection = new SqlConnection(myConnectionString);
                selectCmd = "SELECT  [Client],isnull([CustomerPO],''),isnull([CustomerProductName],''),isnull([CustomerProductNo],''),[vchBoxs],isnull(PalletNo,'') FROM [ShippingHead] where [ListDate]='" + ListDateListBox.SelectedItem + "' and [ProductName]='" + ProductComboBox.SelectedItem + "'and [vchBoxs]='" + BoxsListBox.SelectedItem + "'";
                conn = new SqlConnection(myConnectionString);
                conn.Open();
                cmd = new SqlCommand(selectCmd, conn);
                reader = cmd.ExecuteReader();
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
                    else if (Client.ToUpper().StartsWith("EMB"))
                    {
                        oSheet.Shapes.AddPicture(Application.StartupPath + @".\LOGO_EMB.png", Microsoft.Office.Core.MsoTriState.msoFalse,
                                            Microsoft.Office.Core.MsoTriState.msoTrue, 2, 17, 212, 125);
                    }
                }
                reader.Close();
                conn.Close();

                //////////
                int serialnooneX = 10, serialnooneY = 239;
                string serialnooneadd = @"C:\SerialNoCode\";
                //////

                string FirstCNO = "";
                //載入嘜頭氣瓶序號位子
                myConnection = new SqlConnection(myConnectionString);
                selectCmd = "SELECT  * FROM [ShippingBody] where [ListDate]='" + ListDateListBox.SelectedItem + "' and [ProductName]='" + ProductComboBox.SelectedItem + "'and [WhereBox]='" + BoxsListBox.SelectedItem + "'";
                conn = new SqlConnection(myConnectionString);
                conn.Open();
                cmd = new SqlCommand(selectCmd, conn);
                reader = cmd.ExecuteReader();
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
                    serialnooneX = serialnooneX + ((Convert.ToInt32(reader.GetString(5)) + 3) % 5) * 127;
                    serialnooneY = serialnooneY + ((Convert.ToInt32(reader.GetString(5)) - 1) / 5) * 111;
                    oSheet.Shapes.AddPicture(serialnooneadd + reader.GetString(3) + ".png", Microsoft.Office.Core.MsoTriState.msoFalse,
                    Microsoft.Office.Core.MsoTriState.msoTrue, serialnooneX, serialnooneY, 60, 60);//, 110, 25);
                }
                reader.Close();

                if ((Client == "SGA" || Client == "Airtanks Limited") && PackingMarks.Trim().CompareTo("SGA-GLADIATAIR") == 0)
                {
                    string ProductNO = "";
                    //該客戶要其自己的logo  PartNo   Part Description
                    selectCmd = "SELECT  Product_NO FROM MSNBody,Manufacturing where [CylinderNo]='" + FirstCNO + "' and vchManufacturingNo=  Manufacturing_NO";
                    cmd = new SqlCommand(selectCmd, conn);
                    reader = cmd.ExecuteReader();
                    if (reader.Read())
                    {
                        ProductNO = reader.GetValue(0).ToString();
                    }
                    reader.Close();

                    selectCmd = "SELECT  ProductCode, ProductDescription FROM CustomerPackingMark where ProductNo='" + ProductNO + "' and LogoType='" + (PackingMarks.Trim().Contains("-") == true ? PackingMarks.Trim().Split('-')[1].Trim().ToUpper() : PackingMarks.Trim()) + "'";
                    cmd = new SqlCommand(selectCmd, conn);
                    reader = cmd.ExecuteReader();
                    if (reader.Read())
                    {
                        //載入客戶產品名稱
                        oSheet.Cells[1, 8] = reader.GetString(1);

                        //載入客戶產品型號
                        oSheet.Cells[2, 8] = reader.GetString(0);
                    }
                    reader.Close();

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

                conn.Close();

                //if (StorageStatus == "N")//20190212
                {
                    //預設位子在X:680,Y:155
                    //預設QRCODE圖片大小250*250

                    //插入圖片
                    int picX = 680, picY = 185;
                    string picadd = @"C:\QRCode\";
                    myConnection = new SqlConnection(myConnectionString);
                    selectCmd = "SELECT  * FROM [ShippingHead] where [ListDate]='" + ListDateListBox.SelectedItem + "' and [ProductName]='" + ProductComboBox.Text + "' and [vchBoxs]='" + BoxsListBox.SelectedItem + "'";
                    conn = new SqlConnection(myConnectionString);
                    conn.Open();
                    cmd = new SqlCommand(selectCmd, conn);
                    reader = cmd.ExecuteReader();
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
            else if (Aboxof == "25")
            {
                //載入嘜頭資料
                myConnection = new SqlConnection(myConnectionString);
                selectCmd = "SELECT  [Client],isnull([CustomerPO],''),isnull([CustomerProductName],''),isnull([CustomerProductNo],''),[vchBoxs],isnull(PalletNo,'') FROM [ShippingHead] where [ListDate]='" + ListDateListBox.SelectedItem + "' and [ProductName]='" + ProductComboBox.SelectedItem + "'and [vchBoxs]='" + BoxsListBox.SelectedItem + "'";
                conn = new SqlConnection(myConnectionString);
                conn.Open();
                cmd = new SqlCommand(selectCmd, conn);
                reader = cmd.ExecuteReader();
                while (reader.Read())
                {
                    Client = reader.GetString(0).Trim();
                    //載入客戶產品名稱
                    oSheet.Cells[1, 7] = reader.GetString(2);

                    //載入客戶產品型號
                    oSheet.Cells[2, 7] = reader.GetString(3);

                    //載入一箱幾隻
                    oSheet.Cells[4, 7] = Getcount;

                    //載入箱號
                    oSheet.Cells[11, 2] = reader.GetString(4);
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
                    else if (Client.ToUpper().StartsWith("EMB"))
                    {
                        oSheet.Shapes.AddPicture(Application.StartupPath + @".\LOGO_EMB.png", Microsoft.Office.Core.MsoTriState.msoFalse,
                                            Microsoft.Office.Core.MsoTriState.msoTrue, 2, 17, 212, 125);
                    }
                }
                reader.Close();
                conn.Close();

                //////////
                int serialnooneX = 8, serialnooneY = 192;
                string serialnooneadd = @"C:\SerialNoCode\";
                string FirstCNO = "";
                //////
                
                //載入嘜頭氣瓶序號位子
                myConnection = new SqlConnection(myConnectionString);
                selectCmd = "SELECT  * FROM [ShippingBody] where [ListDate]='" + ListDateListBox.SelectedItem + "' and [ProductName]='" + ProductComboBox.SelectedItem + "'and [WhereBox]='" + BoxsListBox.SelectedItem + "'";
                conn = new SqlConnection(myConnectionString);
                conn.Open();
                cmd = new SqlCommand(selectCmd, conn);
                reader = cmd.ExecuteReader();
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
                reader.Close();

                if ((Client == "SGA" || Client == "Airtanks Limited") && PackingMarks.Trim().CompareTo("SGA-GLADIATAIR") == 0)
                {
                    string ProductNO = "";
                    //該客戶要其自己的logo  PartNo   Part Description
                    selectCmd = "SELECT  Product_NO FROM MSNBody,Manufacturing where [CylinderNo]='" + FirstCNO + "' and vchManufacturingNo=  Manufacturing_NO";
                    cmd = new SqlCommand(selectCmd, conn);
                    reader = cmd.ExecuteReader();
                    if (reader.Read())
                    {
                        ProductNO = reader.GetValue(0).ToString();
                    }
                    reader.Close();

                    selectCmd = "SELECT  ProductCode, ProductDescription FROM CustomerPackingMark where ProductNo='" + ProductNO + "' and LogoType='" + (PackingMarks.Trim().Contains("-") == true ? PackingMarks.Trim().Split('-')[1].Trim().ToUpper() : PackingMarks.Trim()) + "'";
                    cmd = new SqlCommand(selectCmd, conn);
                    reader = cmd.ExecuteReader();
                    if (reader.Read())
                    {
                        //載入客戶產品名稱
                        oSheet.Cells[1, 7] = reader.GetString(1);

                        //載入客戶產品型號
                        oSheet.Cells[2, 7] = reader.GetString(0);
                    }
                    reader.Close();

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

                conn.Close();

                //if (StorageStatus == "N")//20190212
                {

                    //預設位子在X:680,Y:155
                    //預設QRCODE圖片大小250*250

                    //插入二維條碼
                    int picX = 730, picY = 179;
                    string picadd = @"C:\QRCode\";
                    myConnection = new SqlConnection(myConnectionString);
                    selectCmd = "SELECT  * FROM [ShippingHead] where [ListDate]='" + ListDateListBox.SelectedItem + "' and [ProductName]='" + ProductComboBox.Text + "' and [vchBoxs]='" + BoxsListBox.SelectedItem + "'";
                    conn = new SqlConnection(myConnectionString);
                    conn.Open();
                    cmd = new SqlCommand(selectCmd, conn);
                    reader = cmd.ExecuteReader();
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
            else if (Aboxof == "30")
            {
                //載入嘜頭資料
                myConnection = new SqlConnection(myConnectionString);
                selectCmd = "SELECT  [Client],isnull([CustomerPO],''),isnull([CustomerProductName],''),isnull([CustomerProductNo],''),[vchBoxs],isnull(PalletNo,'') FROM [ShippingHead] where [ListDate]='" + ListDateListBox.SelectedItem + "' and [ProductName]='" + ProductComboBox.SelectedItem + "'and [vchBoxs]='" + BoxsListBox.SelectedItem + "'";
                conn = new SqlConnection(myConnectionString);
                conn.Open();
                cmd = new SqlCommand(selectCmd, conn);
                reader = cmd.ExecuteReader();
                while (reader.Read())
                {
                    Client = reader.GetString(0).Trim();
                    //載入客戶產品名稱
                    oSheet.Cells[1, 7] = reader.GetString(2);

                    //載入客戶產品型號
                    oSheet.Cells[2, 7] = reader.GetString(3);

                    //載入一箱幾隻
                    oSheet.Cells[4, 7] = Getcount;

                    //載入箱號
                    oSheet.Cells[12, 2] = reader.GetString(4);
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

                }
                reader.Close();
                conn.Close();

                //////

                string FirstCNO = "";
                //載入嘜頭氣瓶序號位子
                myConnection = new SqlConnection(myConnectionString);
                selectCmd = "SELECT  * FROM [ShippingBody] where [ListDate]='" + ListDateListBox.SelectedItem + "' and [ProductName]='" + ProductComboBox.SelectedItem + "'and [WhereBox]='" + BoxsListBox.SelectedItem + "'";
                conn = new SqlConnection(myConnectionString);
                conn.Open();
                cmd = new SqlCommand(selectCmd, conn);
                reader = cmd.ExecuteReader();
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
                reader.Close();

                if ((Client == "SGA" || Client == "Airtanks Limited") && PackingMarks.Trim().CompareTo("SGA-GLADIATAIR") == 0)
                {
                    string ProductNO = "";
                    //該客戶要其自己的logo  PartNo   Part Description
                    selectCmd = "SELECT  Product_NO FROM MSNBody,Manufacturing where [CylinderNo]='" + FirstCNO + "' and vchManufacturingNo=  Manufacturing_NO";
                    cmd = new SqlCommand(selectCmd, conn);
                    reader = cmd.ExecuteReader();
                    if (reader.Read())
                    {
                        ProductNO = reader.GetValue(0).ToString();
                    }
                    reader.Close();

                    selectCmd = "SELECT  ProductCode, ProductDescription FROM CustomerPackingMark where ProductNo='" + ProductNO + "' and LogoType='" + (PackingMarks.Trim().Contains("-") == true ? PackingMarks.Trim().Split('-')[1].Trim().ToUpper() : PackingMarks.Trim()) + "'";
                    cmd = new SqlCommand(selectCmd, conn);
                    reader = cmd.ExecuteReader();
                    if (reader.Read())
                    {
                        //載入客戶產品名稱
                        oSheet.Cells[1, 7] = reader.GetString(1);

                        //載入客戶產品型號
                        oSheet.Cells[2, 7] = reader.GetString(0);
                    }
                    reader.Close();

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
                conn.Close();

                //if (StorageStatus == "N")//20190212
                {
                    //預設位子在X:680,Y:155
                    //預設QRCODE圖片大小250*250

                    //插入二維條碼
                    int picX = 730, picY = 179;
                    string picadd = @"C:\QRCode\";
                    myConnection = new SqlConnection(myConnectionString);
                    selectCmd = "SELECT  * FROM [ShippingHead] where [ListDate]='" + ListDateListBox.SelectedItem + "' and [ProductName]='" + ProductComboBox.Text + "' and [vchBoxs]='" + BoxsListBox.SelectedItem + "'";
                    conn = new SqlConnection(myConnectionString);
                    conn.Open();
                    cmd = new SqlCommand(selectCmd, conn);
                    reader = cmd.ExecuteReader();
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
            else if (Aboxof == "117")
            {
                //載入嘜頭資料
                myConnection = new SqlConnection(myConnectionString);
                selectCmd = "SELECT  [Client],isnull([CustomerPO],''),isnull([CustomerProductName],''),isnull([CustomerProductNo],''),[vchBoxs],isnull(PalletNo,'') FROM [ShippingHead] where [ListDate]='" + ListDateListBox.SelectedItem + "' and [ProductName]='" + ProductComboBox.SelectedItem + "'and [vchBoxs]='" + BoxsListBox.SelectedItem + "'";
                conn = new SqlConnection(myConnectionString);
                conn.Open();
                cmd = new SqlCommand(selectCmd, conn);
                reader = cmd.ExecuteReader();
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
                    else if (Client.ToUpper().StartsWith("EMB"))
                    {
                        oSheet.Shapes.AddPicture(Application.StartupPath + @".\LOGO_EMB.png", Microsoft.Office.Core.MsoTriState.msoFalse,
                                            Microsoft.Office.Core.MsoTriState.msoTrue, 2, 17, 212, 125);
                    }
                }
                reader.Close();
                conn.Close();

                string FirstCNO = "";
                //載入嘜頭氣瓶序號位子
                myConnection = new SqlConnection(myConnectionString);
                selectCmd = "SELECT  * FROM [ShippingBody] where [ListDate]='" + ListDateListBox.SelectedItem + "' and [ProductName]='" + ProductComboBox.SelectedItem + "'and [WhereBox]='" + BoxsListBox.SelectedItem + "'";
                conn = new SqlConnection(myConnectionString);
                conn.Open();
                cmd = new SqlCommand(selectCmd, conn);
                reader = cmd.ExecuteReader();
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
                reader.Close();

                if ((Client == "SGA" || Client == "Airtanks Limited") && PackingMarks.Trim().CompareTo("SGA-GLADIATAIR") == 0)
                {
                    string ProductNO = "";
                    //該客戶要其自己的logo  PartNo   Part Description
                    selectCmd = "SELECT  Product_NO FROM MSNBody,Manufacturing where [CylinderNo]='" + FirstCNO + "' and vchManufacturingNo=  Manufacturing_NO";
                    cmd = new SqlCommand(selectCmd, conn);
                    reader = cmd.ExecuteReader();
                    if (reader.Read())
                    {
                        ProductNO = reader.GetValue(0).ToString();
                    }
                    reader.Close();

                    selectCmd = "SELECT  ProductCode, ProductDescription FROM CustomerPackingMark where ProductNo='" + ProductNO + "' and LogoType='" + (PackingMarks.Trim().Contains("-") == true ? PackingMarks.Trim().Split('-')[1].Trim().ToUpper() : PackingMarks.Trim()) + "'";
                    cmd = new SqlCommand(selectCmd, conn);
                    reader = cmd.ExecuteReader();
                    if (reader.Read())
                    {
                        //載入客戶產品名稱
                        oSheet.Cells[1, 9] = reader.GetString(1);

                        //載入客戶產品型號
                        oSheet.Cells[2, 9] = reader.GetString(0);
                    }
                    reader.Close();

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
                conn.Close();

                //Aboxof == "117"其資料太長，造成QR code 無法全部紀錄，僅序號最多41組
                //if (StorageStatus == "N")
                //{

                //    //預設位子在X:680,Y:155
                //    //預設QRCODE圖片大小250*250

                //    //插入二維條碼

                //    string picadd = @"C:\QRCode\";
                //    myConnection = new SqlConnection(myConnectionString);
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
                myConnection = new SqlConnection(myConnectionString);
                selectCmd = "SELECT  [Client],isnull([CustomerPO],''),isnull([CustomerProductName],''),isnull([CustomerProductNo],''),[vchBoxs] ,isnull(PalletNo,'') FROM [ShippingHead] where [ListDate]='" + ListDateListBox.SelectedItem + "' and [ProductName]='" + ProductComboBox.SelectedItem + "'and [vchBoxs]='" + BoxsListBox.SelectedItem + "'";
                conn = new SqlConnection(myConnectionString);
                conn.Open();
                cmd = new SqlCommand(selectCmd, conn);
                reader = cmd.ExecuteReader();
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
                    else if (Client.ToUpper().StartsWith("EMB"))
                    {
                        oSheet.Shapes.AddPicture(Application.StartupPath + @".\LOGO_EMB.png", Microsoft.Office.Core.MsoTriState.msoFalse,
                                            Microsoft.Office.Core.MsoTriState.msoTrue, 2, 17, 212, 125);
                    }
                }
                reader.Close();
                conn.Close();
                string FirstCNO = "";
                //載入嘜頭氣瓶序號位子
                myConnection = new SqlConnection(myConnectionString);
                selectCmd = "SELECT  * FROM [ShippingBody] where [ListDate]='" + ListDateListBox.SelectedItem + "' and [ProductName]='" + ProductComboBox.SelectedItem + "'and [WhereBox]='" + BoxsListBox.SelectedItem + "'";
                conn = new SqlConnection(myConnectionString);
                conn.Open();
                cmd = new SqlCommand(selectCmd, conn);
                reader = cmd.ExecuteReader();
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
                reader.Close();

                if ((Client == "SGA" || Client == "Airtanks Limited") && PackingMarks.Trim().CompareTo("SGA-GLADIATAIR") == 0)
                {
                    string ProductNO = "";
                    //該客戶要其自己的logo  PartNo   Part Description
                    selectCmd = "SELECT  Product_NO FROM MSNBody,Manufacturing where [CylinderNo]='" + FirstCNO + "' and vchManufacturingNo=  Manufacturing_NO";
                    cmd = new SqlCommand(selectCmd, conn);
                    reader = cmd.ExecuteReader();
                    if (reader.Read())
                    {
                        ProductNO = reader.GetValue(0).ToString();
                    }
                    reader.Close();

                    selectCmd = "SELECT  ProductCode, ProductDescription FROM CustomerPackingMark where ProductNo='" + ProductNO + "' and LogoType='" + (PackingMarks.Trim().Contains("-") == true ? PackingMarks.Trim().Split('-')[1].Trim().ToUpper() : PackingMarks.Trim()) + "'";
                    cmd = new SqlCommand(selectCmd, conn);
                    reader = cmd.ExecuteReader();
                    if (reader.Read())
                    {
                        //載入客戶產品名稱
                        oSheet.Cells[1, 9] = reader.GetString(1);

                        //載入客戶產品型號
                        oSheet.Cells[2, 9] = reader.GetString(0);
                    }
                    reader.Close();

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

                conn.Close();
            }
            if (Aboxof == "1")
            {
                //載入嘜頭資料
                myConnection = new SqlConnection(myConnectionString);
                selectCmd = "SELECT  [Client],isnull([CustomerPO],''),isnull([CustomerProductName],''),isnull([CustomerProductNo],''),[vchBoxs],isnull(PalletNo,'') FROM [ShippingHead] where [ListDate]='" + ListDateListBox.SelectedItem + "' and [ProductName]='" + ProductComboBox.SelectedItem + "'and [vchBoxs]='" + BoxsListBox.SelectedItem + "'";
                conn = new SqlConnection(myConnectionString);
                conn.Open();
                cmd = new SqlCommand(selectCmd, conn);
                reader = cmd.ExecuteReader();
                while (reader.Read())
                {
                    Client = reader.GetString(0).Trim();
                    //載入客戶產品名稱
                    oSheet.Cells[1, 7] = reader.GetString(2);

                    //載入客戶產品型號
                    oSheet.Cells[2, 7] = reader.GetString(3);

                    //載入一箱幾隻
                    oSheet.Cells[4, 7] = Getcount;

                    //載入箱號
                    //oSheet.Cells[11, 3] = reader.GetString(4);

                    //// if (StorageStatus == "N")//20190213
                    {
                        //載入客戶名稱
                        oSheet.Cells[3, 7] = reader.GetString(0);

                        ////載入訂單編號(PO)
                        //oSheet.Cells[5, 13] = reader.GetString(1);

                        //載入箱號
                        oSheet.Cells[10, 2] = reader.GetString(4);
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
                    else if (Client.ToUpper().StartsWith("EMB"))
                    {
                        oSheet.Shapes.AddPicture(Application.StartupPath + @".\LOGO_EMB.png", Microsoft.Office.Core.MsoTriState.msoFalse,
                                            Microsoft.Office.Core.MsoTriState.msoTrue, 2, 17, 212, 125);
                    }
                }
                reader.Close();
                conn.Close();

                //////
                string FirstCNO = "";
                //載入嘜頭氣瓶序號位子
                myConnection = new SqlConnection(myConnectionString);
                selectCmd = "SELECT  * FROM [ShippingBody] where [ListDate]='" + ListDateListBox.SelectedItem + "' and [ProductName]='" + ProductComboBox.SelectedItem + "'and [WhereBox]='" + BoxsListBox.SelectedItem + "'";
                conn = new SqlConnection(myConnectionString);
                conn.Open();
                cmd = new SqlCommand(selectCmd, conn);
                reader = cmd.ExecuteReader();
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
                reader.Close();

                if ((Client == "SGA" || Client == "Airtanks Limited") && PackingMarks.Trim().CompareTo("SGA-GLADIATAIR") == 0)
                {
                    string ProductNO = "";
                    //該客戶要其自己的logo  PartNo   Part Description
                    selectCmd = "SELECT  Product_NO FROM MSNBody,Manufacturing where [CylinderNo]='" + FirstCNO + "' and vchManufacturingNo=  Manufacturing_NO";
                    cmd = new SqlCommand(selectCmd, conn);
                    reader = cmd.ExecuteReader();
                    if (reader.Read())
                    {
                        ProductNO = reader.GetValue(0).ToString();
                    }
                    reader.Close();

                    selectCmd = "SELECT  ProductCode, ProductDescription FROM CustomerPackingMark where ProductNo='" + ProductNO + "' and LogoType='" + (PackingMarks.Trim().Contains("-") == true ? PackingMarks.Trim().Split('-')[1].Trim().ToUpper() : PackingMarks.Trim()) + "'";
                    cmd = new SqlCommand(selectCmd, conn);
                    reader = cmd.ExecuteReader();
                    if (reader.Read())
                    {
                        //載入客戶產品名稱
                        oSheet.Cells[1, 7] = reader.GetString(1);

                        //載入客戶產品型號
                        oSheet.Cells[2, 7] = reader.GetString(0);
                    }
                    reader.Close();

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

                conn.Close();

                //if (StorageStatus == "N")
                //{

                //    //預設位子在X:680,Y:155
                //    //預設QRCODE圖片大小250*250

                //    //插入二維條碼
                //    int picX = 730, picY = 179;
                //    string picadd = @"C:\QRCode\";
                //    myConnection = new SqlConnection(myConnectionString);
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
                myConnection = new SqlConnection(myConnectionString);
                selectCmd = "SELECT  * FROM [ShippingHead] where [ListDate]='" + ListDateListBox.SelectedItem + "' and [ProductName]='" + ProductComboBox.SelectedItem + "'";
                conn = new SqlConnection(myConnectionString);
                conn.Open();
                cmd = new SqlCommand(selectCmd, conn);
                reader = cmd.ExecuteReader();
                while (reader.Read())
                {
                    HowMuch = reader.GetString(4);
                    Cumulative++;
                }
                reader.Close();
                conn.Close();

                Total = Convert.ToInt32(HowMuch) * Cumulative;

                //載入嘜頭資料
                myConnection = new SqlConnection(myConnectionString);
                selectCmd = "SELECT  [Client],isnull([CustomerPO],''),isnull([CustomerProductName],''),isnull([CustomerProductNo],''),[vchBoxs],isnull(PalletNo,'') FROM [ShippingHead] where [ListDate]='" + ListDateListBox.SelectedItem + "' and [ProductName]='" + ProductComboBox.SelectedItem + "'and [vchBoxs]='" + BoxsListBox.SelectedItem + "'";
                conn = new SqlConnection(myConnectionString);
                conn.Open();
                cmd = new SqlCommand(selectCmd, conn);
                reader = cmd.ExecuteReader();
                while (reader.Read())
                {
                    Client = reader.GetString(0).Trim();
                    //載入客戶產品名稱
                    oSheet.Cells[1, 7] = reader.GetString(2);

                    //載入客戶產品型號
                    oSheet.Cells[2, 7] = reader.GetString(3);

                    //載入一箱幾隻
                    oSheet.Cells[4, 7] = Getcount;

                    //載入箱號
                    oSheet.Cells[10, 2] = reader.GetString(4);
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
                    else if (Client.ToUpper().StartsWith("EMB"))
                    {
                        oSheet.Shapes.AddPicture(Application.StartupPath + @".\LOGO_EMB.png", Microsoft.Office.Core.MsoTriState.msoFalse,
                                            Microsoft.Office.Core.MsoTriState.msoTrue, 2, 17, 212, 125);
                    }
                }
                reader.Close();
                conn.Close();

                //////////
                int serialnooneX = 10, serialnooneY = 239;
                string serialnooneadd = @"C:\SerialNoCode\";

                //////
                string FirstCNO = "";
                //載入嘜頭氣瓶序號位子
                myConnection = new SqlConnection(myConnectionString);
                selectCmd = "SELECT  * FROM [ShippingBody] where [ListDate]='" + ListDateListBox.SelectedItem + "' and [ProductName]='" + ProductComboBox.SelectedItem + "'and [WhereBox]='" + BoxsListBox.SelectedItem + "'";
                conn = new SqlConnection(myConnectionString);
                conn.Open();
                cmd = new SqlCommand(selectCmd, conn);
                reader = cmd.ExecuteReader();
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
                reader.Close();

                if ((Client == "SGA" || Client == "Airtanks Limited") && PackingMarks.Trim().CompareTo("SGA-GLADIATAIR") == 0)
                {
                    string ProductNO = "";
                    //該客戶要其自己的logo  PartNo   Part Description
                    selectCmd = "SELECT  Product_NO FROM MSNBody,Manufacturing where [CylinderNo]='" + FirstCNO + "' and vchManufacturingNo=  Manufacturing_NO";
                    cmd = new SqlCommand(selectCmd, conn);
                    reader = cmd.ExecuteReader();
                    if (reader.Read())
                    {
                        ProductNO = reader.GetValue(0).ToString();
                    }
                    reader.Close();

                    selectCmd = "SELECT  ProductCode, ProductDescription FROM CustomerPackingMark where ProductNo='" + ProductNO + "' and LogoType='" + (PackingMarks.Trim().Contains("-") == true ? PackingMarks.Trim().Split('-')[1].Trim().ToUpper() : PackingMarks.Trim()) + "'";
                    cmd = new SqlCommand(selectCmd, conn);
                    reader = cmd.ExecuteReader();
                    if (reader.Read())
                    {
                        //載入客戶產品名稱
                        oSheet.Cells[1, 7] = reader.GetString(1);

                        //載入客戶產品型號
                        oSheet.Cells[2, 7] = reader.GetString(0);
                    }
                    reader.Close();

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

                conn.Close();

                //if (StorageStatus == "N")//20190212
                {

                    //預設位子在X:680,Y:155
                    //預設QRCODE圖片大小250*250

                    //插入圖片
                    int picX = 680, picY = 185;
                    string picadd = @"C:\QRCode\";
                    myConnection = new SqlConnection(myConnectionString);
                    selectCmd = "SELECT  * FROM [ShippingHead] where [ListDate]='" + ListDateListBox.SelectedItem + "' and [ProductName]='" + ProductComboBox.Text + "' and [vchBoxs]='" + BoxsListBox.SelectedItem + "'";
                    conn = new SqlConnection(myConnectionString);
                    conn.Open();
                    cmd = new SqlCommand(selectCmd, conn);
                    reader = cmd.ExecuteReader();
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
            else if (Aboxof == "2")
            {
                string HowMuch = "";
                int Cumulative = 0;
                int Total = 0;

                //載入嘜頭資料
                myConnection = new SqlConnection(myConnectionString);
                selectCmd = "SELECT  * FROM [ShippingHead] where [ListDate]='" + ListDateListBox.SelectedItem + "' and [ProductName]='" + ProductComboBox.SelectedItem + "'";
                conn = new SqlConnection(myConnectionString);
                conn.Open();
                cmd = new SqlCommand(selectCmd, conn);
                reader = cmd.ExecuteReader();
                while (reader.Read())
                {
                    HowMuch = reader.GetString(4);
                    Cumulative++;
                }
                reader.Close();
                conn.Close();

                Total = Convert.ToInt32(HowMuch) * Cumulative;

                //載入嘜頭資料
                myConnection = new SqlConnection(myConnectionString);
                selectCmd = "SELECT  [Client],isnull([CustomerPO],''),isnull([CustomerProductName],''),isnull([CustomerProductNo],''),[vchBoxs],isnull(PalletNo,'') FROM [ShippingHead] where [ListDate]='" + ListDateListBox.SelectedItem + "' and [ProductName]='" + ProductComboBox.SelectedItem + "'and [vchBoxs]='" + BoxsListBox.SelectedItem + "'";
                conn = new SqlConnection(myConnectionString);
                conn.Open();
                cmd = new SqlCommand(selectCmd, conn);
                reader = cmd.ExecuteReader();
                while (reader.Read())
                {
                    Client = reader.GetString(0).Trim();
                    //載入客戶產品名稱
                    oSheet.Cells[1, 7] = reader.GetString(2);

                    //載入客戶產品型號
                    oSheet.Cells[2, 7] = reader.GetString(3);

                    //載入一箱幾隻
                    oSheet.Cells[4, 7] = Getcount;

                    //載入箱號
                    oSheet.Cells[10, 2] = reader.GetString(4);
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
                    else if (Client.ToUpper().StartsWith("EMB"))
                    {
                        oSheet.Shapes.AddPicture(Application.StartupPath + @".\LOGO_EMB.png", Microsoft.Office.Core.MsoTriState.msoFalse,
                                            Microsoft.Office.Core.MsoTriState.msoTrue, 2, 17, 212, 125);
                    }
                }
                reader.Close();
                conn.Close();

                //////////
                int serialnooneX = 10, serialnooneY = 239;
                string serialnooneadd = @"C:\SerialNoCode\";

                //////
                string FirstCNO = "";
                //載入嘜頭氣瓶序號位子
                myConnection = new SqlConnection(myConnectionString);
                selectCmd = "SELECT  * FROM [ShippingBody] where [ListDate]='" + ListDateListBox.SelectedItem + "' and [ProductName]='" + ProductComboBox.SelectedItem + "'and [WhereBox]='" + BoxsListBox.SelectedItem + "'";
                conn = new SqlConnection(myConnectionString);
                conn.Open();
                cmd = new SqlCommand(selectCmd, conn);
                reader = cmd.ExecuteReader();
                while (reader.Read())
                {
                    serialnooneX = 3; serialnooneY = 270;
                    switch (reader.GetString(5))
                    {
                        case "1":
                            oSheet.Cells[6, 1] = reader.GetString(3);
                            FirstCNO = reader.GetString(0);
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
                reader.Close();

                if ((Client == "SGA" || Client == "Airtanks Limited") && PackingMarks.Trim().CompareTo("SGA-GLADIATAIR") == 0)
                {
                    string ProductNO = "";
                    //該客戶要其自己的logo  PartNo   Part Description
                    selectCmd = "SELECT  Product_NO FROM MSNBody,Manufacturing where [CylinderNo]='" + FirstCNO + "' and vchManufacturingNo=  Manufacturing_NO";
                    cmd = new SqlCommand(selectCmd, conn);
                    reader = cmd.ExecuteReader();
                    if (reader.Read())
                    {
                        ProductNO = reader.GetValue(0).ToString();
                    }
                    reader.Close();

                    selectCmd = "SELECT  ProductCode, ProductDescription FROM CustomerPackingMark where ProductNo='" + ProductNO + "' and LogoType='" + (PackingMarks.Trim().Contains("-") == true ? PackingMarks.Trim().Split('-')[1].Trim().ToUpper() : PackingMarks.Trim()) + "'";
                    cmd = new SqlCommand(selectCmd, conn);
                    reader = cmd.ExecuteReader();
                    if (reader.Read())
                    {
                        //載入客戶產品名稱
                        oSheet.Cells[1, 7] = reader.GetString(1);

                        //載入客戶產品型號
                        oSheet.Cells[2, 7] = reader.GetString(0);
                    }
                    reader.Close();

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

                conn.Close();

                //if (StorageStatus == "N")//20190212
                {
                    //預設位子在X:680,Y:155
                    //預設QRCODE圖片大小250*250

                    //插入圖片
                    int picX = 680, picY = 183;
                    string picadd = @"C:\QRCode\";
                    myConnection = new SqlConnection(myConnectionString);
                    selectCmd = "SELECT  * FROM [ShippingHead] where [ListDate]='" + ListDateListBox.SelectedItem + "' and [ProductName]='" + ProductComboBox.Text + "' and [vchBoxs]='" + BoxsListBox.SelectedItem + "'";
                    conn = new SqlConnection(myConnectionString);
                    conn.Open();
                    cmd = new SqlCommand(selectCmd, conn);
                    reader = cmd.ExecuteReader();
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

                oSheet.Shapes.AddPicture(oneadd + BoxsListBox.SelectedItem + ".png", Microsoft.Office.Core.MsoTriState.msoFalse,
                                    Microsoft.Office.Core.MsoTriState.msoTrue, oneX, oneY, 170, 35);
                //載入嘜頭資料
                selectCmd = "SELECT  [Client],isnull([CustomerPO],''),isnull([CustomerProductName],''),isnull([CustomerProductNo],''),[vchBoxs] FROM [ShippingHead] where [ListDate]='" + ListDateListBox.SelectedItem + "' and [ProductName]='" + ProductComboBox.SelectedItem + "'and [vchBoxs]='" + BoxsListBox.SelectedItem + "'";
                conn = new SqlConnection(myConnectionString);
                conn.Open();
                cmd = new SqlCommand(selectCmd, conn);
                reader = cmd.ExecuteReader();
                if (reader.Read())
                {
                    //載入P/O No.
                    oSheet.Cells[8, 3] = reader.GetString(1);
                    //載入一箱幾隻
                    oSheet.Cells[7, 3] = Getcount;
                    //載入箱號
                    oSheet.Cells[9, 3] = reader.GetString(4);
                }
                reader.Close();
                conn.Close();

                //////////
                int serialnooneX = 308, serialnooneY = 128;
                string serialnooneadd = @"C:\SerialNoCode\";
                string CylinderNo = "", HydrostaticTestDate = "", ProductNO = "";
                //載入嘜頭氣瓶序號位子v
                selectCmd = "SELECT  * FROM [ShippingBody] where [ListDate]='" + ListDateListBox.SelectedItem + "' and [ProductName]='" + ProductComboBox.SelectedItem + "'and [WhereBox]='" + BoxsListBox.SelectedItem + "'";
                conn = new SqlConnection(myConnectionString);
                conn.Open();
                cmd = new SqlCommand(selectCmd, conn);
                reader = cmd.ExecuteReader();
                while (reader.Read())
                {
                    switch (reader.GetString(5))
                    {
                        case "1":
                            oSheet.Cells[2, 4] = reader.GetString(3);
                            CylinderNo = reader.GetString(3);
                            MarkSerialNoBarCode(CylinderNo);
                            oSheet.Shapes.AddPicture(serialnooneadd + reader.GetString(3) + ".png", Microsoft.Office.Core.MsoTriState.msoFalse,
                            Microsoft.Office.Core.MsoTriState.msoTrue, serialnooneX, serialnooneY, 255, 44);
                            break;
                    }
                }
                reader.Close();
                //載入由序號找水壓年月
                selectCmd = "SELECT vchManufacturingNo, Product_NO, vchMarkingType, vchHydrostaticTestDate FROM MSNBody,Manufacturing where [CylinderNo]='" + CylinderNo + "' and vchManufacturingNo=  Manufacturing_NO";
                cmd = new SqlCommand(selectCmd, conn);
                reader = cmd.ExecuteReader();
                if (reader.Read())
                {
                    HydrostaticTestDate = reader.GetValue(3).ToString();
                    ProductNO = reader.GetValue(1).ToString();
                }
                reader.Close();
                if (HydrostaticTestDate.Contains("/") == true)
                {
                    oSheet.Cells[9, 7] = HydrostaticTestDate.Split('/')[1] + HydrostaticTestDate.Split('/')[0].Substring(2, 2);
                }
                else
                {
                    oSheet.Cells[9, 7] = HydrostaticTestDate;
                }
                //由序號找出產品型號再找出Part Description、Part No.
                selectCmd = "SELECT CustomerName, BottleType, BottleCatalog, ProductNo, LogoType, ProductCode, ProductDescription FROM CustomerPackingMark where ProductNo='" + ProductNO + "' and LogoType='" + (PackingMarks.Trim().Contains("-") == true ? PackingMarks.Trim().Split('-')[1].Trim().ToUpper() : PackingMarks.Trim()) + "'";
                cmd = new SqlCommand(selectCmd, conn);
                reader = cmd.ExecuteReader();
                if (reader.Read())
                {
                    //載入Part Description
                    oSheet.Cells[5, 3] = reader.GetString(6).Contains("-") == true ? reader.GetString(6).Replace("- ", "\n") : reader.GetString(6);
                    if (reader.GetString(6).Contains("-") == true)
                    {
                        oSheet.get_Range("C5").Font.Size = 22;
                    }
                    //oSheet.get_Range("C5").ShrinkToFit = true;// '設定為縮小字型以適合欄寬
                    //載入Part No.
                    oSheet.Cells[6, 3] = reader.GetString(5);
                }
                reader.Close();
                conn.Close();

                //預設位子在X:446,Y:228
                //預設QRCODE圖片大小190*190

                //插入二維條碼
                int picX = 444, picY = 228;
                string picadd = @"C:\QRCode\";
                selectCmd = "SELECT  * FROM [ShippingHead] where [ListDate]='" + ListDateListBox.SelectedItem + "' and [ProductName]='" + ProductComboBox.Text + "' and [vchBoxs]='" + BoxsListBox.SelectedItem + "'";
                conn = new SqlConnection(myConnectionString);
                conn.Open();
                cmd = new SqlCommand(selectCmd, conn);
                reader = cmd.ExecuteReader();
                if (reader.Read())
                {
                    oSheet.Shapes.AddPicture(picadd + (reader.GetString(0) + reader.GetString(1) + reader.GetString(3)) + ".png", Microsoft.Office.Core.MsoTriState.msoFalse,
                    Microsoft.Office.Core.MsoTriState.msoTrue, picX, picY, 190, 190);
                }
                reader.Close();
                conn.Close();

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
            if (PackingMarks.Trim().CompareTo("Regulator 3000psi") == 0)
            {
                //該客戶要其自己的logo
                oSheet.Shapes.AddPicture(Application.StartupPath + @".\LOGO_Regulator 3000psi.png", Microsoft.Office.Core.MsoTriState.msoFalse,
                                Microsoft.Office.Core.MsoTriState.msoTrue, 45, 2, 545, 450);
            }
            else if (PackingMarks.Trim().CompareTo("Estratego-48ci 3000psi+Regulator") == 0)
            {
                //該客戶要其自己的logo
                oSheet.Shapes.AddPicture(Application.StartupPath + @".\LOGO_Estratego_48ci 3000psi_Regulator.png", Microsoft.Office.Core.MsoTriState.msoFalse,
                                Microsoft.Office.Core.MsoTriState.msoTrue, 45, 2, 545, 450);
            }
            else if (PackingMarks.Trim().CompareTo("Estratego-48ci 3000psi") == 0)
            {
                //該客戶要其自己的logo
                oSheet.Shapes.AddPicture(Application.StartupPath + @".\LOGO_Estratego_48ci 3000psi.png", Microsoft.Office.Core.MsoTriState.msoFalse,
                                Microsoft.Office.Core.MsoTriState.msoTrue, 45, 2, 545, 450);
            }
            else if (PackingMarks.Trim().CompareTo("Estratego-13ci 3000psi+Regulator") == 0)
            {
                //該客戶要其自己的logo
                oSheet.Shapes.AddPicture(Application.StartupPath + @".\LOGO_Estratego_13ci 3000psi_Regulator.png", Microsoft.Office.Core.MsoTriState.msoFalse,
                                Microsoft.Office.Core.MsoTriState.msoTrue, 45, 2, 545, 450);
            }
            else if (PackingMarks.Trim().CompareTo("Estratego-13ci 3000psi") == 0)
            {
                //該客戶要其自己的logo
                oSheet.Shapes.AddPicture(Application.StartupPath + @".\LOGO_Estratego_13ci 3000psi.png", Microsoft.Office.Core.MsoTriState.msoFalse,
                                Microsoft.Office.Core.MsoTriState.msoTrue, 45, 2, 545, 450);
            }
            else if (PackingMarks.Trim().CompareTo("Estratego-12oz") == 0)
            {
                //該客戶要其自己的logo
                oSheet.Shapes.AddPicture(Application.StartupPath + @".\LOGO_Estratego_12oz.png", Microsoft.Office.Core.MsoTriState.msoFalse,
                                Microsoft.Office.Core.MsoTriState.msoTrue, 45, 2, 545, 450);
            }
            else if (PackingMarks.Trim().CompareTo("Estratego-20oz") == 0)
            {
                //該客戶要其自己的logo
                oSheet.Shapes.AddPicture(Application.StartupPath + @".\LOGO_Estratego_20oz.png", Microsoft.Office.Core.MsoTriState.msoFalse,
                                Microsoft.Office.Core.MsoTriState.msoTrue, 45, 2, 545, 450);
            }
            else if (PackingMarks.Trim().CompareTo("Estratego-68ci(Assault) 4500psi") == 0)
            {
                //該客戶要其自己的logo
                oSheet.Shapes.AddPicture(Application.StartupPath + @".\LOGO_Estratego_68ci(Assault) 4500psi.png", Microsoft.Office.Core.MsoTriState.msoFalse,
                                Microsoft.Office.Core.MsoTriState.msoTrue, 45, 2, 545, 450);
            }
            else if (PackingMarks.Trim().CompareTo("Estratego-68ci(Snow White) 4500psi") == 0)
            {
                //該客戶要其自己的logo
                oSheet.Shapes.AddPicture(Application.StartupPath + @".\LOGO_Estratego_68ci(Snow White) 4500psi.png", Microsoft.Office.Core.MsoTriState.msoFalse,
                                Microsoft.Office.Core.MsoTriState.msoTrue, 45, 2, 545, 450);
            }
            else if (PackingMarks.Trim().CompareTo("Estratego_68ci_UL") == 0)
            {
                //該客戶要其自己的logo
                oSheet.Shapes.AddPicture(Application.StartupPath + @".\LOGO_Estratego_68ci_UL.png", Microsoft.Office.Core.MsoTriState.msoFalse,
                                Microsoft.Office.Core.MsoTriState.msoTrue, 45, 2, 545, 450);
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
            myConnection = new SqlConnection(myConnectionString);
            selectCmd = "Update [LaserMarkDirection] SET  [vchWhere]=0";
            conn = new SqlConnection(myConnectionString);
            conn.Open();
            cmd = new SqlCommand(selectCmd, conn);
            reader = cmd.ExecuteReader();
            reader.Read();
            reader.Close();
            conn.Close();
        }

        private void button4_Click(object sender, EventArgs e)
        {
            //更新氣瓶相關資料進入MSNBody資料表
            myConnection = new SqlConnection(myConnectionString);
            selectCmd = "Update [LaserMarkDirection] SET  [vchWhere]=1";
            conn = new SqlConnection(myConnectionString);
            conn.Open();
            cmd = new SqlCommand(selectCmd, conn);
            reader = cmd.ExecuteReader();
            reader.Read();
            reader.Close();
            conn.Close();
        }

        private bool PerformanceTest(string ManufacturingNo, string CylinderNo)
        {
            //可參考20160714之註解

            string ProductNo = "", ProductType = ""; conn = new SqlConnection(myConnectionString);
            conn.Open();
            selectCmd = "SELECT  Product.Type, Product.Product_No  FROM Manufacturing,Product where  Manufacturing.Product_NO=Product.Product_No and Manufacturing.Manufacturing_NO='" + ManufacturingNo + "'";

            cmd = new SqlCommand(selectCmd, conn);
            reader = cmd.ExecuteReader();
            if (reader.Read())
            {
                ProductNo = reader.GetValue(1).ToString();
                if (reader.IsDBNull(0) == false && reader.GetString(0).CompareTo("鋁瓶") == 0)
                {
                    ProductType = reader.GetValue(0).ToString();
                    //鋁瓶
                }
                else if (reader.IsDBNull(0) == false && reader.GetString(0).CompareTo("鋁內膽") == 0)
                {
                    ProductType = reader.GetValue(0).ToString();
                    //鋁內膽
                }
                else if (reader.IsDBNull(0) == false && reader.GetString(0).CompareTo("複合瓶") == 0)
                {
                    ProductType = reader.GetValue(0).ToString();
                    //複合瓶
                }
            }
            reader.Close();
            conn.Close();

            //判別產品類型
            if (ProductType.CompareTo("鋁瓶") == 0)
            {
                if (ProductNo != "")
                {
                    if (ProductNo.Contains("1-A-") == true)
                    {
                        //以副程式顯示 回傳True or False
                        //彎曲(Bending)or壓扁(Squash)、拉伸(tensile)、硬度(Hardness)、爆破(Burst)、水壓(Hydro)(2019/04/01)
                        if (BendingFinish(ManufacturingNo) == false && SquashFinish(ManufacturingNo) == false)
                        {
                            BottleTextBox.Text = "";
                            BottomTextBox.Text = "";
                            HistoryListBox.Items.Add("無彎曲或壓扁資料。批號:" + ManufacturingNo);
                            MessageBox.Show("無彎曲或壓扁資料，請聯繫品保人員。");
                            return false;
                        }
                        if (TensileFinish(ManufacturingNo) == false)
                        {
                            BottleTextBox.Text = "";
                            BottomTextBox.Text = "";
                            HistoryListBox.Items.Add("無拉伸資料。批號:" + ManufacturingNo);
                            MessageBox.Show("無拉伸資料，請聯繫品保人員。");
                            return false;
                        }
                        if (HardnessFinish(ManufacturingNo) == false)
                        {
                            BottleTextBox.Text = "";
                            BottomTextBox.Text = "";
                            HistoryListBox.Items.Add("無硬度資料。批號:" + ManufacturingNo);
                            MessageBox.Show("無硬度資料，請聯繫品保人員。");
                            return false;
                        }
                        if (BurstFinish(ManufacturingNo) == false)
                        {
                            BottleTextBox.Text = "";
                            BottomTextBox.Text = "";
                            HistoryListBox.Items.Add("無爆破資料。批號:" + ManufacturingNo);
                            MessageBox.Show("無爆破資料，請聯繫品保人員。");
                            return false;
                        }
                        if (HydroFinish(ManufacturingNo, CylinderNo) == false) //水壓(2019/04/01)
                        {
                            BottleTextBox.Text = "";
                            BottomTextBox.Text = "";
                            HistoryListBox.Items.Add("無水壓報告資料。批號:" + ManufacturingNo);
                            MessageBox.Show("無水壓報告資料。");
                            return false;
                        }
                    }
                    else if (ProductNo.Contains("3-A-") == true)
                    {
                        //拉伸、壓扁、水壓(2019/04/01)
                        if (TensileFinish(ManufacturingNo) == false)
                        {
                            BottleTextBox.Text = "";
                            BottomTextBox.Text = "";
                            HistoryListBox.Items.Add("無拉伸資料。批號:" + ManufacturingNo);
                            MessageBox.Show("無拉伸資料，請聯繫品保人員。");
                            return false;
                        }
                        if (SquashFinish(ManufacturingNo) == false)
                        {
                            BottleTextBox.Text = "";
                            BottomTextBox.Text = "";
                            HistoryListBox.Items.Add("無拉伸資料。批號:" + ManufacturingNo);
                            MessageBox.Show("無拉伸資料，請聯繫品保人員。");
                            return false;
                        }
                        if (HydroFinish(ManufacturingNo,CylinderNo) == false) //水壓(2019/04/01)
                        {
                            BottleTextBox.Text = "";
                            BottomTextBox.Text = "";
                            HistoryListBox.Items.Add("無水壓報告資料。批號:" + ManufacturingNo);
                            MessageBox.Show("無水壓報告資料。");
                            return false;
                        }
                    }
                    else if (ProductNo.Contains("5-A-") == true)
                    {
                        //爆破 水壓(2019/04/01)
                        if (BurstFinish(ManufacturingNo) == false)
                        {
                            BottleTextBox.Text = "";
                            BottomTextBox.Text = "";
                            HistoryListBox.Items.Add("無爆破資料。批號:" + ManufacturingNo);
                            MessageBox.Show("無爆破資料，請聯繫品保人員。");
                            return false;
                        }
                        if (HydroFinish(ManufacturingNo,CylinderNo) == false) //水壓(2019/04/01)
                        {
                            BottleTextBox.Text = "";
                            BottomTextBox.Text = "";
                            HistoryListBox.Items.Add("無水壓報告資料。批號:" + ManufacturingNo);
                            MessageBox.Show("無水壓報告資料。");
                            return false;
                        }
                    }
                }
            }
            else if (ProductType.CompareTo("鋁內膽") == 0)
            {
                //拉伸、爆破、水壓(2019/04/01)
                if (TensileFinish(ManufacturingNo) == false)
                {
                    BottleTextBox.Text = "";
                    BottomTextBox.Text = "";
                    HistoryListBox.Items.Add("無拉伸資料。批號:" + ManufacturingNo);
                    MessageBox.Show("無拉伸資料，請聯繫品保人員。");
                    return false;
                }
                if (BurstFinish(ManufacturingNo) == false)
                {
                    BottleTextBox.Text = "";
                    BottomTextBox.Text = "";
                    HistoryListBox.Items.Add("無爆破資料。批號:" + ManufacturingNo);
                    MessageBox.Show("無爆破資料，請聯繫品保人員。");
                    return false;
                }
                if (HydroFinish(ManufacturingNo,CylinderNo) == false) //水壓(2019/04/01)
                {
                    BottleTextBox.Text = "";
                    BottomTextBox.Text = "";
                    HistoryListBox.Items.Add("無水壓報告資料。批號:" + ManufacturingNo);
                    MessageBox.Show("無水壓報告資料。");
                    return false;
                }
            }
            else if (ProductType.CompareTo("複合瓶") == 0)
            {
                //20180905 判別是否有做出貨檢驗，無出貨檢驗資料不允許包裝
                connP = new SqlConnection(ESIGNmyConnectionString);
                connP.Open();

                selectCmdP = "SELECT  * FROM  CH_ShippingInspection where LotNo='" + ManufacturingNo + "' and SerialNo='" + CylinderNo + "'";
                cmdP = new SqlCommand(selectCmdP, connP);
                readerP = cmdP.ExecuteReader();

                if (readerP.Read())
                {
                    readerP.Close();
                    connP.Close();
                }
                else
                {
                    readerP.Close();
                    connP.Close();
                    HistoryListBox.Items.Add("無出貨檢驗資料。序號:" + CylinderNo);
                    MessageBox.Show("無出貨檢驗資料，請聯繫品保人員。");
                    return false;
                }

                //對應內膽  拉伸、爆破
                //找出對應內膽批號
                string BuildUp = "";
                selectCmd = "SELECT BuildUp FROM Manufacturing where Manufacturing_NO='" + ManufacturingNo + "'";
                conn = new SqlConnection(myConnectionString);
                conn.Open();
                cmd = new SqlCommand(selectCmd, conn);
                reader = cmd.ExecuteReader();
                if (reader.Read())
                {
                    BuildUp = reader.GetValue(0).ToString();
                }
                reader.Close();
                conn.Close();

                if (BuildUp != "")
                {
                    for (int SubLiner = 0; SubLiner < BuildUp.Split(',').Length; SubLiner++)
                    {
                        //拉伸、爆破、水壓(2019/04/01)
                        if (TensileFinish(BuildUp.Split(',')[SubLiner]) == false)
                        {
                            BottleTextBox.Text = "";
                            BottomTextBox.Text = "";
                            HistoryListBox.Items.Add("無對應內膽拉伸資料。批號:" + BuildUp.Split(',')[SubLiner]);
                            MessageBox.Show("無對應內膽拉伸資料，請聯繫品保人員。");
                            return false;
                        }
                        if (BurstFinish(BuildUp.Split(',')[SubLiner]) == false)
                        {
                            BottleTextBox.Text = "";
                            BottomTextBox.Text = "";
                            HistoryListBox.Items.Add("無對應內膽爆破資料。批號:" + BuildUp.Split(',')[SubLiner]);
                            MessageBox.Show("無對應內膽爆破資料，請聯繫品保人員。");
                            return false;
                        }
                        if (HydroFinish(ManufacturingNo,CylinderNo) == false) //水壓(2019/04/01)
                        {
                            BottleTextBox.Text = "";
                            BottomTextBox.Text = "";
                            HistoryListBox.Items.Add("無水壓報告資料。批號:" + ManufacturingNo);
                            MessageBox.Show("無水壓報告資料。");
                            return false;
                        }

                    }
                }

                //爆破、循環(Cycle)、進料(Feed)(碳纖、玻纖、樹脂)、水壓(2019/04/01)
                if (BurstFinish(ManufacturingNo) == false)
                {
                    BottleTextBox.Text = "";
                    BottomTextBox.Text = "";
                    HistoryListBox.Items.Add("無爆破資料。批號:" + ManufacturingNo);
                    MessageBox.Show("無爆破資料，請聯繫品保人員。");
                    return false;
                }
                if (CycleFinish(ManufacturingNo) == false)
                {
                    BottleTextBox.Text = "";
                    BottomTextBox.Text = "";
                    HistoryListBox.Items.Add("無循環資料。批號:" + ManufacturingNo);
                    MessageBox.Show("無循環資料，請聯繫品保人員。");
                    return false;
                }
                if (FeedFinish(ManufacturingNo) == false)
                {
                    HistoryListBox.Items.Add("無碳纖、玻纖、樹脂檢驗資料。批號:" + ManufacturingNo);
                    MessageBox.Show("無碳纖、玻纖、樹脂檢驗資料，請聯繫品保人員。");
                    return false;
                }
                if (HydroFinish(ManufacturingNo,CylinderNo) == false) //水壓(2019/04/01)
                {
                    BottleTextBox.Text = "";
                    BottomTextBox.Text = "";
                    HistoryListBox.Items.Add("無水壓報告資料。批號:" + ManufacturingNo);
                    MessageBox.Show("無水壓報告資料。");
                    return false;
                }
            }
            return true;//表示有做功性能測試
        }

        private bool HydroFinish(string MNO,string CylinderNo)
        {//水壓
            bool IsFinish = false;
            bool IsReportFinish = false;
            using (connP = new SqlConnection(ESIGNmyConnectionString))
            {
                connP.Open();

                selectCmdP = "SELECT  * FROM [AMS2].[dbo].[PPT_Hydro] WHERE [LotNo] = '" + MNO + "' and FinalResult = 'PASS' And ([AcceptanceNo] Like 'Q019%' Or [AcceptanceNo] Like '%Q030%')";
                cmdP = new SqlCommand(selectCmdP, connP);
                using (readerP = cmdP.ExecuteReader())
                {
                    if (readerP.Read())
                    {
                        IsFinish = true;
                    }
                    else
                    {
                        IsFinish = false;
                    }
                }

                selectCmdP = "SELECT [SerialNo] FROM [PPT_Hydro_Details] WHERE [SerialNo] = '" + CylinderNo + "' And ([AcceptanceNo] Like 'Q019%' Or [AcceptanceNo] Like '%Q030%')";
                cmdP = new SqlCommand(selectCmdP, connP);
                using (readerP = cmdP.ExecuteReader())
                {
                    if (readerP.Read())
                    {
                        IsReportFinish = true;
                    }
                    else
                    {
                        IsReportFinish = false;
                    }
                }
            }
            return IsFinish && IsReportFinish;
        }

        private bool TensileFinish(string MNO)
        {//拉伸
            bool IsFinish = false;
            connP = new SqlConnection(ESIGNmyConnectionString);
            connP.Open();

            selectCmdP = "SELECT  * FROM [PPT_Tensile] WHERE [ManufacturingNo] = '" + MNO + "' and FinalResult='PASS' order by ID desc";
            cmdP = new SqlCommand(selectCmdP, connP);
            readerP = cmdP.ExecuteReader();

            if (readerP.HasRows)
            {
                IsFinish = true;
            }
            else
            {
                IsFinish = false;
            }
            readerP.Close();
            connP.Close();

            return IsFinish;
        }

        private bool BendingFinish(string MNO)
        {//彎曲
            bool IsFinish = false;
            connP = new SqlConnection(ESIGNmyConnectionString);
            connP.Open();
            selectCmdP = "SELECT  * FROM [PPT_FlatBend] WHERE [ManufacturingNo] = '" + MNO + "' and FinalResult='PASS' and Method='彎曲' order by id desc";
            cmdP = new SqlCommand(selectCmdP, connP);
            readerP = cmdP.ExecuteReader();

            if (readerP.HasRows)
            {
                IsFinish = true;
            }
            else
            {
                IsFinish = false;
            }
            readerP.Close();
            connP.Close();

            return IsFinish;
        }

        private bool SquashFinish(string MNO)
        {//壓扁
            bool IsFinish = false;
            connP = new SqlConnection(ESIGNmyConnectionString);
            connP.Open();
            selectCmdP = "SELECT  * FROM [PPT_FlatBend] WHERE [ManufacturingNo] = '" + MNO + "' and FinalResult='PASS' and Method='壓扁' order by id desc";
            cmdP = new SqlCommand(selectCmdP, connP);
            readerP = cmdP.ExecuteReader();

            if (readerP.HasRows)
            {
                IsFinish = true;
            }
            else
            {
                IsFinish = false;
            }
            readerP.Close();
            connP.Close();

            return IsFinish;
        }

        private bool BurstFinish(string MNO)
        {//爆破
            bool IsFinish = false;
            connP = new SqlConnection(ESIGNmyConnectionString);
            connP.Open();

            selectCmdP = "SELECT  * FROM [PPT_Burst] WHERE [ManufacturingNo] = '" + MNO + "' and FinalResult='PASS' order by AcceptanceNo desc";
            cmdP = new SqlCommand(selectCmdP, connP);
            readerP = cmdP.ExecuteReader();

            if (readerP.HasRows)
            {
                IsFinish = true;
            }
            else
            {
                IsFinish = false;
            }
            readerP.Close();
            connP.Close();

            return IsFinish;
        }

        private bool HardnessFinish(string MNO)
        {//硬度
            bool IsFinish = false;
            connP = new SqlConnection(ESIGNmyConnectionString);
            connP.Open();

            selectCmdP = "SELECT  * FROM QCDocument INNER JOIN Esign2 ON QCDocument.AcceptanceNo = Esign2.AcceptanceNo WHERE (QCDocument.LotNo = '" + MNO + "') AND (Esign2.Type LIKE '硬度%')";
            cmdP = new SqlCommand(selectCmdP, connP);
            readerP = cmdP.ExecuteReader();

            if (readerP.HasRows)
            {
                IsFinish = true;
            }
            else
            {
                IsFinish = false;
            }
            readerP.Close();
            connP.Close();

            return IsFinish;
        }

        private bool CycleFinish(string MNO)
        {//循環
            bool IsFinish = false;
            connP = new SqlConnection(ESIGNmyConnectionString);
            connP.Open();
            selectCmdP = "SELECT  * FROM [PPT_Cycling] WHERE [LotNo] = '" + MNO + "' and FinalResult='PASS'";
            cmdP = new SqlCommand(selectCmdP, connP);
            readerP = cmdP.ExecuteReader();

            if (readerP.HasRows)
            {
                IsFinish = true;
            }
            else
            {
                IsFinish = false;
            }
            readerP.Close();
            connP.Close();

            return IsFinish;
        }

        private bool FeedFinish(string MNO)
        {//進料  碳纖、玻纖、樹脂
            bool IsFinish = false;
            string ResinLotNo = "", CarbonLotNo = "", GlassLotNo = "";
            string CarbonSpec = "", GlassSpec = "";
            connP = new SqlConnection(ESIGNmyConnectionString);
            connP.Open();
            selectCmdP = "SELECT  ResinLotNo, CarbonLotNo, GlassLotNo,CarbonSpec,GlassSpec FROM [FilamentWinding] WHERE [LotNo] = '" + MNO + "' order by id desc";
            cmdP = new SqlCommand(selectCmdP, connP);
            readerP = cmdP.ExecuteReader();
            if (readerP.Read())
            {
                ResinLotNo = readerP.GetValue(0).ToString();
                CarbonLotNo = readerP.GetValue(1).ToString();
                GlassLotNo = readerP.GetValue(2).ToString();
                CarbonSpec = readerP.GetValue(3).ToString();
                GlassSpec = readerP.GetValue(4).ToString();
            }
            readerP.Close();
            if (ResinLotNo == "" || CarbonLotNo == "" || GlassLotNo == "")
            {
                connP.Close();
                return true;
            }

            selectCmdP = "SELECT * FROM [IQC] A, [Esign2] B WHERE A.[AcceptanceNo]=B.[AcceptanceNo] AND A.[Type] = '碳纖' AND A.[LotNo] LIKE '%" + CarbonLotNo + "%'";
            cmdP = new SqlCommand(selectCmdP, connP);
            readerP = cmdP.ExecuteReader();
            if (readerP.Read())
            {
                IsFinish = true;
            }
            else
            {
                readerP.Close();
                connP.Close();
                IsFinish = false;
                return IsFinish;
            }
            readerP.Close();

            selectCmdP = "SELECT * FROM [PPT] A, [Esign2] B WHERE A.[AcceptanceNo]=B.[AcceptanceNo] AND A.[Type] = '玻纖' AND A.[LotNo] LIKE '%" + GlassLotNo + "%'";
            cmdP = new SqlCommand(selectCmdP, connP);
            readerP = cmdP.ExecuteReader();
            if (readerP.Read())
            {
                IsFinish = true;
            }
            else
            {
                readerP.Close();
                connP.Close();
                IsFinish = false;
                return IsFinish;
            }
            readerP.Close();

            //selectCmdP = "SELECT * FROM [PPT] A, [Esign2] B WHERE A.[AcceptanceNo]=B.[AcceptanceNo] AND A.[Type] = '樹脂' AND A.[LotNo] LIKE '%" + ResinLotNo + "%' and FiberType ='玻' and FiberLotNo like '%" + GlassLotNo + "%'";
            selectCmdP = "SELECT * FROM [PPT] A, [Esign2] B WHERE A.[AcceptanceNo]=B.[AcceptanceNo] AND A.[Type] = '樹脂' AND A.[LotNo] LIKE '%" + ResinLotNo + "%' and FiberType ='玻' and (FiberLotNo like '%" + GlassLotNo + "%' or FiberSpec like '%" + GlassSpec + "%')";//20180912品保系統檢驗組組長 說只要規格一樣沒有對應批號也可以。當初為CE0086有問題
            cmdP = new SqlCommand(selectCmdP, connP);
            readerP = cmdP.ExecuteReader();
            if (readerP.Read())
            {
                IsFinish = true;
            }
            else
            {
                readerP.Close();
                connP.Close();
                IsFinish = false;
                return IsFinish;
            }
            readerP.Close();

            //selectCmdP = "SELECT * FROM [PPT] A, [Esign2] B WHERE A.[AcceptanceNo]=B.[AcceptanceNo] AND A.[Type] = '樹脂' AND A.[LotNo] LIKE '%" + ResinLotNo + "%' and FiberType ='碳' and FiberLotNo like '%" + CarbonLotNo + "%'";
            selectCmdP = "SELECT * FROM [PPT] A, [Esign2] B WHERE A.[AcceptanceNo]=B.[AcceptanceNo] AND A.[Type] = '樹脂' AND A.[LotNo] LIKE '%" + ResinLotNo + "%' and FiberType ='碳' and (FiberLotNo like '%" + CarbonLotNo + "%' or FiberSpec like '%" + CarbonSpec + "%')";//20180912品保系統檢驗組組長 說只要規格一樣沒有對應批號也可以。當初為CE0086有問題
            cmdP = new SqlCommand(selectCmdP, connP);
            readerP = cmdP.ExecuteReader();
            if (readerP.Read())
            {
                IsFinish = true;
            }
            else
            {
                readerP.Close();
                connP.Close();
                IsFinish = false;
                return IsFinish;
            }
            readerP.Close();
            connP.Close();

            return IsFinish;
        }

        private void timer1_Tick(object sender, EventArgs e)
        {//AutoAccumulate

            //用來記錄氣瓶序號
            string CylinderNumbers = "";

            if (((BottleTextBox.Text == BottomTextBox.Text) && (BottleTextBox.Text != "" || BottomTextBox.Text != "")) || (Pass == "Y" && (BottleTextBox.Text != "" || BottomTextBox.Text != "")))
            {
                // string FredlovCSV = "N";
                //string CalisoCSV = "N";

                string HydrostaticPass = "N";

                if (BottleTextBox.Text != "")
                {
                    CylinderNumbers = BottleTextBox.Text;
                }
                else if (BottomTextBox.Text != "")
                {
                    CylinderNumbers = BottomTextBox.Text;
                }
                //判別是否為報廢氣瓶
                selectCmd = "SELECT  * FROM [ComplexScrapData] where [ComplexCylinderNo]='" + CylinderNumbers + "'";
                conn = new SqlConnection(myConnectionString);
                conn.Open();
                cmd = new SqlCommand(selectCmd, conn);
                reader = cmd.ExecuteReader();
                if (reader.Read())
                {
                    BottleTextBox.Text = "";
                    BottomTextBox.Text = "";
                    MessageBox.Show("此序號之氣瓶為報廢氣瓶，故不允許加入", "警告-W006");
                    HistoryListBox.Items.Add(NowTime());
                    HistoryListBox.Items.Add("此序號為報廢氣瓶：" + CylinderNumbers);
                    BottleTextBox.Focus();
                    reader.Close();
                    conn.Close();
                    return;
                }
                reader.Close();
                conn.Close();

                //20181002 報工報廢
                selectCmd = "SELECT  * FROM [RePortScrapReason] where [ScrapCylinderNO]='" + CylinderNumbers + "'";
                conn = new SqlConnection(myConnectionString);
                conn.Open();
                cmd = new SqlCommand(selectCmd, conn);
                reader = cmd.ExecuteReader();
                if (reader.Read())
                {
                    BottleTextBox.Text = "";
                    BottomTextBox.Text = "";
                    MessageBox.Show("此序號之氣瓶為報廢氣瓶，故不允許加入", "警告-W006");
                    HistoryListBox.Items.Add(NowTime());
                    HistoryListBox.Items.Add("此序號為報廢氣瓶：" + CylinderNumbers);
                    BottleTextBox.Focus();
                    reader.Close();
                    conn.Close();
                    return;
                }
                reader.Close();
                conn.Close();

                //判斷是否已經有相同的序號入嘜頭
                myConnection = new SqlConnection(myConnectionString);
                selectCmd = "SELECT  * FROM [ShippingBody] where [CylinderNumbers]='" + CylinderNumbers + "'";
                conn = new SqlConnection(myConnectionString);
                conn.Open();
                cmd = new SqlCommand(selectCmd, conn);
                reader = cmd.ExecuteReader();
                if (reader.Read())
                {
                    BottleTextBox.Text = "";
                    BottomTextBox.Text = "";
                    MessageBox.Show("此序號已存入嘜頭資訊！(在第" + reader.GetString(4) + "箱，第" + reader.GetString(5) + "位置)", "警告-W001");
                    HistoryListBox.Items.Add(NowTime());
                    HistoryListBox.Items.Add("此序號已重複：" + CylinderNumbers);
                    BottleTextBox.Focus();
                    reader.Close();
                    conn.Close();
                    return;
                }
                reader.Close();
                conn.Close();

                string ManufacturingNo = "";
                string SpecialUses = "N";

                selectCmd = "SELECT Manufacturing_NO, isnull([H_SpecialUses],'N') FROM [MSNBody],[Manufacturing] where [CylinderNo]='" + CylinderNumbers + "' and Manufacturing_NO=vchManufacturingNo";
                conn = new SqlConnection(myConnectionString);
                conn.Open();
                cmd = new SqlCommand(selectCmd, conn);
                reader = cmd.ExecuteReader();
                if (reader.Read())
                {
                    ManufacturingNo = reader.GetString(0);
                    if (reader.GetValue(1).ToString() == "Y")
                    {
                        SpecialUses = "Y";
                    }
                }
                reader.Close();
                conn.Close();

                ////取得製造批號
                //myConnection = new SqlConnection(myConnectionString);
                //selectCmd = "SELECT  * FROM [MSNBody] where [vchCylinderCode]+[vchCylinderNo]='" + CylinderNumbers + "'";
                //conn = new SqlConnection(myConnectionString);
                //conn.Open();
                //cmd = new SqlCommand(selectCmd, conn);
                //reader = cmd.ExecuteReader();
                //if (reader.Read())
                //{
                //    ManufacturingNo = reader.GetString(0);
                //}
                //reader.Close();
                //conn.Close();


                //if (ManufacturingNo != "")
                //{
                //    //判斷此批號是否是走特採的批號
                //    myConnection = new SqlConnection(myConnectionString);
                //    selectCmd = "SELECT  * FROM [Manufacturing] where [Manufacturing_NO]='" + ManufacturingNo + "' and [H_SpecialUses]='Y'";
                //    conn = new SqlConnection(myConnectionString);
                //    conn.Open();
                //    cmd = new SqlCommand(selectCmd, conn);
                //    reader = cmd.ExecuteReader();
                //    if (reader.Read())
                //    {
                //        SpecialUses = "Y";
                //    }
                //    reader.Close();
                //    conn.Close();
                //}

                if (SpecialUses == "N")
                {
                    conn = new SqlConnection(myConnectionString);
                    conn.Open();

                    selectCmd = "SELECT  * FROM [HydrostaticPass] where [ManufacturingNo]='" + ManufacturingNo + "' and [CylinderNo]='" + CylinderNumbers + "' and [HydrostaticPass]='Y'";
                    cmd = new SqlCommand(selectCmd, conn);
                    reader = cmd.ExecuteReader();
                    if (reader.Read())
                    {
                        HydrostaticPass = "Y";
                    }
                    reader.Close();

                    if (HydrostaticPass == "N")
                    {
                        //找對應的舊序號，若有序號則依此序號查是否有做過水壓
                        string OriCNo = "", OriMNO = "";
                        selectCmd = "SELECT  OriCylinderNo,OriManufacturingNo, NewCylinderNo FROM [ChangeCylinderNo] where [NewManufacturingNo]='" + ManufacturingNo + "' and [NewCylinderNo]='" + CylinderNumbers + "' ";
                        cmd = new SqlCommand(selectCmd, conn);
                        reader = cmd.ExecuteReader();
                        if (reader.Read())
                        {
                            OriCNo = reader.GetString(0);
                            OriMNO = reader.GetString(1);
                        }
                        reader.Close();

                        if (OriCNo != "")
                        {
                            selectCmd = "SELECT  * FROM [HydrostaticPass] where [ManufacturingNo]='" + OriMNO + "' and [CylinderNo]='" + OriCNo + "' and [HydrostaticPass]='Y'";
                            cmd = new SqlCommand(selectCmd, conn);
                            reader = cmd.ExecuteReader();
                            if (reader.Read())
                            {
                                HydrostaticPass = "Y";
                            }
                            reader.Close();
                        }
                    }

                    conn.Close();

                    if (HydrostaticPass == "N")
                    {
                        BottleTextBox.Text = "";
                        BottomTextBox.Text = "";
                        MessageBox.Show("此序號查詢不到水壓測試資料！", "警告-W002");
                        HistoryListBox.Items.Add(NowTime());
                        HistoryListBox.Items.Add("此序號查無水壓資訊：" + CylinderNumbers);
                        BottleTextBox.Focus();
                        return;
                    }
                    //20170515判別是否有做過成品檢驗，有才允許繼續，否則不允許包裝
                    //研發瓶轉正式出貨產品時，有可能之前的研發瓶試認證瓶所以沒有成品檢驗，因此要有成品檢驗的記錄
                    if (CheckProductAcceptanceIsWork(ManufacturingNo) == false)
                    {
                        BottleTextBox.Text = "";
                        BottomTextBox.Text = "";
                        MessageBox.Show("此批號查詢不到成品檢驗資料！", "警告-W007");
                        HistoryListBox.Items.Add(NowTime());
                        HistoryListBox.Items.Add("此批號查無成品檢驗資訊：" + ManufacturingNo);
                        BottleTextBox.Focus();
                    }
                }
                //功性能測試檢查
                if (PerformanceTest(ManufacturingNo, CylinderNumbers) == false)
                {
                    //BottleTextBox.Text = "";
                    //BottomTextBox.Text = "";
                    BottleTextBox.Focus();
                    return;
                }

                //20160714機制未完成，故先不使用
                ////複合瓶判別
                //conn = new SqlConnection(myConnectionString);
                //conn.Open();
                //selectCmd = "SELECT  Product.Type  FROM Manufacturing,Product where  Manufacturing.Product_NO=Product.Product_No and Manufacturing.Manufacturing_NO='" + ManufacturingNo + "'";

                //cmd = new SqlCommand(selectCmd, conn);
                //reader = cmd.ExecuteReader();
                //if (reader.Read())
                //{
                //    if (reader.IsDBNull(0) == false && reader.GetString(0).CompareTo("複合瓶")==0)
                //    {
                //        //複合瓶

                //        //判別是否有爆破、循環資料
                //        connP = new SqlConnection(ESIGNmyConnectionString);
                //        connP.Open();
                //        //爆破
                //        selectCmdP = "SELECT  * FROM [PPT_Burst] WHERE [ManufacturingNo] = '" + ManufacturingNo + "'";
                //        cmdP = new SqlCommand(selectCmdP, connP);
                //        readerP = cmdP.ExecuteReader();

                //        if (readerP.HasRows)
                //        {
                //            ;
                //        }
                //        else
                //        {
                //            reader.Close();
                //            conn.Close();

                //            readerP.Close();
                //            connP.Close();
                //            MessageBox.Show("查詢不到複合瓶" + ManufacturingNo + " 爆破資料！", "警告-W004");
                //            HistoryListBox.Items.Add(NowTime());
                //            HistoryListBox.Items.Add("查詢不到複合瓶" + ManufacturingNo + " 爆破資料！");
                //            BottleTextBox.Focus();
                //            return;
                //        }
                //        readerP.Close();

                //        //循環
                //        selectCmdP = "SELECT  * FROM [PPT_Cycling] WHERE [LotNo] = '" + ManufacturingNo + "'";
                //        cmdP = new SqlCommand(selectCmdP, connP);
                //        readerP = cmdP.ExecuteReader();

                //        if (readerP.HasRows)
                //        {
                //            ;
                //        }
                //        else
                //        {
                //            reader.Close();
                //            conn.Close();

                //            readerP.Close();
                //            connP.Close();
                //            MessageBox.Show("查詢不到複合瓶" + ManufacturingNo + " 循環資料！", "警告-W005");
                //            HistoryListBox.Items.Add(NowTime());
                //            HistoryListBox.Items.Add("查詢不到複合瓶" + ManufacturingNo + " 循環資料！");
                //            BottleTextBox.Focus();
                //            return;
                //        }
                //        readerP.Close();
                //        connP.Close();
                //    }
                //}
                //reader.Close();
                //conn.Close();

                //判斷新增到那個位子

                string NowSeat = "";

                //判斷[ShippingBody]是否有資料
                myConnection = new SqlConnection(myConnectionString);
                selectCmd = "SELECT  * FROM [ShippingBody] where [ListDate]='" + ListDateListBox.SelectedItem + "' and [ProductName]='" + ProductComboBox.SelectedItem + "' and [WhereBox]='" + BoxsListBox.SelectedItem + "' order by Convert(INT,[WhereSeat]) DESC ";
                conn = new SqlConnection(myConnectionString);
                conn.Open();
                cmd = new SqlCommand(selectCmd, conn);
                reader = cmd.ExecuteReader();
                if (reader.Read())
                {
                    NowSeat = reader.GetString(5);

                    if (NowSeat == Aboxof())
                    {
                        BottleTextBox.Text = "";
                        BottomTextBox.Text = "";
                        MessageBox.Show("此嘜頭已滿箱！", "警告-W003");
                        BottleTextBox.Focus();
                        return;
                    }
                }
                else
                {
                    NowSeat = "0";
                }
                reader.Close();
                conn.Close();

                //取得現在時間
                DateTime currentTime = DateTime.Now;
                //轉成字串   
                String timeString = currentTime.ToLocalTime().ToString();

                string PassselectCmd = "";

                //取得氣瓶批號
                string LotNumber=null;

                myConnection = new SqlConnection(myConnectionString);
                selectCmd = "SELECT [vchManufacturingNo] FROM [MSNBody] where [CylinderNo]='" + CylinderNumbers +"'";
                conn = new SqlConnection(myConnectionString);
                conn.Open();
                cmd = new SqlCommand(selectCmd, conn);
                reader = cmd.ExecuteReader();
                while (reader.Read())
                {
                    LotNumber = reader.GetString(0);
                }
                reader.Close();
                conn.Close();

                //如果Pass=Y SQL系統記錄此事件
                if (Pass == "Y")
                {
                    PassselectCmd = "INSERT INTO [ShippingBody] ([ListDate],[ProductName],[CylinderNumbers],[WhereBox],[WhereSeat],[vchUser],[Time],[Incomplete],[LotNumber])VALUES(" + "'" + ListDateListBox.SelectedItem + "'" + "," + "'" + ProductComboBox.SelectedItem + "'" + "," + "'" + CylinderNumbers + "'" + "," + "'" + BoxsListBox.SelectedItem + "'" + "," + "'" + (Convert.ToInt32(NowSeat) + 1) + "'," + "'" + UserListComboBox.Text + "'," + "'" + timeString + "'," + "'Y'" + ",'" + LotNumber + "')";
                }
                else
                {
                    PassselectCmd = "INSERT INTO [ShippingBody] ([ListDate],[ProductName],[CylinderNumbers],[WhereBox],[WhereSeat],[vchUser],[Time],[LotNumber])VALUES(" + "'" + ListDateListBox.SelectedItem + "'" + "," + "'" + ProductComboBox.SelectedItem + "'" + "," + "'" + CylinderNumbers + "'" + "," + "'" + BoxsListBox.SelectedItem + "'" + "," + "'" + (Convert.ToInt32(NowSeat) + 1) + "'," + "'" + UserListComboBox.Text + "'," + "'" + timeString +"','"+ LotNumber + "')";
                }

                //雷刻掃描完確認瓶身瓶底相同後載入資料
                myConnection = new SqlConnection(myConnectionString);
                selectCmd = PassselectCmd;
                conn = new SqlConnection(myConnectionString);
                conn.Open();
                cmd = new SqlCommand(selectCmd, conn);
                reader = cmd.ExecuteReader();
                reader.Close();
                conn.Close();

                //更新登出時間
                selectCmd = "UPDATE [LoginPackage] SET  [LogoutTime] = '" + DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss") + "' , [IsUpdate]='0' WHERE [ID] = '" + toolStripStatusLabel1.Text + "'";
                conn = new SqlConnection(myConnectionString);
                conn.Open();
                cmd = new SqlCommand(selectCmd, conn);
                cmd.ExecuteNonQuery();
                conn.Close();

                selectCmd = "INSERT INTO [WorkTimePackage] ([CylinderNo],[Operator],[OperatorId],[AddTime],[Date],[WorkType])VALUES(" + "'" + CylinderNumbers + "','" + User + "','" + ID + "','" + DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss") + "','" + DateTime.Now.ToString("yyyy-MM-dd") + "','" + worktype + "')";
                conn = new SqlConnection(myConnectionString);
                conn.Open();
                cmd = new SqlCommand(selectCmd, conn);
                cmd.ExecuteNonQuery();
                conn.Close();

                time = 420;

                if (LinkLMCheckBox.Checked == true)
                {
                    ;
                }
                string BoxsListBoxIndex = "";
                string NowSeat2 = "";

                //用來自動跳下一箱                
                myConnection = new SqlConnection(myConnectionString);
                selectCmd = "SELECT  * FROM [ShippingBody] where [ListDate]='" + ListDateListBox.SelectedItem + "' and [ProductName]='" + ProductComboBox.SelectedItem + "' and [WhereBox]='" + BoxsListBox.SelectedItem + "' order by Convert(INT,[WhereSeat]) DESC ";
                conn = new SqlConnection(myConnectionString);
                conn.Open();
                cmd = new SqlCommand(selectCmd, conn);
                reader = cmd.ExecuteReader();
                if (reader.Read())
                {
                    NowSeat2 = reader.GetString(5);
                    BoxsListBoxIndex = BoxsListBox.SelectedIndex.ToString();

                    reader.Close();
                    conn.Close();
                    //如果箱號已經超過最大箱數則不自動跳箱
                    if ((Convert.ToInt32(BoxsListBoxIndex) >= (BoxsListBox.Items.Count - 1)) && BoxsListBox.Items.Count != 1 && NowSeat2 == Aboxof())
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
                            BoxsListBox.SelectedIndex = (Convert.ToInt32(BoxsListBoxIndex) + 1);
                        }
                    }
                }
                else
                {
                    reader.Close();
                    conn.Close();
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
                {//如果按下的不是回退键，则取消本次(按键)动作
                    e.Handled = true;
                }
            }

            if (e.KeyChar == (char)Keys.Back)
            {
                e.Handled = true;
            }
        }

        private string NowTime()
        {

            //取得現在時間
            DateTime currentTime = DateTime.Now;
            //轉成字串   
            string timeString = currentTime.ToLocalTime().ToString();

            return timeString;
        }

        private void CheckFull()
        {
            //確定滿箱才可以列印
            myConnection = new SqlConnection(myConnectionString);
            selectCmd = "SELECT  count([WhereBox]) FROM [ShippingBody] where [ListDate]='" + ListDateListBox.SelectedItem + "' and [ProductName]='" + ProductComboBox.SelectedItem + "' and [WhereBox]='" + BoxsListBox.SelectedItem + "' ";
            conn = new SqlConnection(myConnectionString);
            conn.Open();
            cmd = new SqlCommand(selectCmd, conn);
            reader = cmd.ExecuteReader();
            if (reader.Read())
            {
                if (reader.GetInt32(0) == Convert.ToInt32(Aboxof()))
                {
                    PrintButton.Enabled = true;
                }
                else
                {
                    PrintButton.Enabled = false;
                }
            }
            reader.Close();
            conn.Close();
        }

        private string GetManufacturingCode(string Code)
        {
            char[] b = new char[12];
            StringReader sr = new StringReader(Code);
            sr.Read(b, 0, 12);
            sr.Close();
            string bb = "";
            for (int i = 0; i <= Code.Length; i++)
            {
                if (ASC(b[i]) >= 65 && ASC(b[i]) <= 90)
                {
                    bb += b[i];
                }
            }
            return bb;
        }

        private string GetManufacturingNumber(string Code)
        {
            char[] b = new char[12];
            StringReader sr = new StringReader(Code);
            sr.Read(b, 0, 12);
            sr.Close();
            string bb = "";
            for (int i = 0; i <= Code.Length; i++)
            {
                if (ASC(b[i]) >= 48 && ASC(b[i]) <= 57)
                {
                    bb += b[i];
                }
            }
            return bb;
        }

        public static int ASC(char C)
        {
            int N = Convert.ToInt32(C);

            return N;
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
                    //20160714機制未完成，故先不使用
                    ////複合瓶判別
                    //if (ComplexQRCodeCheckBox.CheckState ==CheckState.Checked)
                    //{
                    //    string CylinderNOFind = NoLMCylinderNOTextBox.Text.ToString();
                    //    if (CylinderNOFind.Length > 8)
                    //    {
                    //        CylinderNOFind = CylinderNOFind.Split((Char)13)[0];//換行符號
                    //        if (CylinderNOFind.Contains("AMS") == true)
                    //        {
                    //            CylinderNOFind = CylinderNOFind.Split(new string[] { "AMS " }, StringSplitOptions.RemoveEmptyEntries)[1];
                    //            CylinderNOFind = CylinderNOFind.Split(' ')[0];
                    //        }
                    //        else if (CylinderNOFind.Contains("TW") == true)
                    //        {
                    //            CylinderNOFind = CylinderNOFind.Split(new string[] { "TW " }, StringSplitOptions.RemoveEmptyEntries)[1];
                    //            CylinderNOFind = CylinderNOFind.Split(' ')[0];
                    //        }
                    //    }
                    //    NoLMCylinderNOTextBox.Text = CylinderNOFind;
                    //}

                    ////20141029 修改成不跳出視窗，直接在該畫面作操作。因有不連號(跳號)，原方式耗時
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
                ////20141029 修改成不跳出視窗，直接在該畫面作操作。因有不連號(跳號)，原方式耗時
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
            selectCmd = "SELECT  Client, City FROM ShippingHead where  ListDate='" + ListDateListBox.SelectedItem.ToString() + "' and  ProductName='" + ProductComboBox.Text.Trim().ToString() + "'and vchBoxs='" + BoxsListBox.SelectedItem.ToString() + "' and City is not null";
            conn = new SqlConnection(myConnectionString);
            conn.Open();
            cmd = new SqlCommand(selectCmd, conn);
            reader = cmd.ExecuteReader();
            if (reader.Read())
            {
                Client = reader.GetString(0);
                City = reader.GetString(1);
                reader.Close();
                conn.Close();
            }
            else
            {
                reader.Close();
                conn.Close();
            }
            //找出對應的
            selectCmd = "SELECT   Client, City, SCNO, ECNO FROM  ShippingCityCNo WHERE  ('" + NoLMCylinderNOTextBox.Text.Trim() + "' >= SCNO) AND ('" + NoLMCylinderNOTextBox.Text.Trim() + "' <= ECNO)";
            conn = new SqlConnection(myConnectionString);
            conn.Open();
            cmd = new SqlCommand(selectCmd, conn);
            reader = cmd.ExecuteReader();
            if (reader.Read())
            {
                if (Client == reader.GetString(0) && City == reader.GetString(1))
                {
                    reader.Close();
                    conn.Close();
                    return true;
                }
                else
                {
                    MessageBox.Show("該序號歸屬" + reader.GetString(1));
                    reader.Close();
                    conn.Close();
                    return false;
                }
            }
            reader.Close();
            conn.Close();
            //SELECT   Client, City, SCNO, ECNO FROM  ShippingCityCNo WHERE  ('5347835T' >= SCNO) AND ('5347835T' <= ECNO)
            return true;
        }

        private bool CheckCylinderNOTextBox()
        {
            if (NoLMCylinderNOTextBox.Text.Length < 6 || NoLMCylinderNOTextBox.Text.Length > 10)
            {
                MessageBox.Show("所輸入之氣瓶序號長度錯誤，請重新輸入!", "提示");
                return false;
            }
            //if (Convert.ToChar(NoLMCylinderNOTextBox.Text.Substring(0, 1)) < 64 || Convert.ToChar(NoLMCylinderNOTextBox.Text.Substring(0, 1)) > 90)
            //{
            //    MessageBox.Show("所輸入之氣瓶序號格式(第一碼)錯誤，請重新輸入!", "提示");
            //    return false;
            //}
            //if ((Convert.ToChar(NoLMCylinderNOTextBox.Text.Substring(1, 1)) < 64 || Convert.ToChar(NoLMCylinderNOTextBox.Text.Substring(1, 1)) > 90) && (Convert.ToChar(NoLMCylinderNOTextBox.Text.Substring(1, 1)) < 48 || Convert.ToChar(NoLMCylinderNOTextBox.Text.Substring(1, 1)) > 57))
            //{
            //    MessageBox.Show("所輸入之氣瓶序號格式(第二碼)錯誤，請重新輸入!", "提示");
            //    return false;
            //}
            //for (int i = 0; i < NoLMCylinderNOTextBox.Text.Substring(2, NoLMCylinderNOTextBox.Text.Length - 2).Length; i++)
            //{
            //    if ((Convert.ToChar(NoLMCylinderNOTextBox.Text.Substring(2+i, 1)) < 48 || Convert.ToChar(NoLMCylinderNOTextBox.Text.Substring(2+i, 1)) > 57))
            //    {
            //        MessageBox.Show("所輸入之氣瓶序號格式(第"+(3+i)+"碼)錯誤，請重新輸入!", "提示");
            //        return false;
            //    }
            //}
            return true;
        }

        private void AutoAccumulate()
        {
            string HydrostaticPass = "N";

            //判別是否為報廢氣瓶
            selectCmd = "SELECT  * FROM [ComplexScrapData] where [ComplexCylinderNo]='" + NoLMCylinderNOTextBox.Text + "'";
            conn = new SqlConnection(myConnectionString);
            conn.Open();
            cmd = new SqlCommand(selectCmd, conn);
            reader = cmd.ExecuteReader();
            if (reader.Read())
            {
                MessageBox.Show("此序號之氣瓶為報廢氣瓶，故不允許加入", "警告-W006");
                reader.Close();
                conn.Close();
                return;
            }
            reader.Close();
            conn.Close();

            selectCmd = "SELECT  * FROM [RePortScrapReason] where [ScrapCylinderNO]='" + NoLMCylinderNOTextBox.Text + "'";
            conn = new SqlConnection(myConnectionString);
            conn.Open();
            cmd = new SqlCommand(selectCmd, conn);
            reader = cmd.ExecuteReader();
            if (reader.Read())
            {
                MessageBox.Show("此序號之氣瓶為報廢氣瓶，故不允許加入", "警告-W006");
                reader.Close();
                conn.Close();
                return;
            }
            reader.Close();
            conn.Close();

            //判斷是否已經有相同的序號入嘜頭

            selectCmd = "SELECT  * FROM [ShippingBody] where [CylinderNumbers]='" + NoLMCylinderNOTextBox.Text + "'";
            conn = new SqlConnection(myConnectionString);
            conn.Open();
            cmd = new SqlCommand(selectCmd, conn);
            reader = cmd.ExecuteReader();
            if (reader.Read())
            {
                MessageBox.Show("此序號已存入嘜頭資訊！(在第" + reader.GetString(4) + "箱，第" + reader.GetString(5) + "位置)", "警告-W001");
                // MessageBox.Show("此序號已存入嘜頭資訊！", "警告-W004");
                // NextNumber();//序號往下累加
                reader.Close();
                conn.Close();
                return;

            }
            reader.Close();
            conn.Close();

            string ManufacturingNo = "";
            string SpecialUses = "N";

            //取得製造批號

            selectCmd = "SELECT  [MSNBody].vchManufacturingNo,isnull([H_SpecialUses],'N') FROM [MSNBody], [Manufacturing]  where [MSNBody].[CylinderNo]='" + NoLMCylinderNOTextBox.Text + "' and [MSNBody].vchManufacturingNo=[Manufacturing].[Manufacturing_NO] ";
            conn = new SqlConnection(myConnectionString);
            conn.Open();
            cmd = new SqlCommand(selectCmd, conn);
            reader = cmd.ExecuteReader();
            if (reader.Read())
            {
                ManufacturingNo = reader.GetString(0);
                if (reader.GetValue(1).ToString() == "Y")
                {
                    SpecialUses = "Y";
                }
            }
            reader.Close();
            conn.Close();

            //selectCmd = "SELECT  * FROM [MSNBody] where [vchCylinderCode]+[vchCylinderNo]='" + NoLMCylinderNOTextBox.Text + "'";
            //conn = new SqlConnection(myConnectionString);
            //conn.Open();
            //cmd = new SqlCommand(selectCmd, conn);
            //reader = cmd.ExecuteReader();
            //if (reader.Read())
            //{
            //    ManufacturingNo = reader.GetString(0);
            //}
            //reader.Close();
            //conn.Close();

            //if (ManufacturingNo != "")
            //{
            //    //判斷此批號是否是走特採的批號
            //    myConnection = new SqlConnection(myConnectionString);
            //    selectCmd = "SELECT  * FROM [Manufacturing] where [Manufacturing_NO]='" + ManufacturingNo + "' and [H_SpecialUses]='Y'";
            //    conn = new SqlConnection(myConnectionString);
            //    conn.Open();
            //    cmd = new SqlCommand(selectCmd, conn);
            //    reader = cmd.ExecuteReader();
            //    if (reader.Read())
            //    {
            //        SpecialUses = "Y";
            //    }
            //    reader.Close();
            //    conn.Close();
            //}

            if (SpecialUses == "N")
            {
                myConnection = new SqlConnection(myConnectionString);
                selectCmd = "SELECT  * FROM [HydrostaticPass] where [ManufacturingNo]='" + ManufacturingNo + "' and [CylinderNo]='" + NoLMCylinderNOTextBox.Text + "' and [HydrostaticPass]='Y'";
                conn = new SqlConnection(myConnectionString);
                conn.Open();
                cmd = new SqlCommand(selectCmd, conn);
                reader = cmd.ExecuteReader();
                if (reader.Read())
                {
                    HydrostaticPass = "Y";
                }
                reader.Close();
                if (HydrostaticPass == "N")
                {
                    //找對應的舊序號，若有序號則依此序號查是否有做過水壓
                    string OriCNo = "", OriMNO = "";
                    selectCmd = "SELECT  OriCylinderNo,OriManufacturingNo, NewCylinderNo FROM [ChangeCylinderNo] where [NewManufacturingNo]='" + ManufacturingNo + "' and [NewCylinderNo]='" + NoLMCylinderNOTextBox.Text + "' ";
                    cmd = new SqlCommand(selectCmd, conn);
                    reader = cmd.ExecuteReader();
                    if (reader.Read())
                    {
                        OriCNo = reader.GetString(0);
                        OriMNO = reader.GetString(1);
                    }
                    reader.Close();

                    if (OriCNo != "")
                    {
                        selectCmd = "SELECT  * FROM [HydrostaticPass] where [ManufacturingNo]='" + OriMNO + "' and [CylinderNo]='" + OriCNo + "' and [HydrostaticPass]='Y'";
                        cmd = new SqlCommand(selectCmd, conn);
                        reader = cmd.ExecuteReader();
                        if (reader.Read())
                        {
                            HydrostaticPass = "Y";
                        }
                        reader.Close();
                    }
                }
                conn.Close();

                if (HydrostaticPass == "N")
                {
                    MessageBox.Show("此序號查詢不到水壓測試資料！", "警告-W005");
                    //NextNumber();//序號往下累加
                    return;
                }

                //20170515判別是否有做過成品檢驗，有才允許繼續，否則不允許包裝
                //研發瓶轉正式出貨產品時，有可能之前的研發瓶試認證瓶所以沒有成品檢驗，因此要有成品檢驗的記錄
                if (CheckProductAcceptanceIsWork(ManufacturingNo) == false)
                {
                    MessageBox.Show("此序號查詢不到成品檢驗資料！", "警告-W008");
                    //NextNumber();//序號往下累加
                    return;
                }
            }

            //功性能測試檢查
            if (PerformanceTest(ManufacturingNo, NoLMCylinderNOTextBox.Text) == false)
            {
                return;
            }
            //20160714機制未完成，故先不使用
            ////複合瓶判別
            //conn = new SqlConnection(myConnectionString);
            //conn.Open();
            //selectCmd = "SELECT  Product.Type  FROM Manufacturing,Product where  Manufacturing.Product_NO=Product.Product_No and Manufacturing.Manufacturing_NO='" + ManufacturingNo + "'";

            //cmd = new SqlCommand(selectCmd, conn);
            //reader = cmd.ExecuteReader();
            //if (reader.Read())
            //{
            //    if (reader.IsDBNull(0) == false && reader.GetString(0).CompareTo("複合瓶") == 0)
            //    {
            //        //複合瓶

            //        //判別是否有爆破、循環資料
            //        connP = new SqlConnection(ESIGNmyConnectionString);
            //        connP.Open();
            //        //爆破
            //        selectCmdP = "SELECT  * FROM [PPT_Burst] WHERE [ManufacturingNo] = '" + ManufacturingNo + "'";
            //        cmdP = new SqlCommand(selectCmdP, connP);
            //        readerP = cmdP.ExecuteReader();

            //        if (readerP.HasRows)
            //        {
            //            ;
            //        }
            //        else
            //        {
            //            reader.Close();
            //            conn.Close();

            //            readerP.Close();
            //            connP.Close();
            //            MessageBox.Show("查詢不到複合瓶" + ManufacturingNo + " 爆破資料！", "警告-W004");
            //            return;
            //        }
            //        readerP.Close();

            //        //循環
            //        selectCmdP = "SELECT  * FROM [PPT_Cycling] WHERE [LotNo] = '" + ManufacturingNo + "'";
            //        cmdP = new SqlCommand(selectCmdP, connP);
            //        readerP = cmdP.ExecuteReader();

            //        if (readerP.HasRows)
            //        {
            //            ;
            //        }
            //        else
            //        {
            //            reader.Close();
            //            conn.Close();

            //            readerP.Close();
            //            connP.Close();
            //            MessageBox.Show("查詢不到複合瓶" + ManufacturingNo + " 循環資料！", "警告-W005");
            //            return;
            //        }
            //        readerP.Close();
            //        connP.Close();
            //    }
            //}
            //reader.Close();
            //conn.Close();

            //判斷新增到那個位子

            string NowSeat = "";

            //判斷[ShippingBody]是否有資料
            myConnection = new SqlConnection(myConnectionString);
            selectCmd = "SELECT  * FROM [ShippingBody] where [ListDate]='" + ListDateListBox.SelectedItem + "' and [ProductName]='" + ProductComboBox.SelectedItem + "' and [WhereBox]='" + BoxsListBox.SelectedItem + "' order by Convert(INT,[WhereSeat]) DESC ";
            conn = new SqlConnection(myConnectionString);
            conn.Open();
            cmd = new SqlCommand(selectCmd, conn);
            reader = cmd.ExecuteReader();
            if (reader.Read())
            {
                NowSeat = reader.GetString(5);
                WhereSeatLabel.Text = (Convert.ToInt32(reader.GetString(5)) + 2).ToString();

                if (NowSeat == Aboxof())
                {
                    MessageBox.Show("此嘜頭已滿箱！", "警告-W009");
                    NextBoxs();
                    return;
                }
            }
            else
            {
                NowSeat = "0";
                WhereSeatLabel.Text = (Convert.ToInt32(NowSeat) + 1).ToString();
            }
            reader.Close();
            conn.Close();

            //取得氣瓶批號
            string LotNumber = null;

            myConnection = new SqlConnection(myConnectionString);
            selectCmd = "SELECT [vchManufacturingNo] FROM [MSNBody] where [CylinderNo]='" + NoLMCylinderNOTextBox.Text + "'";
            conn = new SqlConnection(myConnectionString);
            conn.Open();
            cmd = new SqlCommand(selectCmd, conn);
            reader = cmd.ExecuteReader();
            while (reader.Read())
            {
                LotNumber = reader.GetString(0);
            }
            reader.Close();
            conn.Close();

            //雷刻掃描完確認瓶身瓶底相同後載入資料
            myConnection = new SqlConnection(myConnectionString);
            selectCmd = "INSERT INTO [ShippingBody] ([ListDate],[ProductName],[CylinderNumbers],[WhereBox],[WhereSeat],[vchUser],[Time],[LotNumber])VALUES(" + "'" + ListDateListBox.SelectedItem + "'" + "," + "'" + ProductComboBox.SelectedItem + "'" + "," + "'" + NoLMCylinderNOTextBox.Text + "'" + "," + "'" + BoxsListBox.SelectedItem + "'" + "," + "'" + (Convert.ToInt32(NowSeat) + 1) + "'," + "'" + UserListComboBox.Text + "'," + "'" + NowTime() +"',"+"'"+ LotNumber + "')";
            conn = new SqlConnection(myConnectionString);
            conn.Open();
            cmd = new SqlCommand(selectCmd, conn);
            reader = cmd.ExecuteReader();
            reader.Close();
            conn.Close();

            //更新登出時間
            selectCmd = "UPDATE [LoginPackage] SET  [LogoutTime]= '" + DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss") + "' , [IsUpdate]='0' WHERE [ID] = '" + toolStripStatusLabel1.Text + "'";
            conn = new SqlConnection(myConnectionString);
            conn.Open();
            cmd = new SqlCommand(selectCmd, conn);
            cmd.ExecuteNonQuery();
            conn.Close();

            selectCmd = "INSERT INTO [WorkTimePackage] ([CylinderNo],[Operator],[OperatorId],[AddTime],[Date],[WorkType])VALUES(" + "'" + NoLMCylinderNOTextBox.Text + "','" + User + "','" + ID + "','" + DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss") + "','" + DateTime.Now.ToString("yyyy-MM-dd") + "','" + worktype + "')";
            conn = new SqlConnection(myConnectionString);
            conn.Open();
            cmd = new SqlCommand(selectCmd, conn);
            cmd.ExecuteNonQuery();
            conn.Close();

            time = 420;

            if (CustomerBarCodeCheckBox.Checked == true)
            {
                CustomerBarCode CBC = new CustomerBarCode();
                CBC.ProductName = ProductComboBox.SelectedItem.ToString();
                CBC.ListDate = ListDateListBox.SelectedItem.ToString();
                CBC.Boxs = BoxsListBox.SelectedItem.ToString();
                CBC.Location = (Convert.ToInt32(NowSeat) + 1).ToString();
                CBC.ShowDialog();
            }
            if (WeightCheckBox.Checked == true && ComPortcomboBox.SelectedIndex != -1)
            {
                CylinderNoWeight CNW = new CylinderNoWeight();
                CNW.ComPort = ComPortcomboBox.SelectedItem.ToString();
                CNW.ProductName = ProductComboBox.SelectedItem.ToString();
                CNW.ListDate = ListDateListBox.SelectedItem.ToString();
                CNW.Boxs = BoxsListBox.SelectedItem.ToString();
                CNW.Location = (Convert.ToInt32(NowSeat) + 1).ToString();
                CNW.CylinderNo = NoLMCylinderNOTextBox.Text.ToString();
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
            BoxsListBox.SelectedItem = GetNowBoxNo();

            //載入入箱狀況的圖片
            LoadPictrue();

            //載入dataGridView資料
            LoadSQLDate();
        }

        private bool CheckProductAcceptanceIsWork(string ManufacturingNo)
        {
            //20170515判別是否有做過成品檢驗，有才允許繼續，否則不允許包裝
            //研發瓶轉正式出貨產品時，有可能之前的研發瓶試認證瓶所以沒有成品檢驗，因此要有成品檢驗的記錄
            bool ProductAcceptanceIsWork = false;
            conn = new SqlConnection(myConnectionString);
            conn.Open();
            selectCmd = "SELECT   * FROM [QC_ProductAcceptanceHead] where ManufacturingNo='" + ManufacturingNo + "' and QualifiedQuantity>0 and substring(AcceptanceNo,1,8)<='" + DateTime.Now.ToString("yyyyMMdd") + "'";
            cmd = new SqlCommand(selectCmd, conn);
            reader = cmd.ExecuteReader();
            if (reader.Read())
            {
                ProductAcceptanceIsWork = true;
            }
            reader.Close();

            if (ProductAcceptanceIsWork == false)
            {
                string OriMNO = "";
                //找是否有對應之批號，有則依此搜尋是否有做成品檢驗
                selectCmd = "SELECT  OriManufacturingNo FROM [TransformProductNo] where TransManufacturingNo='" + ManufacturingNo + "'";
                cmd = new SqlCommand(selectCmd, conn);
                reader = cmd.ExecuteReader();
                if (reader.Read())
                {
                    OriMNO = reader.GetString(0);
                }
                reader.Close();

                if (OriMNO != "")
                {
                    selectCmd = "SELECT   * FROM [QC_ProductAcceptanceHead] where ManufacturingNo='" + OriMNO + "' and QualifiedQuantity>0 and substring(AcceptanceNo,1,8)<='" + DateTime.Now.ToString("yyyyMMdd") + "'";
                    cmd = new SqlCommand(selectCmd, conn);
                    reader = cmd.ExecuteReader();
                    if (reader.Read())
                    {
                        ProductAcceptanceIsWork = true;
                    }
                    reader.Close();
                }
            }
            conn.Close();
            return ProductAcceptanceIsWork;
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
            //===== Encoding performed here =====

            //產生條碼
            System.Drawing.Image img = MyCode.GetCodeImage(BarCodeData, Code128_Label.Encode.Code128A);
            BarCodePictureBox.Width = img.Width;
            BarCodePictureBox.Image = img;

            //Bitmap bmp = new Bitmap(pictureBox1.Width, pictureBox1.Height);
            //pictureBox1.DrawToBitmap(bmp, new Rectangle(0, 0, pictureBox1.Width, pictureBox1.Height));
            //bmp.Save("C:\\barcode\\" + BarCodeData + i.ToString().PadLeft(5, '0') + ".png", ImageFormat.Jpeg);

            //===================================
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
            //Excel.Application oXL = new Excel.Application();
            //Excel.Workbook oWB;
            //Excel.Worksheet oSheet;


            //    // 停用警告訊息
            //oXL.DisplayAlerts = false;

            //// 加入新的活頁簿
            //oXL.Workbooks.Add(Type.Missing);

            //// 引用第一個活頁簿
            //oWB = oXL.Workbooks[1];

            //// 設定活頁簿焦點
            ////oWB=(Excel.Workbook)oXL.ActiveWorkbook;//.Activate();

            //// 引用第一個工作表
            //oSheet = (Excel.Worksheet)oWB.Worksheets[1];

            //    // 命名工作表的名稱
            //oSheet.Name = "工作表";

            //    // 設定工作表焦點
            ////設定工作表
            //oSheet = (Excel.Worksheet)oWB.ActiveSheet;
            ////oSheet.Activate();
            ////string srcFileName = "";

            ////srcFileName = Application.StartupPath + @".\Book1.xlsx";//EXCEL檔案路徑

            ////try
            ////{
            ////    //產生一個Workbook物件，並加入Application//改成.open以及在()中輸入開啟位子
            ////    oWB = oXL.Workbooks.Open(srcFileName, Type.Missing, Type.Missing, Type.Missing, Type.Missing,
            ////                            Type.Missing, Type.Missing, Type.Missing, Type.Missing,
            ////                            Type.Missing, Type.Missing, Type.Missing, Type.Missing,
            ////                            Type.Missing, Type.Missing);
            ////}
            ////catch
            ////{
            ////    MessageBox.Show(@"找不到EXCEL檔案！", "Warning");
            ////    return;
            ////}
            ////設定工作表
            ////oSheet = (Excel.Worksheet)oWB.ActiveSheet;

            //float PicLeft, PicTop, PicWidth, PicHeight;
            //string PicturePath, PicLocation;

            ////PicLocation = "A2";
            //PicLocation = ((Char)(65)).ToString() + (2).ToString();
            //PicturePath =  @"C:\SerialNoCode\" + NoLMCylinderNOTextBox.Text.ToString() + ".png";

            //Excel.Worksheet xSheet = (Excel.Worksheet)oWB.Sheets[1];
            ////xSheet.Cells[ (1 + 2 * (i / 8)),(i % 8) + 1] = SelectCylinderNoListBox.Items[i].ToString();
            //Excel.Range m_objRange = xSheet.get_Range(PicLocation, Type.Missing);
            //m_objRange.Select();

            //PicLeft = Convert.ToSingle(m_objRange.Left);
            //PicTop = Convert.ToSingle(m_objRange.Top);
            //PicWidth = 63;
            //PicHeight = 35;
            //xSheet.Shapes.AddPicture(PicturePath, Microsoft.Office.Core.MsoTriState.msoFalse, Microsoft.Office.Core.MsoTriState.msoTrue, PicLeft + 4, PicTop + 6, PicWidth, PicHeight);
            
            //oXL.Visible = true;
            ////關閉活頁簿
            ////oWB.Close(false, Type.Missing, Type.Missing);
            ////關閉Excel
            ////oXL.Quit();

            ////釋放Excel資源
            ////System.Runtime.InteropServices.Marshal.ReleaseComObject(oXL);
            //GC.Collect();

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
            PicLocation = ((Char)(65)).ToString() + (2).ToString();
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

        private string TrialCarry(int i)
        {
            String fnum = String.Format("{0:00000}", Convert.ToInt32(i + 1));

            //修改部分氣瓶序號為6碼
            if ((Ebb == "CA" || Ebb == "NA") && i >= 100000)
            {
                fnum = String.Format("{0:000000}", Convert.ToInt32(i + 1));
            }
            Ebb = "";
            return fnum;
        }

        private void NextBoxs()
        {
            //用來自動跳下一箱     

            string BoxsListBoxIndex = "";
            string NowSeat2 = "";

            //此處插入一個跳出式的視窗，詢問是否要列印

            myConnection = new SqlConnection(myConnectionString);
            selectCmd = "SELECT  * FROM [ShippingBody] where [ListDate]='" + ListDateListBox.SelectedItem + "' and [ProductName]='" + ProductComboBox.SelectedItem + "' and [WhereBox]='" + BoxsListBox.SelectedItem + "' order by Convert(INT,[WhereSeat]) DESC ";
            conn = new SqlConnection(myConnectionString);
            conn.Open();
            cmd = new SqlCommand(selectCmd, conn);
            reader = cmd.ExecuteReader();
            if (reader.Read())
            {
                NowSeat2 = reader.GetString(5);
                BoxsListBoxIndex = BoxsListBox.SelectedIndex.ToString();

                reader.Close();
                conn.Close();

                //如果箱號已經超過最大箱數則不自動跳箱
                if ((Convert.ToInt32(BoxsListBoxIndex) >= (BoxsListBox.Items.Count - 1)) && BoxsListBox.Items.Count != 1 && NowSeat2 == Aboxof())
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
                        BoxsListBox.SelectedIndex = (Convert.ToInt32(BoxsListBoxIndex) + 1);
                    }
                    WhereSeatLabel.Text = "1";
                }
            }
            else
            {
                reader.Close();
                conn.Close();
            }
        }

        private string GetNowSeat()
        {

            //判斷新增到那個位子
            string NowSeat = "";

            //判斷[ShippingBody]是否有資料
            myConnection = new SqlConnection(myConnectionString);
            selectCmd = "SELECT  * FROM [ShippingBody] where [ListDate]='" + ListDateListBox.SelectedItem + "' and [ProductName]='" + ProductComboBox.SelectedItem + "' and [WhereBox]='" + BoxsListBox.SelectedItem + "' order by Convert(INT,[WhereSeat]) DESC ";
            conn = new SqlConnection(myConnectionString);
            conn.Open();
            cmd = new SqlCommand(selectCmd, conn);
            reader = cmd.ExecuteReader();
            if (reader.Read())
            {
                NowSeat = reader.GetString(5);
            }
            else
            {
                NowSeat = "1";
            }
            reader.Close();
            conn.Close();

            return NowSeat;
        }

        public string GetNowBoxNo()
        {
            return BoxsListBox.SelectedItem.ToString();
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
            selectCmd = "SELECT [ListDate],[ProductName],[vchBoxs],isnull([CustomerPO],''),isnull([CustomerProductName],''),isnull([CustomerProductNo],'') FROM [ShippingHead] where [ListDate]='" + ListDateListBox.SelectedItem + "' and [ProductName]='" + ProductComboBox.Text + "' and [vchBoxs]='" + BoxsListBox.SelectedItem + "' ";
            conn = new SqlConnection(myConnectionString);
            conn.Open();
            cmd = new SqlCommand(selectCmd, conn);
            reader = cmd.ExecuteReader();
            if (reader.Read())
            {
                QRcodeName1 = reader.GetString(0) + reader.GetString(1) + reader.GetString(2);
            }
            reader.Close();
            conn.Close();
            return QRcodeName1;
        }
        private string QRcodDetailData()
        {
            string QRcodDetail1 = ""; string Aboxof = "";
            string QRClient = "", QRProductName = "", PackingMarks = "";
            // int section = 0;
            //找出客戶資訊
            selectCmd = "SELECT isnull(Client,''),isnull(ProductName,''),isnull(PackingMarks,''),vchAboxof FROM [ShippingHead] where [ListDate]='" + ListDateListBox.SelectedItem + "' and [ProductName]='" + ProductComboBox.Text + "' and [vchBoxs]='" + BoxsListBox.SelectedItem + "' ";
            conn = new SqlConnection(myConnectionString);
            conn.Open();
            cmd = new SqlCommand(selectCmd, conn);
            reader = cmd.ExecuteReader();
            if (reader.Read())
            {
                QRClient = reader.GetValue(0).ToString();
                QRProductName = reader.GetValue(1).ToString();
                //找出外箱嘜頭貼紙是否有客製化需求PackingMarks
                PackingMarks = reader.GetValue(2).ToString();
                Aboxof = reader.GetValue(3).ToString();
            }
            reader.Close();
            conn.Close();

            GetThisBoxMaxCount();
            if ((QRClient.ToUpper().Trim().StartsWith("SGA") == true || QRClient.ToUpper().Trim().StartsWith("Airtanks Limited") == true) && PackingMarks.ToUpper().Trim().StartsWith("SGA") == true && Aboxof == "1")
            {
                string CNO = "", HydroDate = "";
                //find SGA Marking //CustomerQRCode
                //找出序號再找出產品型號，找出Marking
                conn = new SqlConnection(myConnectionString);
                conn.Open();
                selectCmd = "SELECT isnull(CylinderNumbers,'') FROM [ShippingBody] where [ListDate]='" + ListDateListBox.SelectedItem + "' and [ProductName]='" + ProductComboBox.Text + "' and [WhereBox]='" + BoxsListBox.SelectedItem + "' order by Convert(int,WhereSeat)";
                cmd = new SqlCommand(selectCmd, conn);
                reader = cmd.ExecuteReader();
                if (reader.Read())
                {
                    CNO = reader.GetValue(0).ToString();
                }
                reader.Close();

                selectCmd = "SELECT CustomerQRCode.ProductNo, CustomerQRCode.QRCodeContent,vchHydrostaticTestDate,isnull(LogoType,'') FROM  Manufacturing INNER JOIN MSNBody"
                    + " ON Manufacturing.Manufacturing_NO = MSNBody.vchManufacturingNo INNER JOIN CustomerQRCode ON Manufacturing.Product_NO = CustomerQRCode.ProductNo "
                    + "WHERE (MSNBody.[CylinderNo] = '" + CNO + "') ";
                cmd = new SqlCommand(selectCmd, conn);
                reader = cmd.ExecuteReader();
                while (reader.Read())
                {
                    if (reader.GetValue(3).ToString() == "")
                    {
                        // CustomerName, BottleType, BottleCatalog, ProductNo
                        QRcodDetail1 = reader.GetValue(1).ToString();
                        HydroDate = reader.GetValue(2).ToString();
                    }
                    else if (PackingMarks.Contains(reader.GetValue(3).ToString()) == true)
                    {
                        QRcodDetail1 = reader.GetValue(1).ToString();
                        HydroDate = reader.GetValue(2).ToString();
                    }
                }
                reader.Close();
                conn.Close();
                if (CNO != "")
                {
                    MarkSerialNoDataMatrix(CNO);
                    //MarkSerialNoBarCode(CNO);
                }
                //補入序號
                QRcodDetail1 = QRcodDetail1.Replace("##", CNO);
                //補入水壓年月(依序號建立時之水壓年月)
                QRcodDetail1 = QRcodDetail1.Replace("*", HydroDate);
                if (QRcodDetail1.Contains("SERVICE") == true)
                {
                    QRcodDetail1 = QRcodDetail1.Replace("SERVICE" + QRcodDetail1.Split(new string[] { "SERVICE" }, StringSplitOptions.RemoveEmptyEntries)[1].Split('/')[0] + "/", "SERVICE " + (Convert.ToInt32(QRcodDetail1.Split(new string[] { "SERVICE" }, StringSplitOptions.RemoveEmptyEntries)[1].Split('/')[0]) + 15).ToString() + "/");
                }
                else if (QRcodDetail1.Contains("FIN") == true)
                {
                    QRcodDetail1 = QRcodDetail1.Replace("FIN" + QRcodDetail1.Split(new string[] { "FIN" }, StringSplitOptions.RemoveEmptyEntries)[1].Split('/')[0] + "/", "FIN " + (Convert.ToInt32(QRcodDetail1.Split(new string[] { "FIN" }, StringSplitOptions.RemoveEmptyEntries)[1].Split('/')[0]) + 20).ToString() + "/");
                }
                QRcodDetail1 = QRcodDetail1.Replace("+ (Char)13 +", "\n");//轉換成換行符號
            }
            else
            {
                List<string> SerialNoArray = new List<string>();
                SerialNoArray.Clear();
                int Cumulative = 0;

                selectCmd = "SELECT  * FROM [ShippingBody] where [ListDate]='" + ListDateListBox.SelectedItem + "' and [ProductName]='" + ProductComboBox.Text + "' and [WhereBox]='" + BoxsListBox.SelectedItem + "' ORDER BY convert(int,[WhereSeat]) asc ";
                conn = new SqlConnection(myConnectionString);
                conn.Open();
                cmd = new SqlCommand(selectCmd, conn);
                reader = cmd.ExecuteReader();
                while (reader.Read())
                {
                    if (QRClient.Contains("Praxair") == true)
                    {//Praxair ->only CylinderNo
                        SerialNoArray.Add(reader.GetString(3));
                    }
                    else
                    {//AMS Default data
                        SerialNoArray.Add((Cumulative + 1) + " " + reader.GetString(3));
                    }
                    MarkSerialNoDataMatrix(reader.GetString(3));
                    //MarkSerialNoBarCode(reader.GetString(3));

                    Cumulative++;
                }
                reader.Close();
                conn.Close();

                if (QRClient.Contains("Praxair") == false)
                {//AMS Default data
                    selectCmd = "SELECT [ListDate],[ProductName],[vchBoxs],isnull([CustomerPO],''),isnull([CustomerProductName],''),isnull([CustomerProductNo],'') FROM [ShippingHead] where [ListDate]='" + ListDateListBox.SelectedItem + "' and [ProductName]='" + ProductComboBox.Text + "' and [vchBoxs]='" + BoxsListBox.SelectedItem + "' ";
                    conn = new SqlConnection(myConnectionString);
                    conn.Open();
                    cmd = new SqlCommand(selectCmd, conn);
                    reader = cmd.ExecuteReader();
                    if (reader.Read())
                    {
                        QRcodDetail1 = "Part Description:" + reader.GetString(4) + "\r\nPart No. " + reader.GetString(5) + "\r\nQuantity: " + Getcount + " pieces\r\nC/NO. " + BoxsListBox.SelectedItem + "\r\nSerial No.\r\n";
                    }
                    reader.Close();
                    conn.Close();
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

        private void LoadBoxsNo()
        {
            int i = 1;
            
            myConnection = new SqlConnection(myConnectionString);
            selectCmd = "SELECT * FROM [ShippingHead] where [ListDate]='" + ListDateListBox.SelectedItem + "' and [ProductName]='" + ProductComboBox.Text + "' order by convert(int,[vchBoxs]) asc ";
            conn = new SqlConnection(myConnectionString);
            conn.Open();
            cmd = new SqlCommand(selectCmd, conn);
            reader = cmd.ExecuteReader();
            while (reader.Read())
            {
                BoxsArray[i] = reader.GetString(3);
                i++;
            }
            reader.Close();
            conn.Close();
        }

        private void LoadBoxsCount()
        {
            for (int i = 1; i < BoxsArray.Length; i++)
            {
                if (BoxsArray[i] == null)
                {
                    break;
                }
                
                myConnection = new SqlConnection(myConnectionString);
                selectCmd = "SELECT count([WhereBox]) FROM [ShippingBody] where [ListDate]='" + ListDateListBox.SelectedItem + "' and [ProductName]='" + ProductComboBox.Text + "' and [WhereBox]='" + BoxsArray[i] + "'";
                conn = new SqlConnection(myConnectionString);
                conn.Open();
                cmd = new SqlCommand(selectCmd, conn);
                reader = cmd.ExecuteReader();
                if (reader.Read())
                {
                    BoxsCountArray[i] = reader.GetInt32(0);
                }
                reader.Close();
                conn.Close();
            }
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
            selectCmd = "SELECT count([WhereSeat]) FROM [ShippingBody] where [ListDate]='" + ListDateListBox.SelectedItem + "' and [ProductName]='" + ProductComboBox.Text + "' and [WhereBox]='" + BoxsListBox.SelectedItem + "'";
            conn = new SqlConnection(myConnectionString);
            conn.Open();
            cmd = new SqlCommand(selectCmd, conn);
            reader = cmd.ExecuteReader();
            if (reader.Read())
            {
                Getcount = reader.GetInt32(0);
            }
            reader.Close();
            conn.Close();
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
                        if (!((p_Text.Length & 1) == 0)) throw new Exception("128C長度必須是偶數");
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
                        if (!((p_Text.Length & 1) == 0)) throw new Exception("EAN128長度必須是偶數");
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
                            if (_ValueCode.Length == 0) throw new Exception("不正確字元集!" + p_Text.Substring(0, 1).ToString());
                            _Text += _ValueCode;
                            _TextNumb.Add(_Temp);
                            p_Text = p_Text.Remove(0, 1);
                        }
                        break;
                }

                if (_TextNumb.Count == 0) throw new Exception("錯誤的編碼,無資料");
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
                if (m_Code128 == null) return "";
                DataRow[] _Row = m_Code128.Select(p_Code.ToString() + "='" + p_Value + "'");
                if (_Row.Length != 1) throw new Exception("錯誤的編碼" + p_Value.ToString());
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
                if (_Row.Length != 1) throw new Exception("驗效位的編碼錯誤" + p_CodeId.ToString());
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
            /// <summary>
            /// 顯示可見條碼文字 如果小於40 不顯示文字
            /// </summary>
            /// <param name="p_Bitmap">圖形</param>

            /*
            private void GetViewText(Bitmap p_Bitmap, string p_ViewText)
            {
                if (m_ValueFont == null) return;

                Graphics _Graphics = Graphics.FromImage(p_Bitmap);
                SizeF _DrawSize = _Graphics.MeasureString(p_ViewText, m_ValueFont);
                if (_DrawSize.Height > p_Bitmap.Height - 10 || _DrawSize.Width > p_Bitmap.Width)
                {
                    _Graphics.Dispose();
                    return;
                }

                int _StarY = p_Bitmap.Height - (int)_DrawSize.Height;

                _Graphics.FillRectangle(Brushes.White, new Rectangle(0, _StarY, p_Bitmap.Width, (int)_DrawSize.Height));
                _Graphics.DrawString(p_ViewText, m_ValueFont, Brushes.Black, 0, _StarY);
            }
             * */
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
                        if (!((p_Text.Length & 1) == 0)) throw new Exception("128C長度必須是偶數");
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
                        if (!((p_Text.Length & 1) == 0)) throw new Exception("EAN128長度必須是偶數");
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
                            if (_ValueCode.Length == 0) throw new Exception("不正確字元集!" + p_Text.Substring(0, 1).ToString());
                            _Text += _ValueCode;
                            _TextNumb.Add(_Temp);
                            p_Text = p_Text.Remove(0, 1);
                        }
                        break;
                }

                if (_TextNumb.Count == 0) throw new Exception("錯誤的編碼,無資料");
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

            /// <summary>
            /// 獲取目標對應的資料
            /// </summary>
            /// <param name="p_Code">編碼</param>
            /// <param name="p_Value">數值 A b 30</param>
            /// <param name="p_SetID">返回編號</param>
            /// <returns>編碼</returns>
            private string GetValue(Encode p_Code, string p_Value, ref int p_SetID)
            {
                if (m_Code128 == null) return "";
                DataRow[] _Row = m_Code128.Select(p_Code.ToString() + "='" + p_Value + "'");
                if (_Row.Length != 1) throw new Exception("錯誤的編碼" + p_Value.ToString());
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
                if (_Row.Length != 1) throw new Exception("驗效位的編碼錯誤" + p_CodeId.ToString());
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
                ////f_Magnify = (float)((float)m_Width / (float)Magnify);
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
            /// <summary>
            /// 顯示可見條碼文字 如果小於40 不顯示文字
            /// </summary>
            /// <param name="p_Bitmap">圖形</param>

            private void GetViewText(Bitmap p_Bitmap, string p_ViewText)
            {
                if (m_ValueFont == null) return;

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
                if (BoxsListBox.Items.Count != 0)
                {
                    for (int BoxListIndex = 0; BoxListIndex < BoxsListBox.Items.Count; BoxListIndex++)
                    {
                        if (BoxsListBox.Items[BoxListIndex].ToString().CompareTo(SelectListBoxTextBox.Text.ToString()) == 0)
                        {
                            BoxsListBox.SelectedIndex = BoxListIndex;
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
            System.Drawing.Image img = MyCode.GetCodeImage(SerialNo, Code128.Encode.Code128A);

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
            selectCmd = "SELECT  CylinderNumbers, WhereBox, WhereSeat,ISNULL(CustomerBarCode,''),ISNULL(CylinderWeight,'0') FROM [ShippingBody]  where  [ListDate]='" + ListDateListBox.SelectedItem.ToString() + "' and [ProductName]='" + ProductComboBox.SelectedItem.ToString() + "' and CONVERT(datetime, SUBSTRING(Time, 0, 11), 111)>='" + DateTime.Now.ToLocalTime().ToString().Split(' ')[0].ToString() + "' and CONVERT(datetime, SUBSTRING(Time, 0, 11), 111)<='" + DateTime.Now.AddDays(1).ToLocalTime().ToString().Split(' ')[0].ToString() + "' ORDER BY RIGHT(REPLICATE('0', 8) + CAST(SUBSTRING(CylinderNumbers, 3, Len(CylinderNumbers)-2) AS NVARCHAR), 8)";
            conn = new SqlConnection(myConnectionString);
            conn.Open();
            cmd = new SqlCommand(selectCmd, conn);
            reader = cmd.ExecuteReader();
            while (reader.Read())
            {
                CylinderNumbersList.Add(reader.GetString(0));
                WhereBoxList.Add(Convert.ToInt32(reader.GetString(1)));
                WhereSeatList.Add(reader.GetString(2));
                CustomerBarCodeList.Add(reader.GetString(3));
                CylinderWeightList.Add(reader.GetValue(4).ToString());
            }
            reader.Close();
            conn.Close();

            if (CylinderNumbersList.Count == 0)
            {
                MessageBox.Show("無產品名稱:" + ProductComboBox.SelectedItem.ToString() + "、嘜頭日期:" + ListDateListBox.SelectedItem.ToString() + "於今天包裝之資料。");
                return;
            }

            //conn = new SqlConnection(myConnectionString);
            //conn.Open();
            //for (int i = 0; i < CylinderNumbersList.Count; i++)
            //{
            //    selectCmd = "SELECT  vchManufacturingNo FROM [MSNBody]  where [vchMarked]='Y' and [vchCylinderCode]+[vchCylinderNo]='" + CylinderNumbersList[i].ToString() + "'";

            //    cmd = new SqlCommand(selectCmd, conn);
            //    reader = cmd.ExecuteReader();
            //    if (reader.Read())
            //    {
            //        ManufacturingNoList.Add(reader.GetString(0));
            //    }
            //    else
            //    {
            //        ManufacturingNoList.Add("");
            //    }
            //    reader.Close();
            //}
            //conn.Close();

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
            oSheet.Cells[1, 2] = ProductComboBox.SelectedItem.ToString();
            oSheet.Cells[2, 2] = ListDateListBox.SelectedItem.ToString();
            oSheet.Cells[3, 2] = DateTime.Now.ToString("yyyy/MM/dd");

            oSheet2.Cells[1, 2] = ProductComboBox.SelectedItem.ToString();
            oSheet2.Cells[2, 2] = ListDateListBox.SelectedItem.ToString();
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

            //ManufacturingNoList.Sort();
            //int MNOAmount = 0,Location=0;
            //string MNO = "";
            //for (int i = 0; i < CylinderNumbersList.Count; i++)
            //{
            //    if (i == 0)
            //    {
            //        MNOAmount = 1;
            //        MNO = ManufacturingNoList[i].ToString();
            //    }
            //    else if (MNO == ManufacturingNoList[i].ToString() && i != CylinderNumbersList.Count - 1)
            //    {
            //        MNOAmount += 1;
            //    }
            //    else if (MNO != ManufacturingNoList[i].ToString() && i != CylinderNumbersList.Count - 1)
            //    {
            //        oSheet.Cells[7 + Location, 1] = MNO;
            //        oSheet.Cells[7 + Location, 2] = MNOAmount.ToString();
            //        Location++;
            //        MNOAmount = 1;
            //        MNO = ManufacturingNoList[i].ToString();
            //    }
            //    else if (MNO != ManufacturingNoList[i].ToString() && i == CylinderNumbersList.Count - 1)
            //    {
            //        oSheet.Cells[7 + Location, 1] = MNO;
            //        oSheet.Cells[7 + Location, 2] = MNOAmount.ToString();
            //        Location++;
            //        oSheet.Cells[7 + Location, 1] = 1;
            //        oSheet.Cells[7 + Location, 2] = ManufacturingNoList[i].ToString();
            //    }
            //    else if (MNO == ManufacturingNoList[i].ToString() && i == CylinderNumbersList.Count - 1)
            //    {
            //        oSheet.Cells[7 + Location, 1] = MNO;
            //        oSheet.Cells[7 + Location, 2] = MNOAmount.ToString();
            //    }
            //}

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

        private void FirstPrinterComboBox_SelectedIndexChanged(object sender, EventArgs e)
        {
            IsChangePrinter = true;
        }

        private void PrinterRefreshButton_Click(object sender, EventArgs e)
        {
            LoadPrinter();
        }

        private void Main_FormClosing(object sender, FormClosingEventArgs e)
        {
            if(time>0)
            {
                DialogResult dr = MessageBox.Show("是否確定要關閉程式? \n Do you really want to exit?", "關閉程式  Exit", MessageBoxButtons.YesNo);

                if (dr == DialogResult.Yes)
                {
                    try
                    {
                        //更新登出時間
                        selectCmd = "UPDATE [LoginPackage] SET  [LogoutTime]= '" + DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss") + "' WHERE [ID] = '" + toolStripStatusLabel1.Text + "'";
                        conn = new SqlConnection(myConnectionString);
                        conn.Open();
                        cmd = new SqlCommand(selectCmd, conn);
                        cmd.ExecuteNonQuery();
                        conn.Close();
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
            else if(time<=0)
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

            if (this.ListDateListBox.SelectedIndex != -1)
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
            if (time>0)
            {
                time=time-1;
            }

            if(time==0)
            {
                try
                {
                    //更新登出時間
                    selectCmd = "UPDATE [LoginPackage] SET  [LogoutTime]= '" + DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss") + "' WHERE [ID] = '" + toolStripStatusLabel1.Text + "'";
                    conn = new SqlConnection(myConnectionString);
                    conn.Open();
                    cmd = new SqlCommand(selectCmd, conn);
                    cmd.ExecuteNonQuery();
                    conn.Close();
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