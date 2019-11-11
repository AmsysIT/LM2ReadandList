using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Text;
using System.Windows.Forms;
using System.Data.SqlClient;
using System.IO;
using System.Reflection;
using Microsoft.Win32;
using System.Text.RegularExpressions;
using Excel = Microsoft.Office.Interop.Excel;

namespace LM2ReadandList
{
    public partial class AutoAccumulate : Form
    {
        public Main F_Main = null;
        string Ebb = "";
        
        //資料庫宣告
        string myConnectionString;
        SqlConnection myConnection;
        string selectCmd;
        SqlConnection conn;
        SqlCommand cmd;
        SqlDataReader reader;


        public AutoAccumulate()
        {
            InitializeComponent();
            myConnectionString = "Server=192.168.0.15;database=amsys;uid=sa;pwd=ams.sql;";       

        }

        private void AutoAccumulate_Load(object sender, EventArgs e)
        {

        }

        private void AutoAccumulate_KeyDown(object sender, KeyEventArgs e)
        {
            if (e.KeyCode == Keys.Escape)
            {
                this.Close();
            }

            if (e.KeyValue == 13)//16=SHIFT 13=ENTER
            {


                // string FredlovCSV = "N";
                //string CalisoCSV = "N";

                string HydrostaticPass = "N";

                
                
                //判斷是否已經有相同的序號入嘜頭
                myConnection = new SqlConnection(myConnectionString);
                selectCmd = "SELECT  * FROM [ShippingBody] where [CylinderNumbers]='"+CylinderNOLabel.Text+"'";
                conn = new SqlConnection(myConnectionString);
                conn.Open();
                cmd = new SqlCommand(selectCmd, conn);
                reader = cmd.ExecuteReader();
                if (reader.Read())
                {                   
                       
                        MessageBox.Show("此序號已存入嘜頭資訊！", "警告-W004");
                        NextNumber();
                        return;
                    
                }
                reader.Close();
                conn.Close();
                
                
                /*
                //用來記錄前面加了幾個零
                int AddZero = 0;

                //用來記錄前面最多偵測補幾個零(X-1)=X-1個0
                int HManyZero = 5;

                //用來存放加零過後的字串
                string AddStr = "";
                
                for (int i = 0; i < HManyZero; i++)
                {

                    switch (AddZero)
                    {
                        case 0:
                            AddStr = GetManufacturingCode(CylinderNOLabel.Text) + GetManufacturingNumber(CylinderNOLabel.Text);
                            break;

                        case 1:
                            AddStr = GetManufacturingCode(CylinderNOLabel.Text) + "0" + GetManufacturingNumber(CylinderNOLabel.Text);
                            break;

                        case 2:
                            AddStr = GetManufacturingCode(CylinderNOLabel.Text) + "00" + GetManufacturingNumber(CylinderNOLabel.Text);
                            break;


                        case 3:
                            AddStr = GetManufacturingCode(CylinderNOLabel.Text) + "000" + GetManufacturingNumber(CylinderNOLabel.Text);
                            break;

                        case 4:
                            AddStr = GetManufacturingCode(CylinderNOLabel.Text) + "0000" + GetManufacturingNumber(CylinderNOLabel.Text);
                            break;


                    }


                    //判斷Fredlov水壓機是否以有測試資料         
                    myConnection = new SqlConnection(myConnectionString);
                    selectCmd = "SELECT  * FROM [FredlovCSV] where [vchCylinderNO]='" + AddStr + "' and [vchStatus]='Pass'";
                    conn = new SqlConnection(myConnectionString);
                    conn.Open();
                    cmd = new SqlCommand(selectCmd, conn);
                    reader = cmd.ExecuteReader();
                    if (reader.Read())
                    {
                        FredlovCSV = "Y";
                        break;
                    }
                    else
                    {

                        if (AddZero < HManyZero)
                        {
                            AddZero++;
                        }

                    }
                    reader.Close();
                    conn.Close();

                }






                //判斷CalisoCSV水壓機是否以有測試資料         
                myConnection = new SqlConnection(myConnectionString);
                selectCmd = "SELECT  * FROM [CalisoCSV] where [vchQuaActTstPrs]>[vchQualTestPres] and [vchQuaActTstPrs]<(convert(int,[vchQualTestPres])*1.1) and (convert(float,[vchQuaActPE]))<'5' and [vchQualDisposit] not like '%F%' and [vchCylindSerial]='" + CylinderNOLabel.Text + "'";
                conn = new SqlConnection(myConnectionString);
                conn.Open();
                cmd = new SqlCommand(selectCmd, conn);
                reader = cmd.ExecuteReader();
                if (reader.Read())
                {
                    CalisoCSV = "Y";
                }
                reader.Close();
                conn.Close();
                 * 
                 */

                string ManufacturingNo = "";

                //取得製造批號
                myConnection = new SqlConnection(myConnectionString);
                selectCmd = "SELECT  * FROM [MSNBody] where [vchCylinderCode]+[vchCylinderNo]='" + CylinderNOLabel.Text + "'";
                conn = new SqlConnection(myConnectionString);
                conn.Open();
                cmd = new SqlCommand(selectCmd, conn);
                reader = cmd.ExecuteReader();
                if (reader.Read())
                {
                    ManufacturingNo = reader.GetString(0);
                }
                reader.Close();
                conn.Close();

                string SpecialUses = "N";

                if (ManufacturingNo != "")
                {
                    //判斷此批號是否是走特採的批號
                    myConnection = new SqlConnection(myConnectionString);
                    selectCmd = "SELECT  * FROM [Manufacturing] where [Manufacturing_NO]='" + ManufacturingNo + "' and [H_SpecialUses]='Y'";
                    conn = new SqlConnection(myConnectionString);
                    conn.Open();
                    cmd = new SqlCommand(selectCmd, conn);
                    reader = cmd.ExecuteReader();
                    if (reader.Read())
                    {
                        SpecialUses = "Y";
                    }
                    reader.Close();
                    conn.Close();

                }


                if (SpecialUses == "N")
                {

                    myConnection = new SqlConnection(myConnectionString);
                    selectCmd = "SELECT  * FROM [HydrostaticPass] where [ManufacturingNo]='" + ManufacturingNo + "' and [CylinderNo]='" + CylinderNOLabel.Text + "' and [HydrostaticPass]='Y'";
                    conn = new SqlConnection(myConnectionString);
                    conn.Open();
                    cmd = new SqlCommand(selectCmd, conn);
                    reader = cmd.ExecuteReader();
                    if (reader.Read())
                    {
                        HydrostaticPass = "Y";
                    }
                    reader.Close();
                    conn.Close();


                    if (HydrostaticPass == "N")
                    {
                        MessageBox.Show("此序號查詢不到水壓測試資料！", "警告-W005");
                        NextNumber();
                        return;
                    }


                }








                /*

                if (SpecialUses == "N")
                {
                    if ((CalisoCSV == "N") && (FredlovCSV == "N"))
                    {

                        MessageBox.Show("此序號查詢不到水壓測試資料！", "警告-W005");
                        NextNumber();
                        return;
                    }


                }
                */
                




                //判斷新增到那個位子

                string NowSeat = "";

                //判斷[ShippingBody]是否有資料
                myConnection = new SqlConnection(myConnectionString);
                selectCmd = "SELECT  * FROM [ShippingBody] where [ListDate]='" + F_Main.ListDateListBox.SelectedItem + "' and [ProductName]='" + F_Main.ProductComboBox.SelectedItem + "' and [WhereBox]='" + F_Main.BoxsListBox.SelectedItem + "' order by Convert(INT,[WhereSeat]) DESC ";
                conn = new SqlConnection(myConnectionString);
                conn.Open();
                cmd = new SqlCommand(selectCmd, conn);
                reader = cmd.ExecuteReader();
                if (reader.Read())
                {
                    NowSeat = reader.GetString(5);
                    WhereSeatLabel.Text = (Convert.ToInt32(reader.GetString(5))+2).ToString();
                   

                    if (NowSeat == F_Main.Aboxof())
                    {
                       
                        MessageBox.Show("此嘜頭以滿箱！", "警告-W006");
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

               
            
          

                //雷刻掃描完確認瓶身瓶底相同後載入資料
                myConnection = new SqlConnection(myConnectionString);
                selectCmd = "INSERT INTO [ShippingBody] ([ListDate],[ProductName],[CylinderNumbers],[WhereBox],[WhereSeat],[vchUser],[Time])VALUES(" + "'" + F_Main.ListDateListBox.SelectedItem + "'" + "," + "'" + F_Main.ProductComboBox.SelectedItem + "'" + "," + "'" + CylinderNOLabel.Text + "'" + "," + "'" + F_Main.BoxsListBox.SelectedItem + "'" + "," + "'" + (Convert.ToInt32(NowSeat) + 1) + "'," + "'" + F_Main.UserListComboBox.Text + "'," + "'" + NowTime() + "')";
                conn = new SqlConnection(myConnectionString);
                conn.Open();
                cmd = new SqlCommand(selectCmd, conn);
                reader = cmd.ExecuteReader();
                reader.Close();
                conn.Close();


                

               //自動跳下一箱 
                NextBoxs();

                //載入目前箱號
                WhereBoxLabel.Text = F_Main.GetNowBoxNo();

                //載入入箱狀況的圖片
                F_Main.LoadPictrue();

                //載入dataGridView資料
                F_Main.LoadSQLDate();
                 
                 

                //序號往下累加
                NextNumber();

                
            }
        }

        private string TrialCarry(int i)
        {
            String fnum = String.Format("{0:00000}", Convert.ToInt32(i+1)); 

           
            //修改部分氣瓶序號為6碼
            if ((Ebb == "CA" || Ebb == "NA") && i >= 100000)
            {
                 fnum = String.Format("{0:000000}", Convert.ToInt32(i + 1)); 
            }
            Ebb = "";
            return fnum;

        }

        private void NextNumber()
        {
            char[] b = new char[12];
            StringReader sr = new StringReader(CylinderNOLabel.Text);
            sr.Read(b, 0, 12);
            sr.Close();

            string Nbb = "";
            int AddNbb = 0;
            for (int i = 0; i <= CylinderNOLabel.Text.Length; i++)
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
            CylinderNOLabel.Text = Ebb + TrialCarry(AddNbb);
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

        private string NowTime()
        {
            //取得現在時間
            DateTime currentTime = DateTime.Now;
            //轉成字串   
            String timeString = currentTime.ToLocalTime().ToString();

            return timeString;

        }


        private void NextBoxs()
        {
            //用來自動跳下一箱     

            string BoxsListBoxIndex = "";
            string NowSeat2 = "";

            //此處插入一個跳出式的視窗，詢問是否要列印


            myConnection = new SqlConnection(myConnectionString);
            selectCmd = "SELECT  * FROM [ShippingBody] where [ListDate]='" + F_Main.ListDateListBox.SelectedItem + "' and [ProductName]='" + F_Main.ProductComboBox.SelectedItem + "' and [WhereBox]='" + F_Main.BoxsListBox.SelectedItem + "' order by Convert(INT,[WhereSeat]) DESC ";
            conn = new SqlConnection(myConnectionString);
            conn.Open();
            cmd = new SqlCommand(selectCmd, conn);
            reader = cmd.ExecuteReader();
            if (reader.Read())
            {
                NowSeat2 = reader.GetString(5);
                BoxsListBoxIndex = F_Main.BoxsListBox.SelectedIndex.ToString();

                //如果箱號已經超過最大箱數則不自動跳箱
                if ((Convert.ToInt32(BoxsListBoxIndex) >= (F_Main.BoxsListBox.Items.Count - 1)) && F_Main.BoxsListBox.Items.Count != 1 && NowSeat2 == F_Main.Aboxof())
                {
                    MessageBox.Show("此日期嘜頭已經完全結束", "提示");
                    return;
                }


                if (NowSeat2 == F_Main.Aboxof())
                {

                    F_Main.BoxsListBox.SelectedIndex = (Convert.ToInt32(BoxsListBoxIndex) + 1);
                    WhereSeatLabel.Text= "1";
                }


            }
            reader.Close();
            conn.Close();
        }
    }
}