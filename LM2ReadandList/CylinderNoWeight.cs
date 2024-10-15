using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Text;
using System.Windows.Forms;
using System.Data.SqlClient;
using System.IO.Ports;

namespace LM2ReadandList
{
    public partial class CylinderNoWeight : Form
    {
        //資料庫宣告
        string myConnectionString = "Server=192.168.0.15;database=amsys;uid=sa;pwd=ams.sql;";
        string selectCmd;
        SqlConnection conn;
        SqlCommand cmd;
        SqlDataReader reader;

        public string ComPort = "", CylinderNo = "", ListDate = "", Boxs = "", check = "";
        public bool HaveBase;
        string ReadWeight = "";
        int i = 0;
        public bool stop = false;

        public CylinderNoWeight()
        {
            InitializeComponent();
        }

        private void CylinderNoWeight_Load(object sender, EventArgs e)
        {
            ReadWeight = "";

            ComPortcomboBox.Items.Clear();
            ComPortcomboBox.Items.Add(ComPort);
            ComPortcomboBox.SelectedIndex = 0;
            ComPortcomboBox.Text = ComPort;

            CylinderNoLabel.Text = "";
            CylinderNoLabel.Text = CylinderNo;

            WeightTextBox.Text = "";

            try
            {
                if (SerialPort1.IsOpen == true)
                {
                    SerialPort1.Close();
                }

                SerialPort1.PortName = ComPortcomboBox.SelectedItem.ToString();
                SerialPort1.BaudRate = 2400;
                SerialPort1.Parity = Parity.Even;
                SerialPort1.DataBits = 7;

                SerialPort1.StopBits = StopBits.One;
                SerialPort1.ReadTimeout = 5000;
                SerialPort1.Open();
                SerialPort1.Parity = Parity.Even;
                SerialPort1.DataBits = 7;
                timer1.Enabled = true;
            }
            catch
            {
                timer1.Enabled = true;
            }
        }

        private void ReflashComportButton_Click(object sender, EventArgs e)
        {
            if (ReflashComportButton.Text == "刷新Com Port")
            {
                string[] ports = SerialPort.GetPortNames();
                List<string> listPorts = new List<string>(ports);
                Comparison<string> comparer = delegate(string name1, string name2)
                {

                    int port1 = Convert.ToInt32(name1.Remove(0, 3));
                    int port2 = Convert.ToInt32(name2.Remove(0, 3));
                    return (port1 - port2);

                };

                listPorts.Sort(comparer);
                ComPortcomboBox.Items.Clear();
                this.ComPortcomboBox.Items.AddRange(listPorts.ToArray());
                ReflashComportButton.Text = "確定變更";
                ComPortcomboBox.Enabled = true;
                WeightTextBox.ReadOnly = false;
                ReLoadButton.Enabled = false;
                OKButton.Enabled = false;
            }
            else
            {
                ComPortcomboBox.Enabled = false;
                WeightTextBox.ReadOnly = true;
                ReLoadButton.Enabled = true;
                OKButton.Enabled = true;
                ReflashComportButton.Text = "刷新Com Port";
            }
        }

        private void timer1_Tick(object sender, EventArgs e)
        {
            try
            {
                SerialPort1.Write(Convert.ToChar(81).ToString() + (Convert.ToChar(13)).ToString() + (Convert.ToChar(10)).ToString());//Q+ctrlM
                ReadWeight = SerialPort1.ReadLine().ToString();
            }
            catch
            {
                timer1.Enabled = false;
                MessageBox.Show("Error 01:未與設備連接或連接錯誤");
                return;

            }

            if (ReadWeight.Contains("ST") == true)
            {
                ReadWeight = ReadWeight.Split(',')[1].Split(' ')[0].ToString();

                if (ReadWeight.Substring(0, 1) == "+")
                {
                    //20200903 扣底座重，20240918扣鎖閥重抓平均130
                    if (check == "True") 
                    {
                        if (HaveBase == true)
                        {
                            WeightTextBox.Text = (Convert.ToDouble(ReadWeight.Substring(1, ReadWeight.Length - 1)) - 236).ToString();

                        }
                        else
                        {
                            WeightTextBox.Text = (Convert.ToDouble(ReadWeight.Substring(1, ReadWeight.Length - 1)) - 130).ToString();
                        }
                    }
                    else
                    {
                        if (HaveBase == true)
                        {
                            WeightTextBox.Text = (Convert.ToDouble(ReadWeight.Substring(1, ReadWeight.Length - 1)) - 111).ToString();

                        }
                        else
                        {
                            WeightTextBox.Text = Convert.ToDouble(ReadWeight.Substring(1, ReadWeight.Length - 1)).ToString();
                        }
                    }

                    if (WeightTextBox.Text.ToString().CompareTo("0") == 1)
                    {
                        timer1.Enabled = false;
                        SerialPort1.Close();
                        return;
                    }
                }
            }
        }

        private void CancelButton_Click(object sender, EventArgs e)
        {
            DialogResult result = MessageBox.Show("取消序號:" + CylinderNoLabel.Text + " 重量量測，是否繼續？", "警告", MessageBoxButtons.OKCancel, MessageBoxIcon.Question);

            if (result == DialogResult.Cancel)
            {
                return;
            }
            stop = true;
            this.Close();
        }

        private void ReLoadButton_Click(object sender, EventArgs e)
        {
            try
            {
                if (SerialPort1.IsOpen == true)
                {
                    SerialPort1.Close();
                }

                SerialPort1.PortName = ComPortcomboBox.SelectedItem.ToString();
                SerialPort1.BaudRate = 2400;
                SerialPort1.Parity = Parity.Even;
                SerialPort1.DataBits = 7;

                SerialPort1.StopBits = StopBits.One;
                SerialPort1.ReadTimeout = 10000;
                SerialPort1.Open();
                SerialPort1.Parity = Parity.Even;
                SerialPort1.DataBits = 7;
                timer1.Enabled = true;
            }
            catch
            {
                timer1.Enabled = true;
            }
        }

        private void WeightTextBox_KeyPress(object sender, KeyPressEventArgs e)
        {
            if (e.KeyChar == (char)13)
            {
                SaveWeight();
            }
        }

        private void OKButton_Click(object sender, EventArgs e)
        {
            SaveWeight();
        }

        private void SaveWeight()
        {
            //判別是否有重量，若無，則不允許處理
            if (WeightTextBox.Text == "0" || WeightTextBox.Text == "")
            {
                MessageBox.Show("無該氣瓶之重量");
                return;
            }

            /*
            //判別該序號中是否已有資料
            conn = new SqlConnection(myConnectionString);
            conn.Open();

            selectCmd = "SELECT [CylinderWeight] From [ShippingBody]  where [CylinderNumbers]='" + CylinderNo + "'";
            cmd = new SqlCommand(selectCmd, conn);
            reader = cmd.ExecuteReader();
            if (reader.Read())
            {
                ;
            }
            else
            {
                //error
                MessageBox.Show("找不到該氣瓶之裝箱資料，故不做任何動作");
                reader.Close();
                conn.Close();
                return;
            }
            reader.Close();

            selectCmd = "UPDATE[ShippingBody] SET [CylinderWeight]='" + WeightTextBox.Text.Trim().ToString() + "' where [CylinderNumbers]='" + CylinderNo + "'";
            cmd = new SqlCommand(selectCmd, conn);
            cmd.ExecuteNonQuery();
            conn.Close();

            */

            this.Close();
        }
    }
}