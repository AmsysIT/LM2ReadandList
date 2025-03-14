using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Windows.Forms;
using System.Data.SqlClient;
using LM2ReadandList_Customized.API;

namespace LM2ReadandList
{
    public partial class Form_Login : Form
    {
        //SQL參數
        SqlConnection conn;
        SqlCommand cmd;
        SqlDataReader reader;
        string AMS21_HR_ConnectionString { get; set; }
        //string ESIGNmyConnectionString;
        string selectCmd;

        //員工資訊
        string EmpName;
        string EmpNo;

        public static bool azure_mode { get; set; }

        public Form_Login()
        {
            InitializeComponent();
        }

        private void Password_TextB_KeyDown(object sender, KeyEventArgs e)
        {
            // Press ENTER
            if (e.KeyCode == Keys.Enter)
            {
                Login_B.PerformClick();
            }
        }

        private void Login_B_Click(object sender, EventArgs e)
        {
            string ErrorMsg = string.Empty;
            string Password = string.Empty;

            // Check login 
            selectCmd = "SELECT * FROM [Employee]  WHERE [Onjob] = '1' AND [EmployeeNo] = @EmployeeNo  " ;
            conn = new SqlConnection(AMS21_HR_ConnectionString);
            conn.Open();
            cmd = new SqlCommand(selectCmd, conn);
            cmd.Parameters.Add("@EmployeeNo", SqlDbType.VarChar).Value = AccountTextBox.Text;
            cmd.Parameters.Add("@Password", SqlDbType.VarChar).Value = Password_TextB.Text;
            reader = cmd.ExecuteReader();

            if (reader.HasRows)
            {
                if (reader.Read())
                {
                    EmpNo = reader.GetString(reader.GetOrdinal("EmployeeNo"));
                    EmpName = reader.GetString(reader.GetOrdinal("Name"));
                    Password = reader.GetString(reader.GetOrdinal("Password"));

                    if (Password != Password_TextB.Text)
                    {
                        ErrorMsg = "密碼錯誤。Password Error。";
                    }
                }
            }
            else
            {
                MessageBox.Show("查無此工號。EmployeeNo Error。", "警告");
                Password_TextB.Text = "";
                return;
            }

            if (ErrorMsg.Any())
            {
                MessageBox.Show(ErrorMsg, "AMSYS");
                return;
            }

            Main MA = new Main
            {
                EmpName = EmpName,
                EmpNo = EmpNo,
            };
            //Hide(); //20241031
            MA.ShowDialog();
            this.Close(); //20241031

        }

        public void Init_ConnectionString()
        {
            //TODO remove
            AMS21_HR_ConnectionString = Api_Core.get_connectstring(db_name: "AMS_HR", test_mode: azure_mode);
            Console.WriteLine(this.AMS21_HR_ConnectionString);
        }


        private void Form_Login_Load(object sender, EventArgs e)
        {
#if DEBUG
            //開發端自行選擇模式
            var result = MessageBox.Show("是否使用雲端字串模式?", "確認", MessageBoxButtons.YesNo, MessageBoxIcon.Question);
            if (result == DialogResult.Yes)
            {
                azure_mode = true;
            }
            AccountTextBox.Text = "A00699";
            Password_TextB.Text = "Yz54338923";
#else
            //用戶端則使用API取得模式
            azure_mode = Api_Core.get_status();
#endif
            if (azure_mode)
            {
                Text += " - 雲端模式";
            }

            Init_ConnectionString();
        }
    }
}
