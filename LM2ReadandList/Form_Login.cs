using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Windows.Forms;
using System.Data.SqlClient;

namespace LM2ReadandList
{
    public partial class Form_Login : Form
    {
        //SQL參數
        SqlConnection conn;
        SqlCommand cmd;
        SqlDataReader reader;
        string AMS21_HR_ConnectionString = "Server = 192.168.0.21; DataBase = AMS_HR; uid = sa; pwd = dsc;";
        //string ESIGNmyConnectionString;
        string selectCmd;

        //員工資訊
        string EmpName;
        string EmpNo;

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
    }
}
