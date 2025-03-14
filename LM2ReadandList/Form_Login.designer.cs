namespace LM2ReadandList
{
    partial class Form_Login
    {
        /// <summary>
        /// Required designer variable.
        /// </summary>
        private System.ComponentModel.IContainer components = null;

        /// <summary>
        /// Clean up any resources being used.
        /// </summary>
        /// <param name="disposing">true if managed resources should be disposed; otherwise, false.</param>
        protected override void Dispose(bool disposing)
        {
            if (disposing && (components != null))
            {
                components.Dispose();
            }
            base.Dispose(disposing);
        }

        #region Windows Form Designer generated code

        /// <summary>
        /// Required method for Designer support - do not modify
        /// the contents of this method with the code editor.
        /// </summary>
        private void InitializeComponent()
        {
            this.panel1 = new System.Windows.Forms.Panel();
            this.label2 = new System.Windows.Forms.Label();
            this.label5 = new System.Windows.Forms.Label();
            this.label1 = new System.Windows.Forms.Label();
            this.AccountTextBox = new System.Windows.Forms.TextBox();
            this.Password_TextB = new System.Windows.Forms.TextBox();
            this.Login_B = new System.Windows.Forms.Button();
            this.panel1.SuspendLayout();
            this.SuspendLayout();
            // 
            // panel1
            // 
            this.panel1.AccessibleDescription = "A";
            this.panel1.Controls.Add(this.label2);
            this.panel1.Controls.Add(this.label5);
            this.panel1.Controls.Add(this.label1);
            this.panel1.Controls.Add(this.AccountTextBox);
            this.panel1.Controls.Add(this.Password_TextB);
            this.panel1.Location = new System.Drawing.Point(24, 12);
            this.panel1.Name = "panel1";
            this.panel1.Size = new System.Drawing.Size(288, 238);
            this.panel1.TabIndex = 2;
            // 
            // label2
            // 
            this.label2.AccessibleDescription = "A";
            this.label2.AutoSize = true;
            this.label2.Font = new System.Drawing.Font("新細明體", 15.75F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(136)));
            this.label2.Location = new System.Drawing.Point(22, 32);
            this.label2.Margin = new System.Windows.Forms.Padding(4, 0, 4, 0);
            this.label2.Name = "label2";
            this.label2.Size = new System.Drawing.Size(210, 27);
            this.label2.TabIndex = 38;
            this.label2.Text = "工號 Employee ID";
            // 
            // label5
            // 
            this.label5.AccessibleDescription = "A";
            this.label5.AutoSize = true;
            this.label5.Font = new System.Drawing.Font("新細明體", 12F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(136)));
            this.label5.ForeColor = System.Drawing.Color.Red;
            this.label5.Location = new System.Drawing.Point(23, 188);
            this.label5.Margin = new System.Windows.Forms.Padding(4, 0, 4, 0);
            this.label5.Name = "label5";
            this.label5.Size = new System.Drawing.Size(229, 20);
            this.label5.TabIndex = 41;
            this.label5.Text = "※工號與密碼有分大小寫";
            // 
            // label1
            // 
            this.label1.AccessibleDescription = "A";
            this.label1.AutoSize = true;
            this.label1.Font = new System.Drawing.Font("新細明體", 15.75F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(136)));
            this.label1.Location = new System.Drawing.Point(22, 112);
            this.label1.Margin = new System.Windows.Forms.Padding(4, 0, 4, 0);
            this.label1.Name = "label1";
            this.label1.Size = new System.Drawing.Size(170, 27);
            this.label1.TabIndex = 39;
            this.label1.Text = "密碼 Password";
            // 
            // AccountTextBox
            // 
            this.AccountTextBox.AccessibleDescription = "A";
            this.AccountTextBox.CharacterCasing = System.Windows.Forms.CharacterCasing.Upper;
            this.AccountTextBox.Font = new System.Drawing.Font("新細明體", 15.75F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(136)));
            this.AccountTextBox.Location = new System.Drawing.Point(27, 68);
            this.AccountTextBox.Margin = new System.Windows.Forms.Padding(4);
            this.AccountTextBox.Name = "AccountTextBox";
            this.AccountTextBox.Size = new System.Drawing.Size(232, 39);
            this.AccountTextBox.TabIndex = 0;
            // 
            // Password_TextB
            // 
            this.Password_TextB.AccessibleDescription = "A";
            this.Password_TextB.Font = new System.Drawing.Font("新細明體", 15.75F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(136)));
            this.Password_TextB.Location = new System.Drawing.Point(27, 142);
            this.Password_TextB.Margin = new System.Windows.Forms.Padding(4);
            this.Password_TextB.Name = "Password_TextB";
            this.Password_TextB.PasswordChar = '*';
            this.Password_TextB.Size = new System.Drawing.Size(232, 39);
            this.Password_TextB.TabIndex = 1;
            this.Password_TextB.KeyDown += new System.Windows.Forms.KeyEventHandler(this.Password_TextB_KeyDown);
            // 
            // Login_B
            // 
            this.Login_B.AccessibleDescription = "A";
            this.Login_B.Font = new System.Drawing.Font("新細明體", 15.75F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(136)));
            this.Login_B.Location = new System.Drawing.Point(68, 257);
            this.Login_B.Margin = new System.Windows.Forms.Padding(4);
            this.Login_B.Name = "Login_B";
            this.Login_B.Size = new System.Drawing.Size(177, 62);
            this.Login_B.TabIndex = 3;
            this.Login_B.TabStop = false;
            this.Login_B.Text = "登入 Login";
            this.Login_B.UseVisualStyleBackColor = true;
            this.Login_B.Click += new System.EventHandler(this.Login_B_Click);
            // 
            // Form_Login
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(8F, 15F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.ClientSize = new System.Drawing.Size(333, 349);
            this.Controls.Add(this.panel1);
            this.Controls.Add(this.Login_B);
            this.Name = "Form_Login";
            this.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen;
            this.Text = "登入(Login)";
            this.Load += new System.EventHandler(this.Form_Login_Load);
            this.panel1.ResumeLayout(false);
            this.panel1.PerformLayout();
            this.ResumeLayout(false);

        }

        #endregion

        private System.Windows.Forms.Panel panel1;
        private System.Windows.Forms.Label label2;
        private System.Windows.Forms.Label label5;
        private System.Windows.Forms.Label label1;
        private System.Windows.Forms.TextBox AccountTextBox;
        private System.Windows.Forms.TextBox Password_TextB;
        private System.Windows.Forms.Button Login_B;
    }
}