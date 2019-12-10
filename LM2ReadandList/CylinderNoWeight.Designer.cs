﻿namespace LM2ReadandList
{
    partial class CylinderNoWeight
    {
        /// <summary>
        /// 設計工具所需的變數。
        /// </summary>
        private System.ComponentModel.IContainer components = null;

        /// <summary>
        /// 清除任何使用中的資源。
        /// </summary>
        /// <param name="disposing">如果應該公開 Managed 資源則為 true，否則為 false。</param>
        protected override void Dispose(bool disposing)
        {
            if (disposing && (components != null))
            {
                components.Dispose();
            }
            base.Dispose(disposing);
        }

        #region Windows Form 設計工具產生的程式碼

        /// <summary>
        /// 此為設計工具支援所需的方法 - 請勿使用程式碼編輯器修改這個方法的內容。
        ///
        /// </summary>
        private void InitializeComponent()
        {
            this.components = new System.ComponentModel.Container();
            this.OKButton = new System.Windows.Forms.Button();
            this.WeightTextBox = new System.Windows.Forms.TextBox();
            this.label1 = new System.Windows.Forms.Label();
            this.label2 = new System.Windows.Forms.Label();
            this.ReLoadButton = new System.Windows.Forms.Button();
            this.ReflashComportButton = new System.Windows.Forms.Button();
            this.ComPortcomboBox = new System.Windows.Forms.ComboBox();
            this.label10 = new System.Windows.Forms.Label();
            this.CylinderNoLabel = new System.Windows.Forms.Label();
            this.label4 = new System.Windows.Forms.Label();
            this.timer1 = new System.Windows.Forms.Timer(this.components);
            this.SerialPort1 = new System.IO.Ports.SerialPort(this.components);
            this.checkBox1 = new System.Windows.Forms.CheckBox();
            this.SuspendLayout();
            // 
            // OKButton
            // 
            this.OKButton.BackColor = System.Drawing.Color.FromArgb(((int)(((byte)(128)))), ((int)(((byte)(255)))), ((int)(((byte)(128)))));
            this.OKButton.Font = new System.Drawing.Font("新細明體", 20.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(136)));
            this.OKButton.Location = new System.Drawing.Point(175, 218);
            this.OKButton.Margin = new System.Windows.Forms.Padding(4, 4, 4, 4);
            this.OKButton.Name = "OKButton";
            this.OKButton.Size = new System.Drawing.Size(171, 55);
            this.OKButton.TabIndex = 0;
            this.OKButton.Text = "確定";
            this.OKButton.UseVisualStyleBackColor = false;
            this.OKButton.Click += new System.EventHandler(this.OKButton_Click);
            // 
            // WeightTextBox
            // 
            this.WeightTextBox.Font = new System.Drawing.Font("新細明體", 20.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(136)));
            this.WeightTextBox.ForeColor = System.Drawing.Color.Black;
            this.WeightTextBox.Location = new System.Drawing.Point(175, 150);
            this.WeightTextBox.Margin = new System.Windows.Forms.Padding(4, 4, 4, 4);
            this.WeightTextBox.Name = "WeightTextBox";
            this.WeightTextBox.ReadOnly = true;
            this.WeightTextBox.Size = new System.Drawing.Size(193, 48);
            this.WeightTextBox.TabIndex = 2;
            this.WeightTextBox.KeyPress += new System.Windows.Forms.KeyPressEventHandler(this.WeightTextBox_KeyPress);
            // 
            // label1
            // 
            this.label1.AutoSize = true;
            this.label1.Font = new System.Drawing.Font("新細明體", 21.75F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(136)));
            this.label1.Location = new System.Drawing.Point(107, 11);
            this.label1.Margin = new System.Windows.Forms.Padding(4, 0, 4, 0);
            this.label1.Name = "label1";
            this.label1.Size = new System.Drawing.Size(239, 37);
            this.label1.TabIndex = 3;
            this.label1.Text = "氣瓶重量量測";
            // 
            // label2
            // 
            this.label2.AutoSize = true;
            this.label2.Font = new System.Drawing.Font("新細明體", 20.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(136)));
            this.label2.Location = new System.Drawing.Point(7, 109);
            this.label2.Margin = new System.Windows.Forms.Padding(4, 0, 4, 0);
            this.label2.Name = "label2";
            this.label2.Size = new System.Drawing.Size(151, 34);
            this.label2.TabIndex = 4;
            this.label2.Text = "氣瓶序號";
            // 
            // ReLoadButton
            // 
            this.ReLoadButton.BackColor = System.Drawing.Color.FromArgb(((int)(((byte)(128)))), ((int)(((byte)(128)))), ((int)(((byte)(255)))));
            this.ReLoadButton.Font = new System.Drawing.Font("新細明體", 20.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(136)));
            this.ReLoadButton.Location = new System.Drawing.Point(377, 150);
            this.ReLoadButton.Margin = new System.Windows.Forms.Padding(4, 4, 4, 4);
            this.ReLoadButton.Name = "ReLoadButton";
            this.ReLoadButton.Size = new System.Drawing.Size(101, 50);
            this.ReLoadButton.TabIndex = 48;
            this.ReLoadButton.Text = "重讀";
            this.ReLoadButton.UseVisualStyleBackColor = false;
            this.ReLoadButton.Click += new System.EventHandler(this.ReLoadButton_Click);
            // 
            // ReflashComportButton
            // 
            this.ReflashComportButton.Font = new System.Drawing.Font("新細明體", 14.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(136)));
            this.ReflashComportButton.Location = new System.Drawing.Point(307, 59);
            this.ReflashComportButton.Margin = new System.Windows.Forms.Padding(4, 4, 4, 4);
            this.ReflashComportButton.Name = "ReflashComportButton";
            this.ReflashComportButton.Size = new System.Drawing.Size(173, 41);
            this.ReflashComportButton.TabIndex = 51;
            this.ReflashComportButton.Text = "刷新Com Port";
            this.ReflashComportButton.UseVisualStyleBackColor = true;
            this.ReflashComportButton.Click += new System.EventHandler(this.ReflashComportButton_Click);
            // 
            // ComPortcomboBox
            // 
            this.ComPortcomboBox.AutoCompleteMode = System.Windows.Forms.AutoCompleteMode.SuggestAppend;
            this.ComPortcomboBox.AutoCompleteSource = System.Windows.Forms.AutoCompleteSource.ListItems;
            this.ComPortcomboBox.Enabled = false;
            this.ComPortcomboBox.Font = new System.Drawing.Font("新細明體", 20.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(136)));
            this.ComPortcomboBox.FormattingEnabled = true;
            this.ComPortcomboBox.Location = new System.Drawing.Point(179, 56);
            this.ComPortcomboBox.Margin = new System.Windows.Forms.Padding(4, 4, 4, 4);
            this.ComPortcomboBox.Name = "ComPortcomboBox";
            this.ComPortcomboBox.Size = new System.Drawing.Size(119, 42);
            this.ComPortcomboBox.TabIndex = 50;
            // 
            // label10
            // 
            this.label10.AutoSize = true;
            this.label10.Font = new System.Drawing.Font("新細明體", 20.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(136)));
            this.label10.Location = new System.Drawing.Point(19, 61);
            this.label10.Margin = new System.Windows.Forms.Padding(4, 0, 4, 0);
            this.label10.Name = "label10";
            this.label10.Size = new System.Drawing.Size(140, 34);
            this.label10.TabIndex = 49;
            this.label10.Text = "Com Port";
            // 
            // CylinderNoLabel
            // 
            this.CylinderNoLabel.AutoSize = true;
            this.CylinderNoLabel.Font = new System.Drawing.Font("新細明體", 20.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(136)));
            this.CylinderNoLabel.Location = new System.Drawing.Point(179, 109);
            this.CylinderNoLabel.Margin = new System.Windows.Forms.Padding(4, 0, 4, 0);
            this.CylinderNoLabel.Name = "CylinderNoLabel";
            this.CylinderNoLabel.Size = new System.Drawing.Size(93, 34);
            this.CylinderNoLabel.TabIndex = 52;
            this.CylinderNoLabel.Text = "label3";
            // 
            // label4
            // 
            this.label4.AutoSize = true;
            this.label4.Font = new System.Drawing.Font("新細明體", 20.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(136)));
            this.label4.Location = new System.Drawing.Point(7, 158);
            this.label4.Margin = new System.Windows.Forms.Padding(4, 0, 4, 0);
            this.label4.Name = "label4";
            this.label4.Size = new System.Drawing.Size(151, 34);
            this.label4.TabIndex = 53;
            this.label4.Text = "氣瓶重量";
            // 
            // timer1
            // 
            this.timer1.Interval = 300;
            this.timer1.Tick += new System.EventHandler(this.timer1_Tick);
            // 
            // checkBox1
            // 
            this.checkBox1.AutoSize = true;
            this.checkBox1.Font = new System.Drawing.Font("新細明體", 16.2F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(136)));
            this.checkBox1.Location = new System.Drawing.Point(377, 109);
            this.checkBox1.Name = "checkBox1";
            this.checkBox1.Size = new System.Drawing.Size(90, 32);
            this.checkBox1.TabIndex = 54;
            this.checkBox1.Text = "鎖閥";
            this.checkBox1.UseVisualStyleBackColor = true;
            this.checkBox1.CheckedChanged += new System.EventHandler(this.checkBox1_CheckedChanged);
            // 
            // CylinderNoWeight
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(8F, 15F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.ClientSize = new System.Drawing.Size(489, 294);
            this.Controls.Add(this.checkBox1);
            this.Controls.Add(this.label4);
            this.Controls.Add(this.CylinderNoLabel);
            this.Controls.Add(this.ReflashComportButton);
            this.Controls.Add(this.ComPortcomboBox);
            this.Controls.Add(this.label10);
            this.Controls.Add(this.ReLoadButton);
            this.Controls.Add(this.label2);
            this.Controls.Add(this.label1);
            this.Controls.Add(this.WeightTextBox);
            this.Controls.Add(this.OKButton);
            this.FormBorderStyle = System.Windows.Forms.FormBorderStyle.FixedToolWindow;
            this.Margin = new System.Windows.Forms.Padding(4, 4, 4, 4);
            this.Name = "CylinderNoWeight";
            this.StartPosition = System.Windows.Forms.FormStartPosition.CenterParent;
            this.Text = "氣瓶重量量測";
            this.Load += new System.EventHandler(this.CylinderNoWeight_Load);
            this.ResumeLayout(false);
            this.PerformLayout();

        }

        #endregion

        private System.Windows.Forms.Button OKButton;
        private System.Windows.Forms.TextBox WeightTextBox;
        private System.Windows.Forms.Label label1;
        private System.Windows.Forms.Label label2;
        private System.Windows.Forms.Button ReLoadButton;
        private System.Windows.Forms.Button ReflashComportButton;
        private System.Windows.Forms.ComboBox ComPortcomboBox;
        private System.Windows.Forms.Label label10;
        private System.Windows.Forms.Label CylinderNoLabel;
        private System.Windows.Forms.Label label4;
        private System.Windows.Forms.Timer timer1;
        private System.IO.Ports.SerialPort SerialPort1;
        private System.Windows.Forms.CheckBox checkBox1;
    }
}