namespace LM2ReadandList
{
    partial class AutoAccumulate
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
            this.CylinderLabel = new System.Windows.Forms.Label();
            this.label2 = new System.Windows.Forms.Label();
            this.label3 = new System.Windows.Forms.Label();
            this.label4 = new System.Windows.Forms.Label();
            this.CylinderNOLabel = new System.Windows.Forms.Label();
            this.WhereBoxLabel = new System.Windows.Forms.Label();
            this.WhereSeatLabel = new System.Windows.Forms.Label();
            this.SuspendLayout();
            // 
            // CylinderLabel
            // 
            this.CylinderLabel.AutoSize = true;
            this.CylinderLabel.Font = new System.Drawing.Font("新細明體", 26.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(136)));
            this.CylinderLabel.Location = new System.Drawing.Point(12, 42);
            this.CylinderLabel.Name = "CylinderLabel";
            this.CylinderLabel.Size = new System.Drawing.Size(190, 35);
            this.CylinderLabel.TabIndex = 0;
            this.CylinderLabel.Text = "氣瓶序號：";
            // 
            // label2
            // 
            this.label2.AutoSize = true;
            this.label2.Font = new System.Drawing.Font("新細明體", 26.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(136)));
            this.label2.Location = new System.Drawing.Point(12, 93);
            this.label2.Name = "label2";
            this.label2.Size = new System.Drawing.Size(190, 35);
            this.label2.TabIndex = 1;
            this.label2.Text = "裝入箱號：";
            // 
            // label3
            // 
            this.label3.AutoSize = true;
            this.label3.Font = new System.Drawing.Font("新細明體", 26.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(136)));
            this.label3.Location = new System.Drawing.Point(12, 144);
            this.label3.Name = "label3";
            this.label3.Size = new System.Drawing.Size(190, 35);
            this.label3.TabIndex = 2;
            this.label3.Text = "裝入箱位：";
            // 
            // label4
            // 
            this.label4.AutoSize = true;
            this.label4.Font = new System.Drawing.Font("新細明體", 26.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(136)));
            this.label4.Location = new System.Drawing.Point(12, 259);
            this.label4.Name = "label4";
            this.label4.Size = new System.Drawing.Size(426, 35);
            this.label4.TabIndex = 3;
            this.label4.Text = "※按ENTER繼續、ESC離開";
            // 
            // CylinderNOLabel
            // 
            this.CylinderNOLabel.AutoSize = true;
            this.CylinderNOLabel.Font = new System.Drawing.Font("新細明體", 26.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(136)));
            this.CylinderNOLabel.Location = new System.Drawing.Point(187, 42);
            this.CylinderNOLabel.Name = "CylinderNOLabel";
            this.CylinderNOLabel.Size = new System.Drawing.Size(252, 35);
            this.CylinderNOLabel.TabIndex = 4;
            this.CylinderNOLabel.Text = "CylinderNOLabel";
            // 
            // WhereBoxLabel
            // 
            this.WhereBoxLabel.AutoSize = true;
            this.WhereBoxLabel.Font = new System.Drawing.Font("新細明體", 26.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(136)));
            this.WhereBoxLabel.Location = new System.Drawing.Point(187, 93);
            this.WhereBoxLabel.Name = "WhereBoxLabel";
            this.WhereBoxLabel.Size = new System.Drawing.Size(232, 35);
            this.WhereBoxLabel.TabIndex = 5;
            this.WhereBoxLabel.Text = "WhereBoxLabel";
            // 
            // WhereSeatLabel
            // 
            this.WhereSeatLabel.AutoSize = true;
            this.WhereSeatLabel.Font = new System.Drawing.Font("新細明體", 26.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(136)));
            this.WhereSeatLabel.Location = new System.Drawing.Point(187, 144);
            this.WhereSeatLabel.Name = "WhereSeatLabel";
            this.WhereSeatLabel.Size = new System.Drawing.Size(235, 35);
            this.WhereSeatLabel.TabIndex = 6;
            this.WhereSeatLabel.Text = "WhereSeatLabel";
            // 
            // AutoAccumulate
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 12F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.ClientSize = new System.Drawing.Size(514, 303);
            this.Controls.Add(this.WhereSeatLabel);
            this.Controls.Add(this.WhereBoxLabel);
            this.Controls.Add(this.CylinderNOLabel);
            this.Controls.Add(this.label4);
            this.Controls.Add(this.label3);
            this.Controls.Add(this.label2);
            this.Controls.Add(this.CylinderLabel);
            this.Name = "AutoAccumulate";
            this.Text = "AutoAccumulate";
            this.Load += new System.EventHandler(this.AutoAccumulate_Load);
            this.KeyDown += new System.Windows.Forms.KeyEventHandler(this.AutoAccumulate_KeyDown);
            this.ResumeLayout(false);
            this.PerformLayout();

        }

        #endregion

        private System.Windows.Forms.Label label2;
        private System.Windows.Forms.Label label3;
        private System.Windows.Forms.Label label4;
        internal System.Windows.Forms.Label CylinderLabel;
        internal System.Windows.Forms.Label CylinderNOLabel;
        internal System.Windows.Forms.Label WhereBoxLabel;
        internal System.Windows.Forms.Label WhereSeatLabel;
    }
}