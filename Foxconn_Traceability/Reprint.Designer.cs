namespace Foxconn_Traceability
{
    partial class Reprint
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
            this.btnClose = new System.Windows.Forms.Button();
            this.panel1 = new System.Windows.Forms.Panel();
            this.pictureBox1 = new System.Windows.Forms.PictureBox();
            this.groupBox1 = new System.Windows.Forms.GroupBox();
            this.chkTg3442 = new System.Windows.Forms.CheckBox();
            this.lbTextoSerial = new System.Windows.Forms.Label();
            this.txtSerial = new System.Windows.Forms.TextBox();
            this.chkSelecionar = new System.Windows.Forms.CheckBox();
            this.lbRodape = new System.Windows.Forms.Label();
            this.txtWo = new System.Windows.Forms.TextBox();
            this.lbWo = new System.Windows.Forms.Label();
            this.panel1.SuspendLayout();
            ((System.ComponentModel.ISupportInitialize)(this.pictureBox1)).BeginInit();
            this.groupBox1.SuspendLayout();
            this.SuspendLayout();
            // 
            // btnClose
            // 
            this.btnClose.BackColor = System.Drawing.Color.SlateGray;
            this.btnClose.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F, ((System.Drawing.FontStyle)((System.Drawing.FontStyle.Bold | System.Drawing.FontStyle.Italic))), System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.btnClose.ForeColor = System.Drawing.SystemColors.ControlLightLight;
            this.btnClose.Location = new System.Drawing.Point(14, 252);
            this.btnClose.Name = "btnClose";
            this.btnClose.Size = new System.Drawing.Size(81, 28);
            this.btnClose.TabIndex = 11;
            this.btnClose.Text = "Voltar";
            this.btnClose.UseVisualStyleBackColor = false;
            this.btnClose.Click += new System.EventHandler(this.btnClose_Click);
            // 
            // panel1
            // 
            this.panel1.BackColor = System.Drawing.Color.SkyBlue;
            this.panel1.Controls.Add(this.pictureBox1);
            this.panel1.Controls.Add(this.btnClose);
            this.panel1.Controls.Add(this.groupBox1);
            this.panel1.Controls.Add(this.chkSelecionar);
            this.panel1.Location = new System.Drawing.Point(19, 11);
            this.panel1.Name = "panel1";
            this.panel1.Size = new System.Drawing.Size(507, 292);
            this.panel1.TabIndex = 12;
            // 
            // pictureBox1
            // 
            this.pictureBox1.BackColor = System.Drawing.Color.Transparent;
            this.pictureBox1.Location = new System.Drawing.Point(1, 1);
            this.pictureBox1.Name = "pictureBox1";
            this.pictureBox1.Size = new System.Drawing.Size(129, 33);
            this.pictureBox1.TabIndex = 41;
            this.pictureBox1.TabStop = false;
            // 
            // groupBox1
            // 
            this.groupBox1.Controls.Add(this.lbWo);
            this.groupBox1.Controls.Add(this.txtWo);
            this.groupBox1.Controls.Add(this.chkTg3442);
            this.groupBox1.Controls.Add(this.lbTextoSerial);
            this.groupBox1.Controls.Add(this.txtSerial);
            this.groupBox1.Location = new System.Drawing.Point(4, 42);
            this.groupBox1.Name = "groupBox1";
            this.groupBox1.Size = new System.Drawing.Size(499, 150);
            this.groupBox1.TabIndex = 3;
            this.groupBox1.TabStop = false;
            this.groupBox1.Text = "REPRINT";
            // 
            // chkTg3442
            // 
            this.chkTg3442.AutoSize = true;
            this.chkTg3442.Location = new System.Drawing.Point(390, 27);
            this.chkTg3442.Name = "chkTg3442";
            this.chkTg3442.Size = new System.Drawing.Size(97, 17);
            this.chkTg3442.TabIndex = 42;
            this.chkTg3442.Text = "TG3442-SMT?";
            this.chkTg3442.UseVisualStyleBackColor = true;
            // 
            // lbTextoSerial
            // 
            this.lbTextoSerial.AutoSize = true;
            this.lbTextoSerial.Font = new System.Drawing.Font("Microsoft Sans Serif", 12F, ((System.Drawing.FontStyle)((System.Drawing.FontStyle.Bold | System.Drawing.FontStyle.Italic))), System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.lbTextoSerial.Location = new System.Drawing.Point(6, 23);
            this.lbTextoSerial.Name = "lbTextoSerial";
            this.lbTextoSerial.Size = new System.Drawing.Size(60, 20);
            this.lbTextoSerial.TabIndex = 2;
            this.lbTextoSerial.Text = "Serial:";
            // 
            // txtSerial
            // 
            this.txtSerial.Location = new System.Drawing.Point(71, 23);
            this.txtSerial.MaxLength = 18;
            this.txtSerial.Name = "txtSerial";
            this.txtSerial.Size = new System.Drawing.Size(163, 20);
            this.txtSerial.TabIndex = 0;
            this.txtSerial.KeyPress += new System.Windows.Forms.KeyPressEventHandler(this.txtSerial_KeyPress);
            // 
            // chkSelecionar
            // 
            this.chkSelecionar.AutoSize = true;
            this.chkSelecionar.Location = new System.Drawing.Point(14, 210);
            this.chkSelecionar.Name = "chkSelecionar";
            this.chkSelecionar.Size = new System.Drawing.Size(145, 17);
            this.chkSelecionar.TabIndex = 3;
            this.chkSelecionar.Text = "Usar Impressora Padrão?";
            this.chkSelecionar.UseVisualStyleBackColor = true;
            // 
            // lbRodape
            // 
            this.lbRodape.AutoSize = true;
            this.lbRodape.BackColor = System.Drawing.Color.Transparent;
            this.lbRodape.Font = new System.Drawing.Font("Microsoft Sans Serif", 12F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.lbRodape.ForeColor = System.Drawing.Color.White;
            this.lbRodape.Location = new System.Drawing.Point(12, 359);
            this.lbRodape.Name = "lbRodape";
            this.lbRodape.Size = new System.Drawing.Size(66, 20);
            this.lbRodape.TabIndex = 40;
            this.lbRodape.Text = "Rodapé";
            // 
            // txtWo
            // 
            this.txtWo.Location = new System.Drawing.Point(71, 58);
            this.txtWo.MaxLength = 18;
            this.txtWo.Name = "txtWo";
            this.txtWo.Size = new System.Drawing.Size(163, 20);
            this.txtWo.TabIndex = 43;
            // 
            // lbWo
            // 
            this.lbWo.AutoSize = true;
            this.lbWo.Font = new System.Drawing.Font("Microsoft Sans Serif", 12F, ((System.Drawing.FontStyle)((System.Drawing.FontStyle.Bold | System.Drawing.FontStyle.Italic))), System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.lbWo.Location = new System.Drawing.Point(5, 58);
            this.lbWo.Name = "lbWo";
            this.lbWo.Size = new System.Drawing.Size(40, 20);
            this.lbWo.TabIndex = 44;
            this.lbWo.Text = "Wo:";
            // 
            // Reprint
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 13F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.BackColor = System.Drawing.Color.SteelBlue;
            this.ClientSize = new System.Drawing.Size(543, 388);
            this.Controls.Add(this.lbRodape);
            this.Controls.Add(this.panel1);
            this.Name = "Reprint";
            this.ShowIcon = false;
            this.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen;
            this.Text = "Reprint  V1.0.1.4";
            this.panel1.ResumeLayout(false);
            this.panel1.PerformLayout();
            ((System.ComponentModel.ISupportInitialize)(this.pictureBox1)).EndInit();
            this.groupBox1.ResumeLayout(false);
            this.groupBox1.PerformLayout();
            this.ResumeLayout(false);
            this.PerformLayout();

        }

        #endregion

        private System.Windows.Forms.Button btnClose;
        private System.Windows.Forms.Panel panel1;
        private System.Windows.Forms.Label lbTextoSerial;
        private System.Windows.Forms.TextBox txtSerial;
        private System.Windows.Forms.Label lbRodape;
        private System.Windows.Forms.GroupBox groupBox1;
        private System.Windows.Forms.PictureBox pictureBox1;
        private System.Windows.Forms.CheckBox chkSelecionar;
        private System.Windows.Forms.CheckBox chkTg3442;
        private System.Windows.Forms.Label lbWo;
        private System.Windows.Forms.TextBox txtWo;
    }
}