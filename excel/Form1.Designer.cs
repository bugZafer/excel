namespace excel
{
    partial class Form1
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
            this.dataGridView1 = new System.Windows.Forms.DataGridView();
            this.SiraNo = new System.Windows.Forms.DataGridViewTextBoxColumn();
            this.TezgahKodu = new System.Windows.Forms.DataGridViewTextBoxColumn();
            this.FirmaAdi = new System.Windows.Forms.DataGridViewTextBoxColumn();
            this.Stokod = new System.Windows.Forms.DataGridViewTextBoxColumn();
            this.MalzemeAd = new System.Windows.Forms.DataGridViewTextBoxColumn();
            this.IsEmriNo = new System.Windows.Forms.DataGridViewTextBoxColumn();
            this.Mevcut = new System.Windows.Forms.DataGridViewTextBoxColumn();
            this.IstenenMiktar = new System.Windows.Forms.DataGridViewTextBoxColumn();
            this.IstenenTarih = new System.Windows.Forms.DataGridViewTextBoxColumn();
            this.KullanilanMalzeme = new System.Windows.Forms.DataGridViewTextBoxColumn();
            this.button1 = new System.Windows.Forms.Button();
            this.button2 = new System.Windows.Forms.Button();
            this.dataGridView2 = new System.Windows.Forms.DataGridView();
            this.dataGridView3 = new System.Windows.Forms.DataGridView();
            this.button3 = new System.Windows.Forms.Button();
            this.progressBar = new System.Windows.Forms.ProgressBar();
            this.button4 = new System.Windows.Forms.Button();
            this.button5 = new System.Windows.Forms.Button();
            ((System.ComponentModel.ISupportInitialize)(this.dataGridView1)).BeginInit();
            ((System.ComponentModel.ISupportInitialize)(this.dataGridView2)).BeginInit();
            ((System.ComponentModel.ISupportInitialize)(this.dataGridView3)).BeginInit();
            this.SuspendLayout();
            // 
            // dataGridView1
            // 
            this.dataGridView1.AllowUserToAddRows = false;
            this.dataGridView1.Anchor = ((System.Windows.Forms.AnchorStyles)((((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Bottom) 
            | System.Windows.Forms.AnchorStyles.Left) 
            | System.Windows.Forms.AnchorStyles.Right)));
            this.dataGridView1.AutoSizeColumnsMode = System.Windows.Forms.DataGridViewAutoSizeColumnsMode.Fill;
            this.dataGridView1.ColumnHeadersHeightSizeMode = System.Windows.Forms.DataGridViewColumnHeadersHeightSizeMode.AutoSize;
            this.dataGridView1.Columns.AddRange(new System.Windows.Forms.DataGridViewColumn[] {
            this.SiraNo,
            this.TezgahKodu,
            this.FirmaAdi,
            this.Stokod,
            this.MalzemeAd,
            this.IsEmriNo,
            this.Mevcut,
            this.IstenenMiktar,
            this.IstenenTarih,
            this.KullanilanMalzeme});
            this.dataGridView1.Location = new System.Drawing.Point(0, 0);
            this.dataGridView1.Name = "dataGridView1";
            this.dataGridView1.ReadOnly = true;
            this.dataGridView1.Size = new System.Drawing.Size(844, 599);
            this.dataGridView1.TabIndex = 0;
            // 
            // SiraNo
            // 
            this.SiraNo.HeaderText = "Sıra No";
            this.SiraNo.Name = "SiraNo";
            this.SiraNo.ReadOnly = true;
            // 
            // TezgahKodu
            // 
            this.TezgahKodu.HeaderText = "Tezgah Kodu";
            this.TezgahKodu.Name = "TezgahKodu";
            this.TezgahKodu.ReadOnly = true;
            // 
            // FirmaAdi
            // 
            this.FirmaAdi.HeaderText = "Firma Adı";
            this.FirmaAdi.Name = "FirmaAdi";
            this.FirmaAdi.ReadOnly = true;
            // 
            // Stokod
            // 
            this.Stokod.HeaderText = "Stok Kod";
            this.Stokod.Name = "Stokod";
            this.Stokod.ReadOnly = true;
            // 
            // MalzemeAd
            // 
            this.MalzemeAd.HeaderText = "Malzeme Ad";
            this.MalzemeAd.Name = "MalzemeAd";
            this.MalzemeAd.ReadOnly = true;
            // 
            // IsEmriNo
            // 
            this.IsEmriNo.HeaderText = "İş Emri No";
            this.IsEmriNo.Name = "IsEmriNo";
            this.IsEmriNo.ReadOnly = true;
            // 
            // Mevcut
            // 
            this.Mevcut.HeaderText = "Mevcut";
            this.Mevcut.Name = "Mevcut";
            this.Mevcut.ReadOnly = true;
            // 
            // IstenenMiktar
            // 
            this.IstenenMiktar.HeaderText = "İstenen Miktar";
            this.IstenenMiktar.Name = "IstenenMiktar";
            this.IstenenMiktar.ReadOnly = true;
            // 
            // IstenenTarih
            // 
            this.IstenenTarih.HeaderText = "İstenen Tar.";
            this.IstenenTarih.Name = "IstenenTarih";
            this.IstenenTarih.ReadOnly = true;
            // 
            // KullanilanMalzeme
            // 
            this.KullanilanMalzeme.HeaderText = "Kullanılan Malz.";
            this.KullanilanMalzeme.Name = "KullanilanMalzeme";
            this.KullanilanMalzeme.ReadOnly = true;
            // 
            // button1
            // 
            this.button1.Location = new System.Drawing.Point(875, 12);
            this.button1.Name = "button1";
            this.button1.RightToLeft = System.Windows.Forms.RightToLeft.No;
            this.button1.Size = new System.Drawing.Size(112, 23);
            this.button1.TabIndex = 1;
            this.button1.Text = "Siparişleri Seç";
            this.button1.UseVisualStyleBackColor = true;
            this.button1.Click += new System.EventHandler(this.button1_Click);
            // 
            // button2
            // 
            this.button2.Location = new System.Drawing.Point(875, 70);
            this.button2.Name = "button2";
            this.button2.Size = new System.Drawing.Size(112, 23);
            this.button2.TabIndex = 4;
            this.button2.Text = "Tezgah Bilgisi Seç";
            this.button2.UseVisualStyleBackColor = true;
            this.button2.Click += new System.EventHandler(this.button2_Click);
            // 
            // dataGridView2
            // 
            this.dataGridView2.ColumnHeadersHeightSizeMode = System.Windows.Forms.DataGridViewColumnHeadersHeightSizeMode.AutoSize;
            this.dataGridView2.Location = new System.Drawing.Point(23, 38);
            this.dataGridView2.Name = "dataGridView2";
            this.dataGridView2.Size = new System.Drawing.Size(150, 39);
            this.dataGridView2.TabIndex = 6;
            this.dataGridView2.Visible = false;
            // 
            // dataGridView3
            // 
            this.dataGridView3.CausesValidation = false;
            this.dataGridView3.ColumnHeadersHeightSizeMode = System.Windows.Forms.DataGridViewColumnHeadersHeightSizeMode.AutoSize;
            this.dataGridView3.Location = new System.Drawing.Point(233, 12);
            this.dataGridView3.Name = "dataGridView3";
            this.dataGridView3.Size = new System.Drawing.Size(150, 23);
            this.dataGridView3.TabIndex = 7;
            this.dataGridView3.Visible = false;
            // 
            // button3
            // 
            this.button3.Location = new System.Drawing.Point(875, 128);
            this.button3.Name = "button3";
            this.button3.Size = new System.Drawing.Size(112, 23);
            this.button3.TabIndex = 8;
            this.button3.Text = "Excele Aktar";
            this.button3.UseVisualStyleBackColor = true;
            this.button3.Click += new System.EventHandler(this.button3_Click);
            // 
            // progressBar
            // 
            this.progressBar.Location = new System.Drawing.Point(850, 299);
            this.progressBar.Name = "progressBar";
            this.progressBar.Size = new System.Drawing.Size(156, 23);
            this.progressBar.TabIndex = 9;
            // 
            // button4
            // 
            this.button4.Location = new System.Drawing.Point(875, 41);
            this.button4.Name = "button4";
            this.button4.Size = new System.Drawing.Size(112, 23);
            this.button4.TabIndex = 10;
            this.button4.Text = "İş Emri Seç";
            this.button4.UseVisualStyleBackColor = true;
            this.button4.Click += new System.EventHandler(this.button4_Click);
            // 
            // button5
            // 
            this.button5.Location = new System.Drawing.Point(875, 99);
            this.button5.Name = "button5";
            this.button5.Size = new System.Drawing.Size(112, 23);
            this.button5.TabIndex = 11;
            this.button5.Text = "Stok Aktar";
            this.button5.UseVisualStyleBackColor = true;
            this.button5.Click += new System.EventHandler(this.button5_Click);
            // 
            // Form1
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 13F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.ClientSize = new System.Drawing.Size(1005, 599);
            this.Controls.Add(this.button5);
            this.Controls.Add(this.button4);
            this.Controls.Add(this.progressBar);
            this.Controls.Add(this.button3);
            this.Controls.Add(this.button2);
            this.Controls.Add(this.button1);
            this.Controls.Add(this.dataGridView1);
            this.Controls.Add(this.dataGridView2);
            this.Controls.Add(this.dataGridView3);
            this.MinimumSize = new System.Drawing.Size(16, 39);
            this.Name = "Form1";
            this.Text = "Form1";
            this.WindowState = System.Windows.Forms.FormWindowState.Maximized;
            this.FormClosing += new System.Windows.Forms.FormClosingEventHandler(this.Form1_FormClosing);
            this.FormClosed += new System.Windows.Forms.FormClosedEventHandler(this.Form1_FormClosed_1);
            this.Load += new System.EventHandler(this.Form1_Load);
            this.KeyDown += new System.Windows.Forms.KeyEventHandler(this.Form1_KeyDown);
            ((System.ComponentModel.ISupportInitialize)(this.dataGridView1)).EndInit();
            ((System.ComponentModel.ISupportInitialize)(this.dataGridView2)).EndInit();
            ((System.ComponentModel.ISupportInitialize)(this.dataGridView3)).EndInit();
            this.ResumeLayout(false);

        }

        #endregion

        private System.Windows.Forms.DataGridView dataGridView1;
        private System.Windows.Forms.Button button1;
        private System.Windows.Forms.Button button2;
        private System.Windows.Forms.DataGridView dataGridView2;
        private System.Windows.Forms.DataGridView dataGridView3;
        private System.Windows.Forms.Button button3;
        private System.Windows.Forms.ProgressBar progressBar;
        private System.Windows.Forms.Button button4;
        private System.Windows.Forms.DataGridViewTextBoxColumn SiraNo;
        private System.Windows.Forms.DataGridViewTextBoxColumn TezgahKodu;
        private System.Windows.Forms.DataGridViewTextBoxColumn FirmaAdi;
        private System.Windows.Forms.DataGridViewTextBoxColumn Stokod;
        private System.Windows.Forms.DataGridViewTextBoxColumn MalzemeAd;
        private System.Windows.Forms.DataGridViewTextBoxColumn IsEmriNo;
        private System.Windows.Forms.DataGridViewTextBoxColumn Mevcut;
        private System.Windows.Forms.DataGridViewTextBoxColumn IstenenMiktar;
        private System.Windows.Forms.DataGridViewTextBoxColumn IstenenTarih;
        private System.Windows.Forms.DataGridViewTextBoxColumn KullanilanMalzeme;
        private System.Windows.Forms.Button button5;
    }
}

