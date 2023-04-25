namespace ExcelTemplatesLib
{
    partial class CFGForm
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
            System.ComponentModel.ComponentResourceManager resources = new System.ComponentModel.ComponentResourceManager(typeof(CFGForm));
            this.listLabel = new System.Windows.Forms.Label();
            this.xlsxView = new System.Windows.Forms.ListView();
            this.columnHeader1 = ((System.Windows.Forms.ColumnHeader)(new System.Windows.Forms.ColumnHeader()));
            this.columnHeader2 = ((System.Windows.Forms.ColumnHeader)(new System.Windows.Forms.ColumnHeader()));
            this.columnHeader3 = ((System.Windows.Forms.ColumnHeader)(new System.Windows.Forms.ColumnHeader()));
            this.upBtn = new System.Windows.Forms.Button();
            this.label1 = new System.Windows.Forms.Label();
            this.selStartBox = new System.Windows.Forms.ComboBox();
            this.label2 = new System.Windows.Forms.Label();
            this.info = new System.Windows.Forms.Label();
            this.linkLabel1 = new System.Windows.Forms.LinkLabel();
            this.qrip = new System.Windows.Forms.CheckBox();
            this.matrixBar = new System.Windows.Forms.CheckBox();
            this.code39 = new System.Windows.Forms.CheckBox();
            this.comboBox1 = new System.Windows.Forms.ComboBox();
            this.comboBox2 = new System.Windows.Forms.ComboBox();
            this.SuspendLayout();
            // 
            // listLabel
            // 
            this.listLabel.AutoSize = true;
            this.listLabel.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(204)));
            this.listLabel.Location = new System.Drawing.Point(16, 11);
            this.listLabel.Margin = new System.Windows.Forms.Padding(4, 0, 4, 0);
            this.listLabel.Name = "listLabel";
            this.listLabel.Size = new System.Drawing.Size(373, 17);
            this.listLabel.TabIndex = 0;
            this.listLabel.Text = "Список шаблонов (двойной клик для редактирования):";
            // 
            // xlsxView
            // 
            this.xlsxView.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle;
            this.xlsxView.Columns.AddRange(new System.Windows.Forms.ColumnHeader[] {
            this.columnHeader1,
            this.columnHeader2,
            this.columnHeader3});
            this.xlsxView.FullRowSelect = true;
            this.xlsxView.HideSelection = false;
            this.xlsxView.Location = new System.Drawing.Point(20, 39);
            this.xlsxView.Margin = new System.Windows.Forms.Padding(4);
            this.xlsxView.MultiSelect = false;
            this.xlsxView.Name = "xlsxView";
            this.xlsxView.Size = new System.Drawing.Size(557, 497);
            this.xlsxView.Sorting = System.Windows.Forms.SortOrder.Ascending;
            this.xlsxView.TabIndex = 1;
            this.xlsxView.UseCompatibleStateImageBehavior = false;
            this.xlsxView.View = System.Windows.Forms.View.Details;
            this.xlsxView.DoubleClick += new System.EventHandler(this.xlsxView_DoubleClick);
            // 
            // columnHeader1
            // 
            this.columnHeader1.Text = "Файл";
            this.columnHeader1.Width = 190;
            // 
            // columnHeader2
            // 
            this.columnHeader2.Text = "Размер";
            this.columnHeader2.Width = 75;
            // 
            // columnHeader3
            // 
            this.columnHeader3.Text = "Изменен";
            this.columnHeader3.Width = 120;
            // 
            // upBtn
            // 
            this.upBtn.Location = new System.Drawing.Point(477, 5);
            this.upBtn.Margin = new System.Windows.Forms.Padding(4);
            this.upBtn.Name = "upBtn";
            this.upBtn.Size = new System.Drawing.Size(100, 28);
            this.upBtn.TabIndex = 2;
            this.upBtn.Text = "Обновить";
            this.upBtn.UseVisualStyleBackColor = true;
            this.upBtn.Click += new System.EventHandler(this.upBtn_Click);
            // 
            // label1
            // 
            this.label1.AutoSize = true;
            this.label1.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(204)));
            this.label1.Location = new System.Drawing.Point(16, 548);
            this.label1.Margin = new System.Windows.Forms.Padding(4, 0, 4, 0);
            this.label1.Name = "label1";
            this.label1.Size = new System.Drawing.Size(202, 17);
            this.label1.TabIndex = 3;
            this.label1.Text = "Выбор шаблона при запуске:";
            // 
            // selStartBox
            // 
            this.selStartBox.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList;
            this.selStartBox.FormattingEnabled = true;
            this.selStartBox.Items.AddRange(new object[] {
            "Выбирать актуальный шаблон (ИНН, общий)",
            "Показывать список шаблонов",
            "Выбирать общий шаблон",
            "Выбирать шаблон по ИНН",
            "Последний выбранный шаблон из списка"});
            this.selStartBox.Location = new System.Drawing.Point(231, 544);
            this.selStartBox.Margin = new System.Windows.Forms.Padding(4);
            this.selStartBox.Name = "selStartBox";
            this.selStartBox.Size = new System.Drawing.Size(345, 24);
            this.selStartBox.TabIndex = 4;
            this.selStartBox.SelectedIndexChanged += new System.EventHandler(this.selStartBox_SelectedIndexChanged);
            // 
            // label2
            // 
            this.label2.AutoSize = true;
            this.label2.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(204)));
            this.label2.Location = new System.Drawing.Point(585, 71);
            this.label2.Margin = new System.Windows.Forms.Padding(4, 0, 4, 0);
            this.label2.Name = "label2";
            this.label2.Size = new System.Drawing.Size(106, 17);
            this.label2.TabIndex = 5;
            this.label2.Text = "Примечание:";
            // 
            // info
            // 
            this.info.AutoSize = true;
            this.info.Location = new System.Drawing.Point(585, 101);
            this.info.Margin = new System.Windows.Forms.Padding(4, 0, 4, 0);
            this.info.Name = "info";
            this.info.Size = new System.Drawing.Size(523, 432);
            this.info.TabIndex = 6;
            this.info.Text = resources.GetString("info.Text");
            // 
            // linkLabel1
            // 
            this.linkLabel1.AutoSize = true;
            this.linkLabel1.Location = new System.Drawing.Point(808, 548);
            this.linkLabel1.Margin = new System.Windows.Forms.Padding(4, 0, 4, 0);
            this.linkLabel1.Name = "linkLabel1";
            this.linkLabel1.Size = new System.Drawing.Size(301, 16);
            this.linkLabel1.TabIndex = 7;
            this.linkLabel1.TabStop = true;
            this.linkLabel1.Text = "https://github.com/dkxce/IPUSN2ExcelTemplates";
            this.linkLabel1.LinkClicked += new System.Windows.Forms.LinkLabelLinkClickedEventHandler(this.linkLabel1_LinkClicked);
            // 
            // qrip
            // 
            this.qrip.AutoSize = true;
            this.qrip.Location = new System.Drawing.Point(589, 5);
            this.qrip.Margin = new System.Windows.Forms.Padding(4);
            this.qrip.Name = "qrip";
            this.qrip.Size = new System.Drawing.Size(305, 20);
            this.qrip.TabIndex = 8;
            this.qrip.Text = "Допускать печать QR-кода оплаты для ИП";
            this.qrip.UseVisualStyleBackColor = true;
            this.qrip.CheckedChanged += new System.EventHandler(this.qrip_CheckedChanged);
            // 
            // matrixBar
            // 
            this.matrixBar.AutoSize = true;
            this.matrixBar.Location = new System.Drawing.Point(589, 43);
            this.matrixBar.Margin = new System.Windows.Forms.Padding(4);
            this.matrixBar.Name = "matrixBar";
            this.matrixBar.Size = new System.Drawing.Size(243, 20);
            this.matrixBar.TabIndex = 9;
            this.matrixBar.Text = "Печатать двухмерный штрих-код";
            this.matrixBar.UseVisualStyleBackColor = true;
            this.matrixBar.CheckedChanged += new System.EventHandler(this.matrixBar_CheckedChanged);
            // 
            // code39
            // 
            this.code39.AutoSize = true;
            this.code39.Location = new System.Drawing.Point(589, 24);
            this.code39.Margin = new System.Windows.Forms.Padding(4);
            this.code39.Name = "code39";
            this.code39.Size = new System.Drawing.Size(245, 20);
            this.code39.TabIndex = 10;
            this.code39.Text = "Печатать одномерный штрих-код";
            this.code39.UseVisualStyleBackColor = true;
            this.code39.CheckedChanged += new System.EventHandler(this.code39_CheckedChanged);
            // 
            // comboBox1
            // 
            this.comboBox1.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList;
            this.comboBox1.FormattingEnabled = true;
            this.comboBox1.Items.AddRange(new object[] {
            "AzTec",
            "Codablock-F",
            "Code 16K",
            "Code 49",
            "Code One",
            "Data Matrix",
            "Dot Code",
            "Grid Matrix",
            "Han Xin\t",
            "MaxiCode",
            "MicroPDF417",
            "PDF471",
            "QRCode",
            "rMQR",
            "UltraCode",
            "UPNQR"});
            this.comboBox1.Location = new System.Drawing.Point(914, 46);
            this.comboBox1.Name = "comboBox1";
            this.comboBox1.Size = new System.Drawing.Size(238, 24);
            this.comboBox1.TabIndex = 11;
            this.comboBox1.SelectedIndexChanged += new System.EventHandler(this.comboBox1_SelectedIndexChanged);
            // 
            // comboBox2
            // 
            this.comboBox2.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList;
            this.comboBox2.FormattingEnabled = true;
            this.comboBox2.Items.AddRange(new object[] {
            "Code 128",
            "Code 39",
            "Code 93"});
            this.comboBox2.Location = new System.Drawing.Point(914, 20);
            this.comboBox2.Name = "comboBox2";
            this.comboBox2.Size = new System.Drawing.Size(238, 24);
            this.comboBox2.TabIndex = 12;
            this.comboBox2.SelectedIndexChanged += new System.EventHandler(this.comboBox2_SelectedIndexChanged);
            // 
            // CFGForm
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(8F, 16F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.ClientSize = new System.Drawing.Size(1164, 583);
            this.Controls.Add(this.comboBox2);
            this.Controls.Add(this.comboBox1);
            this.Controls.Add(this.code39);
            this.Controls.Add(this.matrixBar);
            this.Controls.Add(this.qrip);
            this.Controls.Add(this.linkLabel1);
            this.Controls.Add(this.info);
            this.Controls.Add(this.label2);
            this.Controls.Add(this.selStartBox);
            this.Controls.Add(this.label1);
            this.Controls.Add(this.upBtn);
            this.Controls.Add(this.xlsxView);
            this.Controls.Add(this.listLabel);
            this.FormBorderStyle = System.Windows.Forms.FormBorderStyle.FixedSingle;
            this.Icon = ((System.Drawing.Icon)(resources.GetObject("$this.Icon")));
            this.Margin = new System.Windows.Forms.Padding(4);
            this.MaximizeBox = false;
            this.Name = "CFGForm";
            this.Text = "Настройки плагина ExcelTemplates для ИП УСН2";
            this.FormClosed += new System.Windows.Forms.FormClosedEventHandler(this.CFGForm_FormClosed);
            this.Load += new System.EventHandler(this.CFGForm_Load);
            this.ResumeLayout(false);
            this.PerformLayout();

        }

        #endregion

        private System.Windows.Forms.Label listLabel;
        private System.Windows.Forms.ListView xlsxView;
        private System.Windows.Forms.ColumnHeader columnHeader1;
        private System.Windows.Forms.ColumnHeader columnHeader2;
        private System.Windows.Forms.ColumnHeader columnHeader3;
        private System.Windows.Forms.Button upBtn;
        private System.Windows.Forms.Label label1;
        private System.Windows.Forms.ComboBox selStartBox;
        private System.Windows.Forms.Label label2;
        private System.Windows.Forms.Label info;
        private System.Windows.Forms.LinkLabel linkLabel1;
        private System.Windows.Forms.CheckBox qrip;
        private System.Windows.Forms.CheckBox matrixBar;
        private System.Windows.Forms.CheckBox code39;
        private System.Windows.Forms.ComboBox comboBox1;
        private System.Windows.Forms.ComboBox comboBox2;
    }
}