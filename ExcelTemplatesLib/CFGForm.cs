using System;
using System.Collections;
using System.Collections.Generic;
using System.IO;
using System.Windows.Forms;
using System.Xml;

namespace ExcelTemplatesLib
{
    public partial class CFGForm : Form
    {
        public bool IsAlive = true;
        
        private string CurrDir = XMLSaved<int>.CurrentDirectory();
        private string xlsxDir = Path.Combine(XMLSaved<int>.CurrentDirectory(), "Templates");        
        private PluginConfig pluginConfig = new PluginConfig();

        public CFGForm()
        {
            InitializeComponent();
        }

        private void CFGForm_FormClosed(object sender, FormClosedEventArgs e)
        {
            SaveCfg();
            IsAlive = false;            
        }

        private void CFGForm_Load(object sender, EventArgs e)
        {
            info.Text = $"Сохранять шаблоны необходимо в папку:\r\n{xlsxDir}\r\n\r\n{info.Text}";
            LoadCfg();
            Reload();
        }

        private void upBtn_Click(object sender, EventArgs e)
        {
            Reload();
        }

        private void LoadCfg()
        {
            if (File.Exists(PluginConfig.FileName)) pluginConfig = PluginConfig.Load(PluginConfig.FileName);
            selStartBox.SelectedIndex = pluginConfig.StartMode;
            qrip.Checked = pluginConfig.QRIP;
            matrixBar.Checked = pluginConfig.MatrixBar;
            code39.Checked = pluginConfig.Code39Bar;
        }

        private void SaveCfg()
        {
            PluginConfig.Save(PluginConfig.FileName, pluginConfig);
        }

        private void Reload()
        {
            xlsxView.Items.Clear();
            List<string> fls = new List<string>();
            fls.AddRange(Directory.GetFiles(xlsxDir, "*.xlsx", SearchOption.TopDirectoryOnly));
            fls.AddRange(Directory.GetFiles(xlsxDir, "*.xlsm", SearchOption.TopDirectoryOnly));
            foreach ( string f in fls ) 
            {
                FileInfo i = new FileInfo(f);

                string n = Path.GetFileName(f);
                ListViewItem lvi = new ListViewItem(n);
                lvi.SubItems.Add(BytesToString(i.Length));
                lvi.SubItems.Add($"{i.LastWriteTime}");
                lvi.SubItems.Add(f);
                xlsxView.Items.Add(lvi);
            }
        }

        public static string BytesToString(long size)
        {
            string[] sizes = { "байт", "КБ", "МБ", "ГБ", "ТБ" };
            double len = size;
            int order = 0;
            while (len >= 1024 && order < sizes.Length - 1)
            {
                order++;
                len = len / 1024;
            };
            return String.Format(System.Globalization.CultureInfo.InvariantCulture, "{0:0.##} {1}", len, sizes[order]);
        }

        private void xlsxView_DoubleClick(object sender, EventArgs e)
        {
            if (xlsxView.SelectedItems.Count == 0) return;
            try { System.Diagnostics.Process.Start(xlsxView.SelectedItems[0].SubItems[3].Text); } catch { };
        }

        private void selStartBox_SelectedIndexChanged(object sender, EventArgs e)
        {
            pluginConfig.StartMode = (byte)selStartBox.SelectedIndex;
            SaveCfg();
        }

        private void linkLabel1_LinkClicked(object sender, LinkLabelLinkClickedEventArgs e)
        {
            try { System.Diagnostics.Process.Start("https://github.com/dkxce/IPUSN2ExcelTemplates"); } catch { };
        }

        private void qrip_CheckedChanged(object sender, EventArgs e)
        {
            pluginConfig.QRIP = qrip.Checked;
            SaveCfg();
        }

        private void matrixBar_CheckedChanged(object sender, EventArgs e)
        {
            pluginConfig.MatrixBar = matrixBar.Checked;
            SaveCfg();
        }

        private void code39_CheckedChanged(object sender, EventArgs e)
        {
            pluginConfig.Code39Bar = code39.Checked;
            SaveCfg();
        }
    }
}
