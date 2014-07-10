using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Windows.Forms;

namespace JH.IBSH.Diploma
{
    public partial class AddNew : FISCA.Presentation.Controls.BaseForm
    {
        public string name = "";
        public Aspose.Words.Document Template = null;
        public AddNew()
        {
            InitializeComponent();
        }
        private void buttonX1_Click(object sender, EventArgs e)
        {
            this.name = textBoxX1.Text;
            this.Close();
        }
        private void radioButton1_CheckedChanged(object sender, EventArgs e)
        {
            DevComponents.DotNetBar.Controls.CheckBoxX cb = (DevComponents.DotNetBar.Controls.CheckBoxX)sender;
            switch (cb.Name)
            {
                case "check_default"://"使用預設樣板":
                    if (cb.Checked)
                        Template = null;
                    break;
                case "check_custom"://使用自訂樣板
                    if (cb.Checked && Template == null)
                    {
                        try
                        {
                            OpenFileDialog ofd = new OpenFileDialog();
                            ofd.Filter = "Word (*.doc)|*.doc|所有檔案 (*.*)|*.*";

                            if (ofd.ShowDialog() == DialogResult.OK)
                                Template = new Aspose.Words.Document(ofd.FileName);
                            else
                                check_default.Checked = true;
                        }
                        catch
                        {
                            FISCA.Presentation.Controls.MsgBox.Show("檔案開啟錯誤,請檢查檔案是否開啟中!!");
                        }
                    }
                    break;
            }
        }

        private void linkLabel1_LinkClicked(object sender, LinkLabelLinkClickedEventArgs e)
        {
            LinkLabel ll = (LinkLabel)sender;
            Aspose.Words.Document doc = null;
            switch (ll.Text)
            {
                case "檢視樣板":
                    doc = new Campus.Report.ReportTemplate(Properties.Resources.樣板, Campus.Report.TemplateType.Word).ToDocument();
                    break;
                case "檢視合併欄位總表":
                    doc = new Campus.Report.ReportTemplate(Properties.Resources.合併欄位總表, Campus.Report.TemplateType.Word).ToDocument();
                    break;
                default:
                    return;
            }
            try
            {
                SaveFileDialog SaveFileDialog1 = new SaveFileDialog();
                SaveFileDialog1.Filter = "Word (*.doc)|*.doc|所有檔案 (*.*)|*.*";
                SaveFileDialog1.FileName = ll.Text.Replace("檢視", "");
                if (SaveFileDialog1.ShowDialog() == DialogResult.OK)
                {
                    doc.Save(SaveFileDialog1.FileName);
                    System.Diagnostics.Process.Start(SaveFileDialog1.FileName);
                }
                else
                    FISCA.Presentation.Controls.MsgBox.Show("檔案未儲存");
            }
            catch
            {
                FISCA.Presentation.Controls.MsgBox.Show("檔案儲存錯誤,請檢查檔案是否開啟中!!");
            }
        }
    }
}
