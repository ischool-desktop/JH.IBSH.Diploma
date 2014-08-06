using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Diagnostics;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Windows.Forms;
using System.Xml;

using Aspose.Words;
using Aspose.Words.Drawing;
using Campus.Report;
using JHSchool.Data;
using K12.Data;
namespace JH.IBSH.Diploma
{
    public partial class MainForm : FISCA.Presentation.Controls.BaseForm
    {
        private BackgroundWorker _bgw = new BackgroundWorker();

        private const string config = "plugins.student.report.certificate.huangwc.config";
        private Dictionary<string, ReportConfiguration> custConfigs = new Dictionary<string, ReportConfiguration>();
        ReportConfiguration conf = new Campus.Report.ReportConfiguration(config);
        public string current = "";

        public MainForm()
        {
            InitializeComponent();

            #region 設定comboBox選單
            foreach (string item in getCustConfig())
            {
                if (!string.IsNullOrWhiteSpace(item))
                {
                    custConfigs.Add(item, new ReportConfiguration(configNameRule(item)));
                    comboBoxEx1.Items.Add(item);
                }
            }
            comboBoxEx1.Items.Add("新增");
            comboBoxEx1.SelectedIndex = 0;
            #endregion
            _bgw.DoWork += new DoWorkEventHandler(_bgw_DoWork);
            _bgw.RunWorkerCompleted += new RunWorkerCompletedEventHandler(_bgw_RunWorkerCompleted);
        }
        private void linkLabel1_LinkClicked(object sender, LinkLabelLinkClickedEventArgs e)
        {
            string value = (string)comboBoxEx1.SelectedItem;
            if (value == "新增") return;
            //畫面內容(範本內容,預設樣式
            Campus.Report.TemplateSettingForm TemplateForm;
            if (custConfigs[current].Template == null)
            {
                custConfigs[current].Template = new Campus.Report.ReportTemplate(Properties.Resources.樣板, Campus.Report.TemplateType.Word);
            }
            TemplateForm = new Campus.Report.TemplateSettingForm(custConfigs[current].Template, new Campus.Report.ReportTemplate(Properties.Resources.樣板, Campus.Report.TemplateType.Word));
            //預設名稱
            TemplateForm.DefaultFileName = current + "樣板";
            if (TemplateForm.ShowDialog() == DialogResult.OK)
            {
                custConfigs[current].Template = TemplateForm.Template;
                custConfigs[current].Save();
            }
        }
        private void btnPrint_Click(object sender, EventArgs e)
        {
            string value = (string)comboBoxEx1.SelectedItem;
            if (value == "新增") return;
            if (K12.Presentation.NLDPanels.Student.SelectedSource.Count < 1)
            {
                FISCA.Presentation.Controls.MsgBox.Show("請先選擇學生");
                return;
            }
            btnPrint.Enabled = false;
            _bgw.RunWorkerAsync();
        }
        void _bgw_DoWork(object sender, DoWorkEventArgs e)
        {
            Document document = new Document();
            Document template = (custConfigs[current].Template != null) //單頁範本
                 ? custConfigs[current].Template.ToDocument()
                 : new Campus.Report.ReportTemplate(Properties.Resources.樣板, Campus.Report.TemplateType.Word).ToDocument();
            List<string> student_ids = K12.Presentation.NLDPanels.Student.SelectedSource;

            List<JHStudentRecord> jhsrl = JHStudent.SelectByIDs(student_ids);
            List<JHUpdateRecordRecord> jhurrl = JHUpdateRecord.SelectByStudentIDs(student_ids);
            Dictionary<string, JHUpdateRecordRecord> djhurr = new Dictionary<string, JHUpdateRecordRecord>();
            foreach (JHUpdateRecordRecord jhurr in jhurrl)
            {
                if (jhurr.UpdateCode == "2")
                {
                    if (djhurr.ContainsKey(jhurr.StudentID))
                    {
                        if (djhurr[jhurr.StudentID].UpdateDate.CompareTo(jhurr.UpdateDate) == 1)
                        {
                            djhurr[jhurr.StudentID] = jhurr;
                        }
                    }
                    else djhurr.Add(jhurr.StudentID, jhurr);
                }
            }
            //入學照片
            Dictionary<string, string> dic_photo_p = K12.Data.Photo.SelectFreshmanPhoto(K12.Presentation.NLDPanels.Student.SelectedSource);
            Dictionary<string, string> dic_photo_g = K12.Data.Photo.SelectGraduatePhoto(K12.Presentation.NLDPanels.Student.SelectedSource);

            Dictionary<string, object> mailmerge = new Dictionary<string, object>();
            string 校內字號 = textBoxX1.Text;
            string 校內字號英文 = textBoxX2.Text;
            foreach (JHStudentRecord jhsr in jhsrl)
            {
                mailmerge.Clear();
                #region MailMerge

                #region 學生資料
                mailmerge.Add("學生姓名", jhsr.Name);
                mailmerge.Add("學生英文姓名", jhsr.EnglishName);
                mailmerge.Add("學生身分證號", jhsr.IDNumber);
                mailmerge.Add("學生目前班級", jhsr.Class.Name);
                mailmerge.Add("學生目前年級", jhsr.Class.GradeYear);
                mailmerge.Add("學生目前座號", jhsr.StudentNumber);
                if (jhsr.Birthday.HasValue)
                {
                    mailmerge.Add("學生生日民國年", jhsr.Birthday.Value.Year - 1911);
                    mailmerge.Add("學生生日英文年", jhsr.Birthday.Value.Year);
                    mailmerge.Add("學生生日中文年", toZhNumber("" + (jhsr.Birthday.Value.Year - 1911)));
                    mailmerge.Add("學生生日月", jhsr.Birthday.Value.Month);
                    mailmerge.Add("學生生日中文月", toZhNumber("" + jhsr.Birthday.Value.Month));
                    mailmerge.Add("學生生日英文月", jhsr.Birthday.Value.ToString("MMMM", new System.Globalization.CultureInfo("en-US")));
                    mailmerge.Add("學生生日英文月3", jhsr.Birthday.Value.ToString("MMM", new System.Globalization.CultureInfo("en-US")));
                    mailmerge.Add("學生生日上標", daySuffix(jhsr.Birthday.Value.Day.ToString()));
                    mailmerge.Add("學生生日日", jhsr.Birthday.Value.Day);
                    mailmerge.Add("學生生日中文日", toZhNumber("" + jhsr.Birthday.Value.Day));
                    mailmerge.Add("學生生日日序數", dayToOrdinal(jhsr.Birthday.Value.Day));
                }
                else
                {
                    mailmerge.Add("學生生日民國年", "");
                    mailmerge.Add("學生生日英文年", "");
                    mailmerge.Add("學生生日中文年", "");
                    mailmerge.Add("學生生日月", "");
                    mailmerge.Add("學生生日中文月", "");
                    mailmerge.Add("學生生日英文月", "");
                    mailmerge.Add("學生生日英文月3", "");
                    mailmerge.Add("學生生日上標", "");
                    mailmerge.Add("學生生日日", "");
                    mailmerge.Add("學生生日中文日", "");
                    mailmerge.Add("學生生日日序數", "");
                }

                if (dic_photo_p.ContainsKey(jhsr.ID))
                {
                    mailmerge.Add("入學照片1吋", dic_photo_p[jhsr.ID]);
                    mailmerge.Add("入學照片2吋", dic_photo_p[jhsr.ID]);
                }
                if (dic_photo_p.ContainsKey(jhsr.ID))
                {
                    mailmerge.Add("畢業照片1吋", dic_photo_g[jhsr.ID]);
                    mailmerge.Add("畢業照片2吋", dic_photo_g[jhsr.ID]);
                }
                foreach (string tmp in new string[] { "離校學年度", "畢業證書字號", "離校科別中文", "離校科別英文" })
                    mailmerge.Add(tmp, "");
                if (djhurr.ContainsKey(jhsr.ID))//畢業異動
                {
                    mailmerge["離校學年度"] = djhurr[jhsr.ID].SchoolYear;
                    mailmerge["畢業證書字號"] = djhurr[jhsr.ID].GraduateCertificateNumber;
                }
                #endregion
                mailmerge.Add("民國年", DateTime.Today.Year - 1911);
                mailmerge.Add("英文年", DateTime.Today.Year);
                mailmerge.Add("中文年", toZhNumber("" + (DateTime.Today.Year-1911)));
                mailmerge.Add("月", DateTime.Today.Month);
                mailmerge.Add("中文月", toZhNumber("" + DateTime.Today.Month));
                mailmerge.Add("英文月", DateTime.Today.ToString("MMMM", new System.Globalization.CultureInfo("en-US")));
                mailmerge.Add("英文月3", DateTime.Today.ToString("MMM", new System.Globalization.CultureInfo("en-US")));
                mailmerge.Add("日上標", daySuffix(DateTime.Today.Day.ToString()));
                mailmerge.Add("日", DateTime.Today.Day);
                mailmerge.Add("中文日", toZhNumber("" + DateTime.Today.Day));
                mailmerge.Add("日序數", dayToOrdinal(DateTime.Today.Day));
                #region 學校資料
                mailmerge.Add("學校全銜", School.ChineseName);
                mailmerge.Add("學校英文全銜", School.EnglishName);
                mailmerge.Add("校長姓名", "");
                mailmerge.Add("目前學期", School.DefaultSemester);
                mailmerge.Add("目前學年度", School.DefaultSchoolYear);

                if (K12.Data.School.Configuration["學校資訊"] != null && K12.Data.School.Configuration["學校資訊"].PreviousData.SelectSingleNode("ChancellorChineseName") != null)
                    mailmerge["校長姓名"] = K12.Data.School.Configuration["學校資訊"].PreviousData.SelectSingleNode("ChancellorChineseName").InnerText;
                mailmerge.Add("校長姓名英文", "");
                if (K12.Data.School.Configuration["學校資訊"] != null && K12.Data.School.Configuration["學校資訊"].PreviousData.SelectSingleNode("ChancellorChineseName") != null)
                    mailmerge["校長姓名英文"] = K12.Data.School.Configuration["學校資訊"].PreviousData.SelectSingleNode("ChancellorEnglishName").InnerText;
                #endregion
                mailmerge.Add("校內字號", 校內字號);
                mailmerge.Add("校內字號英文", 校內字號英文);
                #endregion
                Document each = (Document)template.Clone(true);
                each.MailMerge.MergeField += new Aspose.Words.Reporting.MergeFieldEventHandler(merge);
                each.MailMerge.Execute(mailmerge.Keys.ToArray(), mailmerge.Values.ToArray());

                document.Sections.Add(document.ImportNode(each.FirstSection, true));
            }
            document.Sections.RemoveAt(0);
            e.Result = document;
        }
        void merge(object sender, Aspose.Words.Reporting.MergeFieldEventArgs e)
        {
            if (e.FieldName == "入學照片1吋" || e.FieldName == "入學照片2吋")
            {
                int tmp_width;
                int tmp_height;
                if (e.FieldName == "入學照片1吋")
                {
                    tmp_width = 25;
                    tmp_height = 35;
                }
                else
                {
                    tmp_width = 35;
                    tmp_height = 45;
                }
                #region 入學照片
                if (!string.IsNullOrEmpty(e.FieldValue.ToString()))
                {
                    byte[] photo = Convert.FromBase64String(e.FieldValue.ToString()); //e.FieldValue as byte[];

                    if (photo != null && photo.Length > 0)
                    {
                        DocumentBuilder photoBuilder = new DocumentBuilder(e.Document);
                        photoBuilder.MoveToField(e.Field, true);
                        e.Field.Remove();
                        //Paragraph paragraph = photoBuilder.InsertParagraph();// new Paragraph(e.Document);
                        Shape photoShape = new Shape(e.Document, ShapeType.Image);
                        photoShape.ImageData.SetImage(photo);
                        photoShape.WrapType = WrapType.Inline;
                        //Cell cell = photoBuilder.CurrentParagraph.ParentNode as Cell;
                        //cell.CellFormat.LeftPadding = 0;
                        //cell.CellFormat.RightPadding = 0;

                        photoShape.Width = ConvertUtil.MillimeterToPoint(tmp_width);
                        photoShape.Height = ConvertUtil.MillimeterToPoint(tmp_height);
                        //paragraph.AppendChild(photoShape);
                        photoBuilder.InsertNode(photoShape);
                    }
                }
                #endregion
            }
            else if (e.FieldName == "畢業照片1吋" || e.FieldName == "畢業照片2吋")
            {
                int tmp_width;
                int tmp_height;
                if (e.FieldName == "畢業照片1吋")
                {
                    tmp_width = 25;
                    tmp_height = 35;
                }
                else
                {
                    tmp_width = 35;
                    tmp_height = 45;
                }
                #region 畢業照片
                if (!string.IsNullOrEmpty(e.FieldValue.ToString()))
                {
                    byte[] photo = Convert.FromBase64String(e.FieldValue.ToString()); //e.FieldValue as byte[];

                    if (photo != null && photo.Length > 0)
                    {
                        DocumentBuilder photoBuilder = new DocumentBuilder(e.Document);
                        photoBuilder.MoveToField(e.Field, true);
                        e.Field.Remove();
                        //Paragraph paragraph = photoBuilder.InsertParagraph();// new Paragraph(e.Document);
                        Shape photoShape = new Shape(e.Document, ShapeType.Image);
                        photoShape.ImageData.SetImage(photo);
                        photoShape.WrapType = WrapType.Inline;
                        //Cell cell = photoBuilder.CurrentParagraph.ParentNode as Cell;
                        //cell.CellFormat.LeftPadding = 0;
                        //cell.CellFormat.RightPadding = 0;
                        photoShape.Width = ConvertUtil.MillimeterToPoint(tmp_width);
                        photoShape.Height = ConvertUtil.MillimeterToPoint(tmp_height);
                        //paragraph.AppendChild(photoShape);
                        photoBuilder.InsertNode(photoShape);
                    }
                }
                #endregion
            }
        }
        void _bgw_RunWorkerCompleted(object sender, RunWorkerCompletedEventArgs e)
        {
            Document inResult = (Document)e.Result;
            btnPrint.Enabled = true;
            try
            {
                SaveFileDialog SaveFileDialog1 = new SaveFileDialog();

                SaveFileDialog1.Filter = "Word (*.doc)|*.doc|所有檔案 (*.*)|*.*";
                SaveFileDialog1.FileName = current;

                if (SaveFileDialog1.ShowDialog() == DialogResult.OK)
                {
                    inResult.Save(SaveFileDialog1.FileName);
                    Process.Start(SaveFileDialog1.FileName);
                    FISCA.Presentation.MotherForm.SetStatusBarMessage(SaveFileDialog1.FileName + ",列印完成!!");
                }
                else
                {
                    FISCA.Presentation.Controls.MsgBox.Show("檔案未儲存");
                    return;
                }
            }
            catch
            {
                string msg = "檔案儲存錯誤,請檢查檔案是否開啟中!!";
                FISCA.Presentation.Controls.MsgBox.Show(msg);
                FISCA.Presentation.MotherForm.SetStatusBarMessage(msg);
            }
        }
        private void btnExit_Click(object sender, EventArgs e)
        {
            this.Close();
        }
        private void comboBoxEx1_SelectedIndexChanged(object sender, EventArgs e)
        {
            var value = (string)comboBoxEx1.SelectedItem;
            switch (value)
            {
                case "新增":
                    AddNew input = new AddNew();
                    if (input.ShowDialog() == DialogResult.OK)
                    {
                        input.name = System.Text.RegularExpressions.Regex.Replace(input.name, @"[\W_]+", "");
                        if (string.IsNullOrWhiteSpace(input.name))
                            FISCA.Presentation.Controls.MsgBox.Show("請輸入樣板名稱(中文或英文字母)");
                        else if (custConfigs.ContainsKey(input.name))
                            FISCA.Presentation.Controls.MsgBox.Show("樣板名稱已存在");
                        else
                        {
                            ReportConfiguration tmp_conf = new ReportConfiguration(configNameRule(input.name));
                            if (input.Template != null)
                                tmp_conf.Template = new ReportTemplate(input.Template);
                            tmp_conf.Save();
                            custConfigs.Add(input.name, tmp_conf);
                            addCustConfig(input.name);
                            comboBoxEx1.Items.Insert(0, input.name);
                            comboBoxEx1.SelectedIndex = 0;
                        }
                    }
                    break;
                default:
                    current = value;
                    break;
            }
        }
        private void delete_LinkClicked(object sender, LinkLabelLinkClickedEventArgs e)
        {
            string value = (string)comboBoxEx1.SelectedItem;
            switch (value)
            {
                case "新增":
                    break;
                default:
                    if (custConfigs.ContainsKey(value))
                    {
                        custConfigs[value].Template = null;
                        custConfigs[value].Save();
                        custConfigs.Remove(value);
                        comboBoxEx1.Items.Remove(value);
                        delCustConfig(value);
                    }
                    break;
            }
            comboBoxEx1.SelectedIndex = 0;
            current = (string)comboBoxEx1.SelectedItem;
        }
        private void addCustConfig(string custConfig)
        {
            List<string> tmp = conf.GetString("customs", "").Split(';').ToList<string>();
            tmp.Add(System.Text.RegularExpressions.Regex.Replace(custConfig, @"[\W_]+", ""));
            conf.SetString("customs", string.Join(";", tmp));
            conf.Save();
        }
        private void delCustConfig(string custConfig)
        {
            List<string> tmp = conf.GetString("customs", "").Split(';').ToList<string>();
            tmp.Remove(custConfig);
            conf.SetString("customs", string.Join(";", tmp));
            conf.Save();
        }
        private string[] getCustConfig()
        {
            return conf.GetString("customs", "").Split(';');
        }
        private static string configNameRule(string custConfigName)
        {
            return config + "." + custConfigName;
        }
        public static string daySuffix(string date)
        {
            switch (int.Parse(date) % 10)
            {
                case 1: return "st";
                case 2: return "nd";
                case 3: return "rd";
                default: return "th";
            }
        }
        public static string toZhNumber(string input)
        {
            string ZhMap = "〇一二三四五六七八九";
            for (int i = 0, len = ZhMap.Length; i < len; i++)
            {
                input = input.Replace("" + i, "" + ZhMap[i]);
            }
            return input;
        }
        public static string dayToOrdinal(int input)
        {
            List<string> l = new List<string>(){
             "zeroth"
            //, "noughth"
            , "first"
            , "second"
            , "third"
            , "fourth"
            , "fifth"
            , "sixth"
            , "seventh"
            , "eighth"
            , "ninth"
            , "tenth"
            , "eleventh"
            , "twelfth"
            , "thirteenth"
            , "fourteenth"
            , "fifteenth"
            , "sixteenth"
            , "seventeenth"
            , "eighteenth"
            , "nineteenth"
            , "twentieth"
            , "twenty-first"
            , "twenty-second"
            , "twenty-third"
            , "twenty-fourth"
            , "twenty-fifth"
            , "twenty-sixth"
            , "twenty-seventh"
            , "twenty-eighth"
            , "twenty-ninth"
            , "thirtieth"
            , "thirty-first"};
            if (input <= 31)
                return l[input];
            else
                return null;
        }
    }
}
