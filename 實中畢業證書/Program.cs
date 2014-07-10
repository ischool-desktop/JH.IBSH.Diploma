using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using FISCA;
using FISCA.Presentation;
using FISCA.Permission;

namespace JH.IBSH.Diploma
{
    public class Program
    {
        [MainMethod()]
        static public void Main()
        {
            MenuButton item = K12.Presentation.NLDPanels.Student.RibbonBarItems["資料統計"]["報表"]["學籍相關報表"];
            //Print["報表"]["學籍相關報表"]["畢業證(明)書"].Enable = 社團幹部證明單權限;
            item["畢業證(明)書"].Click += delegate
            {
                new MainForm().ShowDialog();
            };
            //Catalog detail1 = RoleAclSource.Instance["學生"]["報表"];
            //detail1.Add(new RibbonFeature(社團幹部證明單, "社團幹部證明單"));
        }

        //static public string 社團幹部證明單 { get { return "K12.Club.Universal.CadreProveReport.cs"; } }
        //static public bool 社團幹部證明單權限
        //{
        //    get
        //    {
        //        return FISCA.Permission.UserAcl.Current[社團幹部證明單].Executable;
        //    }
        //}
    }
}
