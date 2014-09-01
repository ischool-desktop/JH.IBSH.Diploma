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
            item["畢業證(明)書"].Enable = Permissions.實中畢業證書權限;
            item["畢業證(明)書"].Click += delegate
            {
                new MainForm().ShowDialog();
            };
            Catalog detail1 = RoleAclSource.Instance["學生"]["報表"];
            detail1.Add(new RibbonFeature(Permissions.實中畢業證書, "畢業證(明)書"));
        }
    }
}
