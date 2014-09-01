using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace JH.IBSH.Diploma
{
    class Permissions
    {
        public static string 實中畢業證書 { get { return "JH.IBSH.Diploma.cs"; } }
        public static bool 實中畢業證書權限
        {
            get
            {
                return FISCA.Permission.UserAcl.Current[實中畢業證書].Executable;
            }
        }
    }
}
