using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace JH.IBSH.Diploma
{
    class Permissions
    {
        public static string 服務學習記錄 { get { return "K12.Service.Learning.Modules.LearningItem.cs"; } }
        public static bool 服務學習記錄權限
        {
            get
            {
                return FISCA.Permission.UserAcl.Current[服務學習記錄].Executable;
            }
        }

        public static string 服務學習快速登錄 { get { return "K12.Service.Learning.Modules.MutiLearning.cs"; } }
        public static bool 服務學習快速登錄權限
        {
            get
            {
                return FISCA.Permission.UserAcl.Current[服務學習快速登錄].Executable;
            }
        }

        public static string 服務學習_證明單 { get { return "K12.Service.Learning.Modules.SLReportForm.cs"; } }
        public static bool 服務學習_證明單權限
        {
            get
            {
                return FISCA.Permission.UserAcl.Current[服務學習_證明單].Executable;
            }
        }

        public static string 匯出服務學習記錄 { get { return "K12.Service.Learning.Modules.ExportSLRecord.cs"; } }
        public static bool 匯出服務學習記錄權限
        {
            get
            {
                return FISCA.Permission.UserAcl.Current[匯出服務學習記錄].Executable;
            }
        }

        public static string 匯入服務學習記錄 { get { return "K12.Service.Learning.Modules.ImportSLRecord.cs"; } }
        public static bool 匯入服務學習記錄權限
        {
            get
            {
                return FISCA.Permission.UserAcl.Current[匯入服務學習記錄].Executable;
            }
        }

        public static string 服務學習批次修改 { get { return "K12.Service.Learning.Modules.ServiceLearningBatch.cs"; } }
        public static bool 服務學習批次修改權限
        {
            get
            {
                return FISCA.Permission.UserAcl.Current[服務學習批次修改].Executable;
            }
        }

        public static string 學生服務時數查詢 { get { return "K12.Service.Learning.Modules.StudentIvForm.cs"; } }
        public static bool 學生服務時數查詢權限
        {
            get
            {
                return FISCA.Permission.UserAcl.Current[學生服務時數查詢].Executable;
            }
        }

        public static string 服務學習事由代碼表 { get { return "K12.Service.Learning.Modules.DigitalCodeForm.cs"; } }
        public static bool 服務學習事由代碼表權限
        {
            get
            {
                return FISCA.Permission.UserAcl.Current[服務學習事由代碼表].Executable;
            }
        }
    }
}
