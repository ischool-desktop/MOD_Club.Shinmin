using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace K12.Club.Shinmin
{
    class Permissions
    {
        public static string 新增社團 { get { return "K12.Club.Shinmin.NewAddClub.cs"; } }
        public static bool 新增社團權限
        {
            get
            {
                return FISCA.Permission.UserAcl.Current[新增社團].Executable;
            }
        }

        public static string 複製社團 { get { return "K12.Club.Shinmin.CopyClub.cs"; } }

        public static bool 複製社團權限
        {
            get
            {
                return FISCA.Permission.UserAcl.Current[複製社團].Executable;
            }
        }

        public static string 刪除社團 { get { return "K12.Club.Shinmin.DeleteClub.cs"; } }
        public static bool 刪除社團權限
        {
            get
            {
                return FISCA.Permission.UserAcl.Current[刪除社團].Executable;
            }
        }

        public static string 社團基本資料 { get { return "K12.Club.Shinmin.ClubDetailItem.cs"; } }
        public static bool 社團基本資料權限
        {
            get
            {
                return FISCA.Permission.UserAcl.Current[社團基本資料].Executable;
            }
        }

        public static string 社團照片 { get { return "K12.Club.Shinmin.ClubImageItem.cs"; } }
        public static bool 社團照片權限
        {
            get
            {
                return FISCA.Permission.UserAcl.Current[社團照片].Executable;
            }
        }

        public static string 社團限制 { get { return "K12.Club.Shinmin.ClubRestrictItem.cs"; } }
        public static bool 社團限制權限
        {
            get
            {
                return FISCA.Permission.UserAcl.Current[社團限制].Executable;
            }
        }

        public static string 社團參與學生 { get { return "K12.Club.Shinmin.ClubStudent.cs"; } }
        public static bool 社團參與學生權限
        {
            get
            {
                return FISCA.Permission.UserAcl.Current[社團參與學生].Executable;
            }
        }

        public static string 未選社團學生 { get { return "K12.Club.Shinmin.CheckStudentIsNotInClub.cs"; } }

        public static bool 未選社團學生權限
        {
            get
            {
                return FISCA.Permission.UserAcl.Current[未選社團學生].Executable;
            }
        }

        public static string 調整社團學生 { get { return "K12.Club.Shinmin.SplitClasses.cs"; } }

        public static bool 調整社團學生權限
        {
            get
            {
                return FISCA.Permission.UserAcl.Current[調整社團學生].Executable;
            }
        }

        public static string 評量項目 { get { return "K12.Club.Shinmin.GradingProjectConfig.cs"; } }

        public static bool 評量項目權限
        {
            get
            {
                return FISCA.Permission.UserAcl.Current[評量項目].Executable;
            }
        }

        public static string 社團幹部 { get { return "K12.Club.Shinmin.CadresItem.cs"; } }
        public static bool 社團幹部權限
        {
            get
            {
                return FISCA.Permission.UserAcl.Current[社團幹部].Executable;
            }
        }

        public static string 成績輸入 { get { return "K12.Club.Shinmin.ClubResultsInput.cs"; } }
        public static bool 成績輸入權限
        {
            get
            {
                return FISCA.Permission.UserAcl.Current[成績輸入].Executable;
            }
        }


        public static string 開放選社時間 { get { return "K12.Club.Shinmin.OpenClubJoinDateTime.cs"; } }

        public static bool 開放選社時間權限
        {
            get
            {
                return FISCA.Permission.UserAcl.Current[開放選社時間].Executable;
            }
        }

        public static string 成績輸入時間 { get { return "K12.Club.Shinmin.ResultsInputDateTime.cs"; } }
        public static bool 成績輸入時間權限
        {
            get
            {
                return FISCA.Permission.UserAcl.Current[成績輸入時間].Executable;
            }
        }

        public static string 社團點名單 { get { return "K12.Club.Shinmin.ClubPointList.cs"; } }
        public static bool 社團點名單權限
        {
            get
            {
                return FISCA.Permission.UserAcl.Current[社團點名單].Executable;
            }
        }

        public static string 班級學生選社_確認表 { get { return "K12.Club.Shinmin.ElectionSocialConfirmation.cs"; } }
        public static bool 班級學生選社_確認表_權限
        {
            get
            {
                return FISCA.Permission.UserAcl.Current[班級學生選社_確認表].Executable;
            }
        }

        public static string 重覆選社檢查 { get { return "K12.Club.Shinmin.RepeatForm.cs"; } }
        public static bool 重覆選社檢查權限
        {
            get
            {
                return FISCA.Permission.UserAcl.Current[重覆選社檢查].Executable;
            }
        }

        public static string 學期結算 { get { return "K12.Club.Shinmin.ClearingForm.cs"; } }
        public static bool 學期結算權限
        {
            get
            {
                return FISCA.Permission.UserAcl.Current[學期結算].Executable;
            }
        }

        public static string 學生社團成績_資料項目 { get { return "K12.Club.Shinmin.StudentResultItem.cs"; } }
        public static bool 學生社團成績_資料項目權限
        {
            get
            {
                return FISCA.Permission.UserAcl.Current[學生社團成績_資料項目].Executable;
            }
        }

        public static string 社團成績單 { get { return "K12.Club.Shinmin.ClubTranscript.cs"; } }
        public static bool 社團成績單權限
        {
            get
            {
                return FISCA.Permission.UserAcl.Current[社團成績單].Executable;
            }
        }

        public static string 班級社團成績單 { get { return "K12.Club.Shinmin.ClassClubTranscript.cs"; } }
        public static bool 班級社團成績單權限
        {
            get
            {
                return FISCA.Permission.UserAcl.Current[班級社團成績單].Executable;
            }
        }

        public static string 社團概況表 { get { return "K12.Club.Shinmin.CLUBFactsTable.cs"; } }
        public static bool 社團概況表權限
        {
            get
            {
                return FISCA.Permission.UserAcl.Current[社團概況表].Executable;
            }
        }

        public static string 匯出社團學期成績 { get { return "K12.Club.Shinmin.ExportClubResult.cs"; } }
        public static bool 匯出社團學期成績權限
        {
            get
            {
                return FISCA.Permission.UserAcl.Current[匯出社團學期成績].Executable;
            }
        }

        public static string 匯出社團成績_資料介接 { get { return "K12.Club.Shinmin.SpecialResult.cs"; } }
        public static bool 匯出社團成績_資料介接權限
        {
            get
            {
                return FISCA.Permission.UserAcl.Current[匯出社團成績_資料介接].Executable;
            }
        }

        public static string 社團幹部證明單 { get { return "K12.Club.Shinmin.CadreProveReport.cs"; } }
        public static bool 社團幹部證明單權限
        {
            get
            {
                return FISCA.Permission.UserAcl.Current[社團幹部證明單].Executable;
            }
        }

        public static string 匯出社團幹部清單 { get { return "K12.Club.Shinmin.ClubCadResult.cs"; } }
        public static bool 匯出社團幹部清單權限
        {
            get
            {
                return FISCA.Permission.UserAcl.Current[匯出社團幹部清單].Executable;
            }
        }
    }
}
