using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using FISCA;
using FISCA.Presentation;
using FISCA.Permission;
using FISCA.UDT;
using FISCA.Presentation.Controls;
using System.Windows.Forms;
using K12.Data;
using K12.Data.Configuration;

namespace K12.Club.Shinmin
{
    public class Program
    {
        [MainMethod()]
        static public void Main()
        {
            ServerModule.AutoManaged("https://module.ischool.com.tw/module/138/Club_Shinmin/udm.xml");

            //FISCA.RTOut.WriteLine("註冊Gadget - 參加社團(學生)：" + WebPackage.RegisterGadget("Student", "fd56eafc-3601-40a0-82d9-808f72a8272b", "參加社團(學生)").Item2);
            //FISCA.RTOut.WriteLine("註冊Gadget - 社團(老師)：" + WebPackage.RegisterGadget("Teacher", "6080a7c0-60e7-443c-bad7-ecccb3a86bcf", "社團(老師)").Item2);


            #region 處理UDT Table沒有的問題

            ConfigData cd = K12.Data.School.Configuration["新民社團UDT載入設定"];
            bool checkClubUDT = false;

            string name = "社團UDT是否已載入";
            //如果尚無設定值,預設為
            if (string.IsNullOrEmpty(cd[name]))
            {
                cd[name] = "false";
            }

            //檢查是否為布林
            bool.TryParse(cd[name], out checkClubUDT);

            if (!checkClubUDT)
            {
                AccessHelper _accessHelper = new AccessHelper();
                _accessHelper.Select<CLUBRecord>("UID = '00000'");
                _accessHelper.Select<SCJoin>("UID = '00000'");
                _accessHelper.Select<WeightProportion>("UID = '00000'");
                _accessHelper.Select<CadresRecord>("UID = '00000'");
                _accessHelper.Select<DTScore>("UID = '00000'");
                _accessHelper.Select<DTClub>("UID = '00000'");
                _accessHelper.Select<ResultScoreRecord>("UID = '00000'");

                cd[name] = "true";
                cd.Save();
            }

            #endregion

            //增加一個社團Tab
            MotherForm.AddPanel(ClubAdmin.Instance);

            //增加一個ListView
            ClubAdmin.Instance.AddView(new ExtracurricularActivitiesView());

            //學生社團成績
            FeatureAce UserPermission = FISCA.Permission.UserAcl.Current[Permissions.學生社團成績_資料項目];
            if (UserPermission.Editable || UserPermission.Viewable)
                K12.Presentation.NLDPanels.Student.AddDetailBulider(new FISCA.Presentation.DetailBulider<StudentResultItem>());

            //社團照片
            UserPermission = FISCA.Permission.UserAcl.Current[Permissions.社團照片];
            if (UserPermission.Editable || UserPermission.Viewable)
                ClubAdmin.Instance.AddDetailBulider(new FISCA.Presentation.DetailBulider<ClubImageItem>());

            //社團基本資料
            UserPermission = FISCA.Permission.UserAcl.Current[Permissions.社團基本資料];
            if (UserPermission.Editable || UserPermission.Viewable)
                ClubAdmin.Instance.AddDetailBulider(new FISCA.Presentation.DetailBulider<ClubDetailItem>());

            //社團限制
            UserPermission = FISCA.Permission.UserAcl.Current[Permissions.社團限制];
            if (UserPermission.Editable || UserPermission.Viewable)
                ClubAdmin.Instance.AddDetailBulider(new FISCA.Presentation.DetailBulider<ClubRestrictItem>());

            //社團學生
            UserPermission = FISCA.Permission.UserAcl.Current[Permissions.社團參與學生];
            if (UserPermission.Editable || UserPermission.Viewable)
                ClubAdmin.Instance.AddDetailBulider(new FISCA.Presentation.DetailBulider<ClubStudent>());

            //社團幹部
            UserPermission = FISCA.Permission.UserAcl.Current[Permissions.社團幹部];
            if (UserPermission.Editable || UserPermission.Viewable)
                ClubAdmin.Instance.AddDetailBulider(new FISCA.Presentation.DetailBulider<CadresItem>());

            #region 功能登錄

            RibbonBarItem edit = ClubAdmin.Instance.RibbonBarItems["編輯"];
            edit["新增社團"].Size = RibbonBarButton.MenuButtonSize.Large;
            edit["新增社團"].Image = Properties.Resources.health_and_leisure_add_64;
            edit["新增社團"].Enable = Permissions.新增社團權限;
            edit["新增社團"].Click += delegate
            {
                NewAddClub insert = new NewAddClub();
                insert.ShowDialog();
            };

            edit["複製社團"].Size = RibbonBarButton.MenuButtonSize.Large;
            edit["複製社團"].Image = Properties.Resources.rotate_64;
            edit["複製社團"].Enable = false;
            edit["複製社團"].Click += delegate
            {
                CopyClub insert = new CopyClub();
                insert.ShowDialog();
            };

            edit["刪除社團"].Size = RibbonBarButton.MenuButtonSize.Large;
            edit["刪除社團"].Image = Properties.Resources.health_and_leisure_remove_64;
            edit["刪除社團"].Enable = false;
            edit["刪除社團"].Click += delegate
            {
                DeleteClub();
            };

            RibbonBarItem totle = ClubAdmin.Instance.RibbonBarItems["資料統計"];
            totle["匯出"].Size = RibbonBarButton.MenuButtonSize.Large;
            totle["匯出"].Image = Properties.Resources.Export_Image;
            totle["匯出"]["匯出聯課活動成績(資料介接)"].Enable = Permissions.匯出社團成績_資料介接權限;
            totle["匯出"]["匯出聯課活動成績(資料介接)"].Click += delegate
            {
                SmartSchool.API.PlugIn.Export.Exporter exporter = new K12.Club.Shinmin.CLUB.SpecialResult();
                K12.Club.Shinmin.CLUB.ExportStudentV2 wizard = new K12.Club.Shinmin.CLUB.ExportStudentV2(exporter.Text, exporter.Image);
                exporter.InitializeExport(wizard);
                wizard.ShowDialog();
            };

            totle["匯出"]["匯出社團幹部清單"].Enable = Permissions.匯出社團幹部清單權限;
            totle["匯出"]["匯出社團幹部清單"].Click += delegate
            {
                SmartSchool.API.PlugIn.Export.Exporter exporter = new K12.Club.Shinmin.CLUB.ClubCadResult();
                K12.Club.Shinmin.CLUB.ExportStudentV2 wizard = new K12.Club.Shinmin.CLUB.ExportStudentV2(exporter.Text, exporter.Image);
                exporter.InitializeExport(wizard);
                wizard.ShowDialog();
            };

            totle["報表"].Size = RibbonBarButton.MenuButtonSize.Large;
            totle["報表"].Image = Properties.Resources.Report;

            totle["報表"]["社團點名單"].Enable = false;
            totle["報表"]["社團點名單"].Click += delegate
            {
                AssociationsPointList insert = new AssociationsPointList();
            };

            totle["報表"]["社團成績單"].Enable = false;
            totle["報表"]["社團成績單"].Click += delegate
            {
                ClubTranscript insert = new ClubTranscript();
            };

            totle["報表"]["社團概況表"].Enable = Permissions.社團概況表權限;
            totle["報表"]["社團概況表"].Click += delegate
            {
                CLUBFactsTable insert = new CLUBFactsTable();
                insert.ShowDialog();
            };

            //totle["報表"]["社團成績單"].Enable = false;
            //totle["報表"]["社團成績單"].Click += delegate
            //{
            //    ClubTranscript insert = new ClubTranscript();
            //};

            RibbonBarItem Print = FISCA.Presentation.MotherForm.RibbonBarItems["學生", "資料統計"];
            Print["匯出"]["社團相關匯出"]["匯出社團學期成績"].Enable = Permissions.匯出社團學期成績權限;
            Print["匯出"]["社團相關匯出"]["匯出社團學期成績"].Click += delegate
            {
                SmartSchool.API.PlugIn.Export.Exporter exporter = new ExportStudentClubResult();
                ExportStudentV2 wizard = new ExportStudentV2(exporter.Text, exporter.Image);
                exporter.InitializeExport(wizard);
                wizard.ShowDialog();
            };

            Print["報表"]["社團相關報表"]["社團幹部證明單"].Enable = Permissions.社團幹部證明單權限;
            Print["報表"]["社團相關報表"]["社團幹部證明單"].Click += delegate
            {
                CadreProveReport cpr = new CadreProveReport();
                cpr.ShowDialog();
            };

            RibbonBarItem InClass = FISCA.Presentation.MotherForm.RibbonBarItems["班級", "資料統計"];
            InClass["報表"]["社團相關報表"]["班級學生選社同意確認單"].Enable = false;
            InClass["報表"]["社團相關報表"]["班級學生選社同意確認單"].Click += delegate
            {
                ElectionForm insert = new ElectionForm();
                insert.ShowDialog();
            };

            InClass["報表"]["社團相關報表"]["班級社團成績單"].Enable = false;
            InClass["報表"]["社團相關報表"]["班級社團成績單"].Click += delegate
            {
                ClassClubTranscript insert = new ClassClubTranscript();
                insert.ShowDialog();
            };

            RibbonBarItem check = ClubAdmin.Instance.RibbonBarItems["檢查"];

            check["未選社團檢查"].Size = RibbonBarButton.MenuButtonSize.Medium;
            check["未選社團檢查"].Image = Properties.Resources.group_help_64;
            check["未選社團檢查"].Enable = Permissions.未選社團學生權限;
            check["未選社團檢查"].Click += delegate
            {
                CheckStudentIsNotInClub insert = new CheckStudentIsNotInClub();
                insert.ShowDialog();
            };

            check["重覆選社檢查"].Size = RibbonBarButton.MenuButtonSize.Medium;
            check["重覆選社檢查"].Image = Properties.Resources.meeting_64;
            check["重覆選社檢查"].Enable = Permissions.重覆選社檢查權限;
            check["重覆選社檢查"].Click += delegate
            {
                RepeatForm insert = new RepeatForm();
                insert.ShowDialog();
            };

            check["調整社團學生"].Size = RibbonBarButton.MenuButtonSize.Medium;
            check["調整社團學生"].Image = Properties.Resources.layers_64;
            check["調整社團學生"].Enable = false;
            check["調整社團學生"].Click += delegate
            {
                if (ClubAdmin.Instance.SelectedSource.Count > 7)
                {
                    MsgBox.Show("所選社團大於7個\n本功能最多僅處理7個社團!!");
                }
                else if (ClubAdmin.Instance.SelectedSource.Count < 2)
                {
                    MsgBox.Show("使用調整社團學生功能\n必須2個以上社團!!");
                }
                else
                {
                    SplitClasses insert = new SplitClasses();
                    insert.ShowDialog();
                }
            };

            RibbonBarItem Results = ClubAdmin.Instance.RibbonBarItems["成績"];
            Results["成績輸入"].Size = RibbonBarButton.MenuButtonSize.Medium;
            Results["成績輸入"].Image = Properties.Resources.marker_fav_64;
            Results["成績輸入"].Enable = false;
            Results["成績輸入"].Click += delegate
            {
                ClubResultsInput insert = new ClubResultsInput();
                insert.ShowDialog();
            };

            Results["評量比例"].Size = RibbonBarButton.MenuButtonSize.Medium;
            Results["評量比例"].Image = Properties.Resources.barchart_64;
            Results["評量比例"].Enable = Permissions.評量項目權限;
            Results["評量比例"].Click += delegate
            {
                GradingProjectConfig insert = new GradingProjectConfig();
                insert.ShowDialog();
            };

            Results["學期結算"].Size = RibbonBarButton.MenuButtonSize.Medium;
            Results["學期結算"].Image = Properties.Resources.brand_write_64;
            Results["學期結算"].Enable = false;
            Results["學期結算"].Click += delegate
            {
                ClearingForm insert = new ClearingForm();
                insert.ShowDialog();
            };

            RibbonBarItem oder = ClubAdmin.Instance.RibbonBarItems["其它"];

            oder["開放選社時間"].Size = RibbonBarButton.MenuButtonSize.Medium;
            oder["開放選社時間"].Image = Properties.Resources.time_frame_refresh_128;
            oder["開放選社時間"].Enable = Permissions.開放選社時間權限;
            oder["開放選社時間"].Click += delegate
            {
                OpenClubJoinDateTime insert = new OpenClubJoinDateTime();
                insert.ShowDialog();
            };

            oder["成績輸入時間"].Size = RibbonBarButton.MenuButtonSize.Medium;
            oder["成績輸入時間"].Image = Properties.Resources.time_frame_refresh_128;
            oder["成績輸入時間"].Enable = Permissions.成績輸入時間權限;
            oder["成績輸入時間"].Click += delegate
            {
                ResultsInputDateTime insert = new ResultsInputDateTime();
                insert.ShowDialog();
            };

            ClubAdmin.Instance.NavPaneContexMenu["重新整理"].Click += delegate
            {
                ClubEvents.RaiseAssnChanged();
            };

            ClubAdmin.Instance.ListPaneContexMenu["刪除社團"].Enable = false;
            ClubAdmin.Instance.ListPaneContexMenu["刪除社團"].Click += delegate
            {
                DeleteClub();
            };

            //RibbonBarItem Other = ClubAdmin.Instance.RibbonBarItems["其它"];
            //Other["場地管裡"].Size = RibbonBarButton.MenuButtonSize.Large;
            //Other["場地管裡"].Image = Properties.Resources.architecture_zoom_64;
            //Other["場地管裡"].Click += delegate
            //{
            //    AddressNameList insert = new AddressNameList();
            //    insert.ShowDialog();
            //};
            K12.Presentation.NLDPanels.Class.SelectedSourceChanged += delegate
            {
                //是否選擇大於0的社團
                bool SourceCount = (K12.Presentation.NLDPanels.Class.SelectedSource.Count > 0);

                bool a = (SourceCount && Permissions.班級學生選社_確認表_權限);
                InClass["報表"]["社團相關報表"]["班級學生選社同意確認單"].Enable = a;


                bool b = (SourceCount && Permissions.班級社團成績單權限);
                InClass["報表"]["社團相關報表"]["班級社團成績單"].Enable = b;

            };


            ClubAdmin.Instance.SelectedSourceChanged += delegate
            {
                //是否選擇大於0的社團
                bool SourceCount = (ClubAdmin.Instance.SelectedSource.Count > 0);
                //刪除社團
                bool a = (SourceCount && Permissions.刪除社團權限);
                ClubAdmin.Instance.ListPaneContexMenu["刪除社團"].Enable = a;
                edit["刪除社團"].Enable = a;

                //複製社團
                bool b = (SourceCount && Permissions.複製社團權限);
                edit["複製社團"].Enable = b;

                bool c = (SourceCount && Permissions.調整社團學生權限);
                check["調整社團學生"].Enable = c;

                bool d = (SourceCount && Permissions.成績輸入權限);
                Results["成績輸入"].Enable = d;

                bool e = (SourceCount && Permissions.社團點名單權限);
                totle["報表"]["社團點名單"].Enable = e;

                bool f = (SourceCount && Permissions.學期結算權限);
                Results["學期結算"].Enable = f;


                bool g = (SourceCount && Permissions.社團成績單權限);
                totle["報表"]["社團成績單"].Enable = g;

                //bool f = (SourceCount && Permissions.社團成績單權限);
                //totle["報表"]["社團成績單"].Enable = f;

                FISCA.Presentation.MotherForm.SetStatusBarMessage("選擇「" + ClubAdmin.Instance.SelectedSource.Count + "」個社團");
            };

            #endregion

            #region 登錄權限代碼

            //是否能夠只用單一代碼,決定此模組之使用
            Catalog detail1;
            detail1 = RoleAclSource.Instance["社團"]["功能按鈕"];
            detail1.Add(new RibbonFeature(Permissions.新增社團, "新增社團"));
            detail1.Add(new RibbonFeature(Permissions.複製社團, "複製社團"));
            detail1.Add(new RibbonFeature(Permissions.刪除社團, "刪除社團"));
            detail1.Add(new RibbonFeature(Permissions.成績輸入, "成績輸入"));
            detail1.Add(new RibbonFeature(Permissions.評量項目, "評量比例"));
            detail1.Add(new RibbonFeature(Permissions.學期結算, "學期結算"));
            detail1.Add(new RibbonFeature(Permissions.未選社團學生, "未選社團學生"));
            detail1.Add(new RibbonFeature(Permissions.調整社團學生, "調整社團學生"));
            detail1.Add(new RibbonFeature(Permissions.開放選社時間, "開放選社時間"));
            detail1.Add(new RibbonFeature(Permissions.成績輸入時間, "成績輸入時間"));
            detail1.Add(new RibbonFeature(Permissions.重覆選社檢查, "重覆選社檢查"));
            detail1.Add(new RibbonFeature(Permissions.匯出社團成績_資料介接, "匯出社團學期成績(資料介接)"));
            detail1.Add(new RibbonFeature(Permissions.匯出社團幹部清單, "匯出社團幹部清單"));

            detail1 = RoleAclSource.Instance["社團"]["報表"];
            detail1.Add(new RibbonFeature(Permissions.社團點名單, "社團點名單"));
            detail1.Add(new RibbonFeature(Permissions.社團成績單, "社團成績單"));
            detail1.Add(new RibbonFeature(Permissions.社團概況表, "社團概況表"));

            detail1 = RoleAclSource.Instance["社團"]["資料項目"];
            detail1.Add(new DetailItemFeature(Permissions.社團基本資料, "基本資料"));
            detail1.Add(new DetailItemFeature(Permissions.社團照片, "社團照片"));
            detail1.Add(new DetailItemFeature(Permissions.社團限制, "社團限制"));
            detail1.Add(new DetailItemFeature(Permissions.社團參與學生, "參與學生"));
            detail1.Add(new DetailItemFeature(Permissions.社團幹部, "社團幹部"));

            detail1 = RoleAclSource.Instance["班級"]["報表"];
            detail1.Add(new RibbonFeature(Permissions.班級學生選社_確認表, "班級學生選社同意確認單"));
            detail1.Add(new RibbonFeature(Permissions.班級社團成績單, "班級社團成績單"));

            detail1 = RoleAclSource.Instance["學生"]["功能按鈕"];
            detail1.Add(new RibbonFeature(Permissions.匯出社團學期成績, "匯出社團學期成績"));    

            detail1 = RoleAclSource.Instance["學生"]["資料項目"];
            detail1.Add(new DetailItemFeature(Permissions.學生社團成績_資料項目, "社團成績"));

            detail1 = RoleAclSource.Instance["學生"]["報表"];
            detail1.Add(new RibbonFeature(Permissions.社團幹部證明單, "社團幹部證明單"));
            #endregion
        }

        static private void DeleteClub()
        {
            DialogResult dr = MsgBox.Show("確認刪除所選社團?", MessageBoxButtons.YesNo, MessageBoxDefaultButton.Button2);
            if (dr == DialogResult.Yes)
            {
                StringBuilder sb = new StringBuilder();
                sb.AppendLine("已刪除選擇社團：");
                List<CLUBRecord> ClubList = tool._A.Select<CLUBRecord>(UDT_S.PopOneCondition("UID", ClubAdmin.Instance.SelectedSource));

                foreach (CLUBRecord each in ClubList)
                {
                    sb.AppendLine(string.Format("學年度「{0}」學期「{1}」社團名稱「{2}」", each.SchoolYear.ToString(), each.Semester.ToString(), each.ClubName));
                }

                List<SCJoin> SCJList = tool._A.Select<SCJoin>(UDT_S.PopOneCondition("ref_club_id", ClubAdmin.Instance.SelectedSource));
                if (SCJList.Count != 0)
                {
                    MsgBox.Show("刪除社團必須清空社團參與學生!");
                    return;
                }

                try
                {
                    tool._A.DeletedValues(ClubList);
                }
                catch (Exception ex)
                {
                    MsgBox.Show("社團刪除失敗!!\n" + ex.Message);
                    SmartSchool.ErrorReporting.ReportingService.ReportException(ex);
                    return;

                }
                FISCA.LogAgent.ApplicationLog.Log("社團", "刪除社團", sb.ToString());
                MsgBox.Show("社團刪除成功!!");
                ClubEvents.RaiseAssnChanged();
            }
            else
            {
                MsgBox.Show("已中止刪除社團操作!!");
            }
        }
    }
}
