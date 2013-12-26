using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Windows.Forms;
using System.Xml;
using FISCA.LogAgent;
using K12.Data;
using FISCA.Data;
using FISCA.UDT;
using FISCA.Presentation.Controls;

namespace K12.Club.Shinmin
{
    public partial class OpenClubJoinDateTime : FISCA.Presentation.Controls.BaseForm
    {
        //2012/7/23
        //group by出目前系統的年級資料
        //取得 DTClub ,並且判斷其開始與結束時間
        private const string DateTimeFormat = "yyyy/MM/dd HH:mm";

        AccessHelper _AccessHelper = new AccessHelper();
        QueryHelper _QueryHelper = new QueryHelper();
        List<DTClub> Low_DTClubList = new List<DTClub>();

        public OpenClubJoinDateTime()
        {
            InitializeComponent();
            
        }

        private void DailyLifeInputControl_Load(object sender, EventArgs e)
        {
            lblSemester.Text = string.Format("{0}學年度　第{1}學期", School.DefaultSchoolYear, School.DefaultSemester);

            //先將 Grid 填入此學校有的年級。
            FillGridViewGradeYear();

            //將對應年級的時間填入。
            FillTimes();
        }

        private void FillTimes()
        {
            Low_DTClubList = _AccessHelper.Select<DTClub>();

            //沒有資料就不顯示資料。
            if (Low_DTClubList.Count == 0) return;

            foreach (DTClub each in Low_DTClubList)
            {
                string grade = each.GradeYear.ToString();
                string startTime = each.Start.HasValue ? each.Start.Value.ToString(DateTimeFormat) : "";
                string endTime = each.End.HasValue ? each.End.Value.ToString(DateTimeFormat) : "";

                foreach (DataGridViewRow eachRow in dgvTimes.Rows)
                {
                    string rowgrade = eachRow.Cells[chGradeYear.Index].Value + "";

                    if (rowgrade == grade)
                    {
                        eachRow.Cells[chStartTime.Index].Value = startTime;
                        eachRow.Cells[chEndTime.Index].Value = endTime;
                        break;
                    }
                }
            }
        }

        private void FillGridViewGradeYear()
        {
            foreach (int each in GroupGradeYear().Keys)
            {
                DataGridViewRow row = new DataGridViewRow();
                row.CreateCells(dgvTimes, each.ToString(), "", "");
                dgvTimes.Rows.Add(row);
            }
        }

        private Dictionary<int, string> GroupGradeYear()
        {
            string gradeyearString = "select grade_year from class group by grade_year ORDER by grade_year";
            DataTable dtTable = _QueryHelper.Select(gradeyearString);

            Dictionary<int, string> years = new Dictionary<int, string>();
            foreach (DataRow row in dtTable.Rows)
            {
                string yearString = "" + row[0];

                int year;

                if (!int.TryParse(yearString, out year))
                    continue;

                if (!years.ContainsKey(year))
                    years.Add(year, yearString);
            }

            return years;
        }

        private void btnClose_Click(object sender, EventArgs e)
        {
            this.Close();
        }

        private void btnSave_Click(object sender, EventArgs e)
        {
            if (IsDataValidity())
            {
                #region 資料正確,進行儲存
                List<DTClub> DTClubList = new List<DTClub>();
                foreach (DataGridViewRow each in dgvTimes.Rows)
                {
                    string Grade = "" + each.Cells[chGradeYear.Index].Value;
                    string Start = "" + each.Cells[chStartTime.Index].Value;
                    string End = "" + each.Cells[chEndTime.Index].Value;

                    DTClub dt = new DTClub();

                    dt.GradeYear = int.Parse(Grade);

                    if (!string.IsNullOrEmpty(Start))
                    {
                        dt.Start = DateTime.Parse(Start);
                    }
                    if (!string.IsNullOrEmpty(End))
                    {
                        dt.End = DateTime.Parse(End);
                    }
                    DTClubList.Add(dt);
                }

                try
                {
                    _AccessHelper.InsertValues(DTClubList);
                    _AccessHelper.DeletedValues(Low_DTClubList);

                    DTClubList.Sort(SortDTClub);
                    StringBuilder sb = new StringBuilder();
                    sb.AppendLine("已修改開放選社時間");
                    foreach (DTClub each in DTClubList)
                    {
                        string start = each.Start.HasValue ? each.Start.Value.ToString("yyyy/MM/dd HH:mm") : "";
                        string end = each.End.HasValue ? each.End.Value.ToString("yyyy/MM/dd HH:mm") : "";
                        sb.AppendLine(string.Format("「{0}」年級：開始時間「{1}」結束時間「{2}」", each.GradeYear, start, end));
                    }

                    FISCA.LogAgent.ApplicationLog.Log("社團", "修改選社時間", sb.ToString());
                }
                catch (Exception ex)
                {
                    MsgBox.Show("儲存失敗!!\n" + ex.Message);
                    SmartSchool.ErrorReporting.ReportingService.ReportException(ex);
                    return;
                }
                MsgBox.Show("儲存成功!!");
                this.Close();

                #endregion
            }
            else
            {
                MsgBox.Show("畫面中含有不正確資料。");
            }
        }

        /// <summary>
        /// 依年級進行排序
        /// </summary>
        private int SortDTClub(DTClub dt1, DTClub dt2)
        {
            return dt1.GradeYear.CompareTo(dt2.GradeYear);
        }

        private bool IsDataValidity()
        {
            bool valid = true;
            foreach (DataGridViewRow each in dgvTimes.Rows)
            {
                if (!string.IsNullOrEmpty(each.ErrorText))
                {
                    valid = false;
                }

                foreach (DataGridViewCell eachCell in each.Cells)
                {
                    if (!string.IsNullOrEmpty(eachCell.ErrorText))
                    {
                        valid = false;
                    }
                }

                if (!valid) break;
            }

            return valid;
        }

        private void dgvTimes_CellClick(object sender, DataGridViewCellEventArgs e)
        {
            //dgvTimes.BeginEdit(true);
        }

        /// <summary>
        /// 開始&結束日期是否有錯誤
        /// </summary>
        private void dgvTimes_RowValidating(object sender, DataGridViewCellCancelEventArgs e)
        {
            DataGridViewRow row = dgvTimes.Rows[e.RowIndex];

            string startTime = row.Cells[chStartTime.Index].Value + "";
            string endTime = row.Cells[chEndTime.Index].Value + "";

            row.ErrorText = "";
            if (string.IsNullOrEmpty(startTime) && string.IsNullOrEmpty(endTime))
            {
                //這裡沒有程式。
            }
            else if (!string.IsNullOrEmpty(startTime) && !string.IsNullOrEmpty(endTime))
            {
                DateTime? objStart = DateTimeHelper.Parse(startTime);
                DateTime? objEnd = DateTimeHelper.Parse(endTime);

                if (objStart.HasValue && objEnd.HasValue)
                {
                    if (objStart.Value >= objEnd.Value)
                        row.ErrorText = "截止時間必須在開始時間之後。";
                }
            }
            else
                row.ErrorText = "請輸入正確的時間限制資料(必需同時有資料或同時沒有資料)。";
        }

        private void dgvTimes_CellValidating(object sender, DataGridViewCellValidatingEventArgs e)
        {
            //ValidateCellData(e.ColumnIndex, e.RowIndex, e.FormattedValue + "");
        }

        private void dgvTimes_CellEndEdit(object sender, DataGridViewCellEventArgs e)
        {
            ValidateCellData(e.ColumnIndex, e.RowIndex, dgvTimes.Rows[e.RowIndex].Cells[e.ColumnIndex].Value + "");
            //TryToCorrectData(e.ColumnIndex, e.RowIndex);
        }

        private void ValidateCellData(int columnIndex, int rowIndex, string value)
        {
            if (columnIndex == chStartTime.Index || columnIndex == chEndTime.Index)
            {
                DataGridViewCell cell = dgvTimes.Rows[rowIndex].Cells[columnIndex];
                cell.ErrorText = "";
                if (string.IsNullOrEmpty(value)) //沒有資料就不作任何檢查。
                    return;

                DateTime dt;
                if (!DateTime.TryParse(value, out dt))
                {
                    cell.ErrorText = "日期格式錯誤。";
                }
                else
                {
                    cell.Value = dt.ToString(DateTimeFormat);
                }
            }
        }

        private void dgvTimes_RowValidated(object sender, DataGridViewCellEventArgs e)
        {
            //TryToCorrectData(e.ColumnIndex, e.RowIndex);
        }

        private void TryToCorrectData(int columnIndex, int rowIndex)
        {
            if (columnIndex == chStartTime.Index || columnIndex == chEndTime.Index)
            {
                DataGridViewRow row = dgvTimes.Rows[rowIndex];
                row.Cells[columnIndex].ErrorText = string.Empty;
                string time = row.Cells[columnIndex].Value + "";

                if (string.IsNullOrEmpty(time)) //沒有資料就不作任何檢查。
                    return;

                DateTime? objStart = DateTimeHelper.ParseGregorian(time, PaddingMethod.First);

                if (objStart.HasValue)
                    row.Cells[columnIndex].Value = objStart.Value.ToString(DateTimeFormat);
            }
        }
    }
}
