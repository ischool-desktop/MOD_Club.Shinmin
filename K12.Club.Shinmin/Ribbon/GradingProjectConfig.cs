using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Windows.Forms;
using FISCA.Presentation.Controls;
using FISCA.UDT;
using FISCA.Data;

namespace K12.Club.Shinmin
{
    public partial class GradingProjectConfig : BaseForm
    {
        private AccessHelper _AccessHelper = new AccessHelper();
        private QueryHelper _QueryHelper = new QueryHelper();

        Dictionary<string, int> rowIndex = new Dictionary<string, int>();

        string PA_Name = "平時活動比例";
        string AR_Name = "出缺率比例";
        string AAS_Name = "活動力及服務比例";
        string FAR_Name = "成品成果考驗比例";

        WeightProportion wp { get; set; }

        public GradingProjectConfig()
        {
            InitializeComponent();
        }

        private void GradingProjectConfig_Load(object sender, EventArgs e)
        {
            dataGridViewX1.Rows.Clear();

            List<WeightProportion> list = _AccessHelper.Select<WeightProportion>();
            if (list.Count == 0)
            {
                this.Text = "社團成績評量項目(尚未設定)";
                wp = new WeightProportion();
                DataGridViewRow row;
                row = SetRow(PA_Name, "");
                rowIndex.Add(PA_Name, row.Index);

                row = SetRow(AR_Name, "");
                rowIndex.Add(AR_Name, row.Index);

                row = SetRow(AAS_Name, "");
                rowIndex.Add(AAS_Name, row.Index);

                row = SetRow(FAR_Name, "");
                rowIndex.Add(FAR_Name, row.Index);

            }
            else
            {

                wp = list[0];

                dataGridViewX1.Tag = wp;

                DataGridViewRow row;
                row = SetRow(PA_Name, wp.PA_Weight.ToString());
                rowIndex.Add(PA_Name, row.Index);

                row = SetRow(AR_Name, wp.AR_Weight.ToString());
                rowIndex.Add(AR_Name, row.Index);

                row = SetRow(AAS_Name, wp.AAS_Weight.ToString());
                rowIndex.Add(AAS_Name, row.Index);

                row = SetRow(FAR_Name, wp.FAR_Weight.ToString());
                rowIndex.Add(FAR_Name, row.Index);

            }

        }

        private DataGridViewRow SetRow(string Name, string wp)
        {
            DataGridViewRow row = new DataGridViewRow();
            row.CreateCells(dataGridViewX1);
            row.Cells[0].Value = Name;
            row.Cells[1].Value = wp;
            int index = dataGridViewX1.Rows.Add(row);
            return dataGridViewX1.Rows[index];
        }

        private void btnSave_Click(object sender, EventArgs e)
        {
            if (CheckData())
            {
                StringBuilder sb = new StringBuilder();
                sb.AppendLine("已修改評量比例");

                wp.PA_Weight = int.Parse("" + dataGridViewX1.Rows[rowIndex[PA_Name]].Cells[1].Value);
                wp.AR_Weight = int.Parse("" + dataGridViewX1.Rows[rowIndex[AR_Name]].Cells[1].Value);
                wp.AAS_Weight = int.Parse("" + dataGridViewX1.Rows[rowIndex[AAS_Name]].Cells[1].Value);
                wp.FAR_Weight = int.Parse("" + dataGridViewX1.Rows[rowIndex[FAR_Name]].Cells[1].Value);

                sb.AppendLine(string.Format("名稱「{0}」比例「{1}」", PA_Name, "" + wp.PA_Weight));
                sb.AppendLine(string.Format("名稱「{0}」比例「{1}」", AR_Name, "" + wp.AR_Weight));
                sb.AppendLine(string.Format("名稱「{0}」比例「{1}」", AAS_Name, "" + wp.AAS_Weight));
                sb.AppendLine(string.Format("名稱「{0}」比例「{1}」", FAR_Name, "" + wp.FAR_Weight));

                try
                {
                    List<WeightProportion> listdelete = _AccessHelper.Select<WeightProportion>();
                    _AccessHelper.DeletedValues(listdelete);

                    List<WeightProportion> list = new List<WeightProportion>();
                    list.Add(wp);
                    _AccessHelper.InsertValues(list);
                    FISCA.LogAgent.ApplicationLog.Log("社團", "修改評量比例", sb.ToString());
                }
                catch (Exception ex)
                {
                    MsgBox.Show("儲存失敗!!\n" + ex.Message);
                    SmartSchool.ErrorReporting.ReportingService.ReportException(ex);
                    return;
                }

                MsgBox.Show("儲存成功!!");
                this.Close();
            }
            else
            {
                MsgBox.Show("資料錯誤請修正後儲存!!");
            }


        }

        //檢查每一個Row的值是否正確
        private bool CheckData()
        {
            //Cell-1必須是數字,且小於100%
            //4個項目,相加後的大小必須小於100
            bool check = true;
            foreach (DataGridViewRow row in dataGridViewX1.Rows)
            {
                foreach (DataGridViewCell cell in row.Cells)
                {
                    if (cell.ColumnIndex == 1)
                    {
                        int x = 0;
                        if (!int.TryParse("" + cell.Value, out x))
                        {
                            check = false;
                            cell.ErrorText = "必須是數字";
                        }
                        else
                        {
                            cell.ErrorText = "";
                        }
                    }
                }
            }

            return check;
        }

        private void btnExit_Click(object sender, EventArgs e)
        {
            this.Close();
        }

        private void buttonX1_Click(object sender, EventArgs e)
        {
            List<WeightProportion> list = _AccessHelper.Select<WeightProportion>();
            _AccessHelper.DeletedValues(list);
        }

        private void dataGridViewX1_CurrentCellDirtyStateChanged(object sender, EventArgs e)
        {
            CheckData();
        }
    }
}
