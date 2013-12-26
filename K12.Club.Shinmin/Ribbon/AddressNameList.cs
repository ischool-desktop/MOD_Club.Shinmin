using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Windows.Forms;
using FISCA.Presentation.Controls;

namespace K12.Club.Shinmin
{
    //上課地點清單管理介面
    public partial class AddressNameList : BaseForm 
    {
        K12.Data.Configuration.ConfigData DateConfig = K12.Data.School.Configuration["社團模組_上課地點清單"];

        public AddressNameList()
        {
            InitializeComponent();

            if (DateConfig.Count != 0)
            {
                foreach (string each in DateConfig["上課地點清單"].Split(','))
                {
                    if (!string.IsNullOrEmpty(each))
                    {
                        DataGridViewRow row = new DataGridViewRow();
                        row.CreateCells(dataGridViewX1);
                        row.Cells[0].Value = each;
                        dataGridViewX1.Rows.Add(row);
                    }
                }
            }
        }

        private void buttonX1_Click(object sender, EventArgs e)
        {
            List<string> list = new List<string>();

            foreach (DataGridViewRow eachRow in dataGridViewX1.Rows)
            {
                if (!eachRow.IsNewRow)
                {
                    list.Add("" + eachRow.Cells[0].Value);
                }
            }

            DateConfig["上課地點清單"] = string.Join(",", list.ToArray());

            DateConfig.Save();
            this.Close();
        }

        private void buttonX2_Click(object sender, EventArgs e)
        {
            this.Close();
        }
    }
}
