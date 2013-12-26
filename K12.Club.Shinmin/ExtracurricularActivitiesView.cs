using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Drawing;
using System.Data;
using System.Linq;
using System.Text;
using System.Windows.Forms;
using FISCA.Presentation;
using K12.Data;
using FISCA.UDT;

namespace K12.Club.Shinmin
{
    public partial class ExtracurricularActivitiesView : NavView
    {

        private Dictionary<string, List<string>> Dic = new Dictionary<string, List<string>>();

        string CurrNode = "";

        public ExtracurricularActivitiesView()
        {
            InitializeComponent();

            NavText = "聯課活動檢視";

            SourceChanged += new EventHandler(ExtracurricularActivitiesView_SourceChanged);
        }

        void ExtracurricularActivitiesView_SourceChanged(object sender, EventArgs e)
        {
            if (advTree1.SelectedNode != null)
            {
                if (advTree1.SelectedNode.Tag != null)
                {
                    CurrNode = "" + advTree1.SelectedNode.Tag;
                }
            }

            Dic.Clear();
            advTree1.Nodes.Clear();

            DevComponents.AdvTree.Node Node1 = new DevComponents.AdvTree.Node();
            Node1.Text = "所有分類(" + Source.Count() + ")";
            Node1.Tag = "All";
            advTree1.Nodes.Add(Node1); //加入

            if (Source.Count != 0)
            {
                AccessHelper _AccessHelper = new AccessHelper();

                string str = string.Join("','", Source);
                List<CLUBRecord> clubList = _AccessHelper.Select<CLUBRecord>("uid in ('" + str + "')");

                clubList.Sort(SortClub); //Sort

                List<string> list = new List<string>();

                foreach (CLUBRecord each in clubList)
                {
                    if (string.IsNullOrEmpty(each.ClubCategory) || each.ClubCategory == "未分類")
                    {
                        list.Add(each.UID);
                        continue;
                    }

                    if (!Dic.ContainsKey(each.ClubCategory))
                        Dic.Add(each.ClubCategory, new List<string>());

                    Dic[each.ClubCategory].Add(each.UID);
                }

                Dic.Add("未分類", list);

                foreach (string each in Dic.Keys)
                {
                    //增加分類Node
                    DevComponents.AdvTree.Node Node2 = new DevComponents.AdvTree.Node();
                    Node2.Text = each + "(" + Dic[each].Count + ")";
                    Node2.Tag = each;
                    Node1.Nodes.Add(Node2);
                }

                if (string.IsNullOrEmpty(CurrNode) || CurrNode == "All")
                {
                    advTree1.SelectedNode = Node1;
                    SetListPaneSource(Source, false, false);
                }
                else
                {
                    foreach (DevComponents.AdvTree.Node each in Node1.Nodes)
                    {
                        if ("" + each.Tag == CurrNode)
                        {
                            if (Dic.ContainsKey(CurrNode))
                            {
                                advTree1.SelectedNode = each;
                                SetListPaneSource(Dic[CurrNode], false, false);
                                return;
                            }
                        }
                    }
                }
            }
            else
            {
                SetListPaneSource(Source, false, false);
            }
        }

        private int SortClub(CLUBRecord cr1, CLUBRecord cr2)
        {
            return cr1.ClubCategory.CompareTo(cr2.ClubCategory);
        }

        private void advTree1_NodeClick(object sender, DevComponents.AdvTree.TreeNodeMouseEventArgs e)
        {
            //判斷是否有按Control,Shift
            bool SelectedAll = (Control.ModifierKeys & Keys.Control) == Keys.Control;
            bool AddToTemp = (Control.ModifierKeys & Keys.Shift) == Keys.Shift;

            if ("" + e.Node.Tag == "All")
            {
                SetListPaneSource(Source, SelectedAll, AddToTemp);
            }
            else if (Dic.ContainsKey("" + e.Node.Tag)) //當使用者是選取類別名稱
            {
                SetListPaneSource(Dic["" + e.Node.Tag], SelectedAll, AddToTemp);
            }
            else //未選取
            {

            }
        }
    }
}
