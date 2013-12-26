using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using DevComponents.DotNetBar.Controls;
using DevComponents.DotNetBar;
using FISCA.Data;
using System.Data;
using System.Drawing;
using System.Drawing.Drawing2D;
using K12.Data;
using FISCA.UDT;

namespace K12.Club.Shinmin
{
    static class tool
    {
        static public AccessHelper _A = new AccessHelper();
        static public QueryHelper _Q = new QueryHelper();

        /// <summary>
        /// 檢查傳入的文字是否為數字型態,
        /// 回傳Parse後的數字,
        /// 當文字不是數字型態時,
        /// 會回傳預設數字為0,
        /// </summary>
        static public int StringIsInt_DefIsZero(string p)
        {
            int k = 0;
            int.TryParse(p, out k);
            return k;
        }

        /// <summary>
        /// 檢查傳入的文字是否為數字型態,
        /// 回傳Parse後的布林值,
        /// 當文字不是數字型態時,
        /// 會回傳預設布林值為false,
        /// </summary>
        static public bool StringIsInt_Bool(string p)
        {
            int k;
            if (int.TryParse(p, out k))
            {
                return true;
            }
            else
            {
                return false;
            }
        }

        /// <summary>
        /// 傳入ComboBox,
        /// 檢查所輸入的文字,
        /// 是否為ItemList內的值,
        /// 當ComboBox Text為空時
        /// 視為資料正確
        /// </summary>
        static public bool ComboBoxValueInItemList(ComboBoxEx p)
        {
            bool b = false;

            if (!string.IsNullOrEmpty(p.Text.Trim()))
            {
                foreach (TeacherObj each in p.Items)
                {
                    if (each.TeacherFullName == p.Text.Trim())
                    {
                        b = true;
                        break;
                    }
                }
            }
            else
            {
                b = true;
            }

            return b;
        }

        /// <summary>
        /// 取得科別對照表文字清單
        /// (具排序效果)
        /// </summary>
        /// <returns></returns>
        static public List<string> GetQueryDeptList()
        {
            QueryHelper _QueryHelper = new QueryHelper();
            DataTable dtable = _QueryHelper.Select("select name from dept");
            List<string> list = new List<string>();
            foreach (DataRow row in dtable.Rows)
            {
                string name = "" + row[0];
                if (!string.IsNullOrEmpty(name))
                {
                    string[] namelist = name.Split(':');
                    string name1 = "" + namelist.GetValue(0);
                    if (!list.Contains(name1))
                    {
                        list.Add(name1);
                    }
                }
            }
            list.Sort();
            return list;
        }

        static public string GetDecimalValue(decimal DecValue, int IntValue)
        {
            string StringValue = "";

            StringValue = (DecValue * IntValue / 100).ToString();
            return StringValue;
        }

        /// <summary>
        /// 確認學生狀態是否正確,
        /// True:一般或延修生
        /// </summary>
        static public bool CheckStatus(StudentRecord student)
        {
            if (student.Status == StudentRecord.StudentStatus.一般 || student.Status == StudentRecord.StudentStatus.延修)
            {
                return true;
            }
            else
            {
                return false;
            }
        }

        /// <summary>
        /// 傳入學生RecordList 取得班級清單
        /// </summary>
        static public Dictionary<string, ClassRecord> GetClassDic(List<StudentRecord> studlist)
        {
            Dictionary<string, ClassRecord> dic = new Dictionary<string, ClassRecord>();
            List<string> list = new List<string>();
            foreach (StudentRecord srud in studlist)
            {
                if (!string.IsNullOrEmpty(srud.RefClassID))
                {
                    if (!list.Contains(srud.RefClassID))
                    {
                        list.Add(srud.RefClassID);
                    }
                }
            }
            List<ClassRecord> classlist = Class.SelectByIDs(list);
            foreach (ClassRecord each in classlist)
            {
                if (!dic.ContainsKey(each.ID))
                {
                    dic.Add(each.ID, each);
                }
            }
            return dic;
        }
    }
}
