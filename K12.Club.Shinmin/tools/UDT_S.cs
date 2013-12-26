using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace K12.Club.Shinmin
{
    static class UDT_S
    {
        /// <summary>
        /// 建立Sql字串
        /// </summary>
        /// <param name="GetColumn">所需要取得的資料名稱</param>
        /// <param name="StringList">資料組合使用之清單</param>
        /// <returns>回傳組成字元</returns>
        static public string PopOneCondition(string GetColumn ,List<string> StringList)
        {
            List<string> test = new List<string>();
            foreach (string each in StringList)
            {
                test.Add(string.Format("'{0}'", each));
            }
            string bulink = string.Format("{0} in(" + string.Join(",", test.ToArray()) + ")", GetColumn);
            return bulink;
        }

        /// <summary>
        /// 建立Sql字串,雙條件
        /// ref_student_id in ('99999') AND (NAME='周小小')
        /// </summary>
        /// <param name="GetColumn">所需要取得的資料名稱</param>
        /// <param name="StringList">資料組合使用之清單</param>
        /// <returns>回傳組成字元</returns>
        static public string PopTwoCondition(string GetColumn, string Two, List<string> StringList)
        {
            List<string> test = new List<string>();
            foreach (string each in StringList)
            {
                test.Add(string.Format("'{0}'", each));
            }
            string bulink = string.Format("{0} in(" + string.Join(",", test.ToArray()) + ")", GetColumn);
            bulink += string.Format("AND (Name ='{0}')", Two);
            return bulink;
        }

        /// <summary>
        /// 建立Sql字串,3條件
        /// ref_student_id in ('99999') AND (NAME='周小小')
        /// </summary>
        /// <param name="GetColumn">所需要取得的資料名稱</param>
        /// <param name="where">條件欄位名稱</param>
        /// <param name="Two">條件內容值</param>
        /// <param name="StringList">資料組合使用之清單</param>
        /// <returns>回傳組成字元</returns>
        static public string PopThereCondition(string GetColumn, string where, string Two, List<string> StringList)
        {
            List<string> test = new List<string>();
            foreach (string each in StringList)
            {
                test.Add(string.Format("'{0}'", each));
            }
            string bulink = string.Format("{0} in(" + string.Join(",", test.ToArray()) + ")", GetColumn);
            bulink += string.Format("AND ({0} ='{1}')",where, Two);
            return bulink;
        }

        /// <summary>
        /// 將DataTable內的DateTime欄位去除多餘字串,
        /// 2008/11/26 00:00:00.000 => 2008/11/26,
        /// (如果非時間值則回傳空字串)
        /// </summary>
        static public string ChangeTime(string Time)
        {
            DateTime dt;
            if (DateTime.TryParse(Time, out dt))
                return dt.ToShortDateString();
            else
                return "";
        }
    }
}
