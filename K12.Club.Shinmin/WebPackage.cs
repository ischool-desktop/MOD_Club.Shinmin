using System;
using System.Data;
using System.IO;
using System.Xml.Linq;
using System.Xml.XPath;
using FISCA.Data;

namespace K12.Club.Shinmin
{
    public class WebPackage
    {
        /// <summary>
        /// 註冊Gadget
        /// </summary>
        /// <param name="role"></param>
        /// <param name="deployPath"></param>
        /// <param name="description"></param>
        /// <returns></returns>
        public static Tuple<bool,string> RegisterGadget(string role,string deployPath,string description)
        {
            QueryHelper helper = new QueryHelper();

            string strWeb1Table = "$system.packagedefinition";
            string strWeb2Table = "_web_package";
            string strWebTable = string.Empty;            

            DataTable table = helper.Select("select * from " + strWeb2Table);

            //若_web_package有資料則是Web 2的設定，否則是Web 1
            if (table.Rows.Count > 0)
                strWebTable = strWeb2Table;
            else
            {
                table = helper.Select("select * from " + strWeb1Table);
                strWebTable = strWeb1Table;
            }

            DataRow Row = null;

            for (int i = 0; i < table.Rows.Count; i++)
            {
                string TargetType =  table.Rows[i].Field<string>("targettype");

                if (role.ToLower().Equals(TargetType.ToLower()))
                {
                    Row = table.Rows[i];
                    break;
                }
            }

            if (Row != null)
            {
                #region 建立Gadget
                string definition = Row.Field<string>("definition");
                string uid = Row.Field<string>("uid");

                StringReader reader = new StringReader(definition);

                XElement Element = XElement.Load(reader);



                XElement elmExist = Element
                    .Element("Gadgets")
                    .XPathSelectElement("./Gadget[@deployPath='" + deployPath + "']");

                if (elmExist == null)
                {
                    XElement elmGadget = new XElement("Gadget");
                    elmGadget.SetAttributeValue("deployPath", deployPath);
                    elmGadget.SetAttributeValue("description", description);

                    Element.Element("Gadgets").Add(elmGadget);
                }
                else
                    return new Tuple<bool,string>(false,"此Gadget已存在，不需新增！");
                #endregion

                #region 更新Gadget
                UpdateHelper updatehelper = new UpdateHelper();

                string strCommand =  "UPDATE " + strWebTable + " SET definition = '"+ Element +"' WHERE uid=" + uid;

                int result = updatehelper.Execute(strCommand);

                if (result == 0)
                    return  new Tuple<bool,string>(true,strWebTable.Equals(strWeb1Table) ? "成功新增Web 1 Gadget" : "成功新增Web 2 Gadget");
                else
                    return new Tuple<bool,string>(false,"更新資料庫失敗，錯誤代碼為" + result);
                #endregion
            }
            else
                return new Tuple<bool,string>(false,"指定的角色不存在！");
        }
    }
}