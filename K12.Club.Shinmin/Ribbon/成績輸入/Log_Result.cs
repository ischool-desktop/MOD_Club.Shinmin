using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using K12.Data;
using System.Xml;
using FISCA.DSAUtil;

namespace K12.Club.Shinmin
{
    class Log_Result
    {

        public StudentRecord _stud { get; set; }

        public ClassRecord _class { get; set; }
        /// <summary>
        /// 舊的資料
        /// </summary>
        public SCJoin _OldSch { get; set; }
        /// <summary>
        /// 修改內容
        /// </summary>
        public SCJoin _NewSch { get; set; }

        /// <summary>
        /// 
        /// </summary>
        public Log_Result(SCJoin sch)
        {
            _OldSch = sch;

            IsChange = false;
        }

        public bool IsChange { get; set; }

        /// <summary>
        /// 取得該筆Log調整內容
        /// </summary>
        public string GetLogString(成績取得器 GetPoint)
        {
            StringBuilder sb = new StringBuilder();

            string SeatNo = "";
            string CLUBName = "";
            if (_stud.SeatNo.HasValue)
                SeatNo = _stud.SeatNo.Value.ToString();

            if (GetPoint._ClubDic.ContainsKey(_OldSch.RefClubID))
            {
                CLUBName = GetPoint._ClubDic[_OldSch.RefClubID].ClubName;
            }

            sb.AppendLine(string.Format("社團「{0}」班級「{1}」座號「{2}」學生「{3}」", CLUBName, _class.Name, SeatNo, _stud.Name));
            if (GetString(_OldSch.PAScore) != GetString(_NewSch.PAScore))
            {
                sb.AppendLine(string.Format("「{0}」由「{1}」修改為「{2}」", "平時活動", GetString(_OldSch.PAScore), GetString(_NewSch.PAScore)));
            }
            if (GetString(_OldSch.ARScore) != GetString(_NewSch.ARScore))
            {
                sb.AppendLine(string.Format("「{0}」由「{1}」修改為「{2}」", "出缺席率", GetString(_OldSch.ARScore), GetString(_NewSch.ARScore)));
            }
            if (GetString(_OldSch.AASScore) != GetString(_NewSch.AASScore))
            {
                sb.AppendLine(string.Format("「{0}」由「{1}」修改為「{2}」", "活動力及服務", GetString(_OldSch.AASScore), GetString(_NewSch.AASScore)));
            }
            if (GetString(_OldSch.FARScore) != GetString(_NewSch.FARScore))
            {
                sb.AppendLine(string.Format("「{0}」由「{1}」修改為「{2}」", "成品成果考驗", GetString(_OldSch.FARScore), GetString(_NewSch.FARScore)));
            }

            return sb.ToString();
        }

        private string GetString(decimal? dec)
        {
            if (dec.HasValue)
            {
                return dec.Value.ToString();
            }
            else
            {
                return "";
            }
        }
    }
}
