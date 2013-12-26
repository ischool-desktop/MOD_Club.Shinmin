using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace K12.Club.Shinmin
{
    class StudentJoinRow : System.ComponentModel.INotifyPropertyChanged
    {
        //年級
        //班級
        //座號
        //學號
        //性別
        //姓名
        //List<SCJoin> - 學生修課記錄
        //SCJoin - 修改後的儲存記錄
        //原社團ClubRecord
        //原社團名稱
        //新社團ClubRecord
        //新社團名稱

        //方法1 - 請調整修課(傳入課程ID)
        //方法2 - 告訴我要儲存的

        public StudentJoinRow()
        {
            SCJoinList = new List<SCJoin>();
            CurrentClubRecord = null;
            AdjustClubRecord = null;

            GradeYear = "";
            ClassName = "";
            SeatNo = "";
        }



        /// <summary>
        /// 學生系統編號
        /// </summary>
        public string Ref_Student_Id { get; set; }

        /// <summary>
        /// 年級
        /// </summary>
        public string GradeYear { get; set; }

        /// <summary>
        /// 班級
        /// </summary>
        public string ClassName { get; set; }

        /// <summary>
        /// 座號
        /// </summary>
        public string SeatNo { get; set; }

        /// <summary>
        /// 學號
        /// </summary>
        public string StudentNumber { get; set; }

        /// <summary>
        /// 性別
        /// </summary>
        public string Gender { get; set; }

        /// <summary>
        /// 姓名
        /// </summary>
        public string Name { get; set; }

        /// <summary>
        /// 學生社團參與清單
        /// </summary>
        public List<SCJoin> SCJoinList { get; set; }

        /// <summary>
        /// 是否重複參與社團
        /// </summary>
        public bool DuplicateJoin
        {
            get
            {
                return SCJoinList.Count > 1;
            }
        }

        /// <summary>
        /// 原社團名稱
        /// </summary>
        public string CurrentClubName
        {
            get
            {
                if (SCJoinList.Count > 1)
                    return "重複參與社團";

                if (CurrentClubRecord != null)
                    return CurrentClubRecord.ClubName;
                else
                    return string.Empty;
            }
        }

        /// <summary>
        /// 原社團Record
        /// </summary>
        public CLUBRecord CurrentClubRecord { get; set; }

        /// <summary>
        /// 調整後社團名稱
        /// </summary>
        public string AdjustClubName
        {
            get
            {
                if (SCJoinList.Count > 1)
                    return "重複參與社團";

                CLUBRecord cr = AdjustClubRecord == null ? CurrentClubRecord : AdjustClubRecord;
                if (cr != null)
                    return cr.ClubName;
                else
                    return string.Empty;
            }
        }

        /// <summary>
        /// 調整後社團Record
        /// </summary>
        public CLUBRecord AdjustClubRecord { get; set; }

        /// <summary>
        /// 取得調整後社團參與記錄
        /// </summary>
        /// <returns>回傳社團參與清單第一筆記錄</returns>
        private SCJoin GetSCJoin()
        {
            if (SCJoinList.Count > 0)
                return SCJoinList[0];
            else
            {
                throw new InvalidOperationException("沒有任何社團參與記錄!");
            }
        }

        /// <summary>
        /// 調整參與社團為newCLUB
        /// </summary>
        /// <param name="newCLUB"></param>
        public void ChangeCLUB(CLUBRecord newCLUB)
        {
            AdjustClubRecord = newCLUB;
            if (PropertyChanged != null)
                PropertyChanged(this, new System.ComponentModel.PropertyChangedEventArgs("AdjustClubName"));
        }

        /// <summary>
        /// 是否變更參與社團
        /// </summary>
        public bool HasChanged
        {
            get
            {
                if (AdjustClubRecord == CurrentClubRecord)
                {
                    return false;
                }
                else if (AdjustClubRecord == null)
                {
                    return false;
                }
                else
                {
                    return true;
                }
            }
        }

        /// <summary>
        /// 取得調整後的"社團參與記錄"
        /// 如果未調整,則回傳null
        /// </summary>
        public SCJoin GetChange()
        {
            if (HasChanged)
            {
                SCJoin scj = GetSCJoin();
                scj.RefClubID = AdjustClubRecord.UID;
                return scj;
            }
            else
            {
                return null;
            }


        }

        public event System.ComponentModel.PropertyChangedEventHandler PropertyChanged;
    }
}
