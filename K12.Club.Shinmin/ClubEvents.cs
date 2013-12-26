using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

//評量資訊更新

namespace K12.Club.Shinmin
{
    public static class ClubEvents
    {
        public static void RaiseAssnChanged()
        {
            if (ClubChanged != null)
                ClubChanged(null, EventArgs.Empty);
        }

        public static event EventHandler ClubChanged;
    }
}