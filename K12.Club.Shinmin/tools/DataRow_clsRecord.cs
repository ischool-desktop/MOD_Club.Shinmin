using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Data;

namespace K12.Club.Shinmin
{
    class DataRow_clsRecord
    {
        public string id { get; set; }
        public string classname { get; set; }
        public string gradeyear { get; set; }

        public DataRow_clsRecord(DataRow row)
        {
            id = "" + row[0];
            classname = "" + row[1];
            gradeyear = "" + row[2];
        }
    }
}
