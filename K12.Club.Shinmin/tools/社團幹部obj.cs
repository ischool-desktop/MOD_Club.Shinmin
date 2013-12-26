using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using K12.Data;

namespace K12.Club.Shinmin
{
    class 社團幹部obj
    {

        //社團幹部 - 社長/副社長
        public Dictionary<string, string> _Cadre1 { get; set; }

        public CLUBRecord _Club { get; set; }


        public 社團幹部obj(CLUBRecord Club)
        {
            _Club = Club;
            _Cadre1 = new Dictionary<string, string>();

            if (!string.IsNullOrEmpty(_Club.President))
            {
                _Cadre1.Add(_Club.President, "社長");
            }

            if (!string.IsNullOrEmpty(_Club.VicePresident))
            {
                if (_Cadre1.ContainsKey(_Club.VicePresident))
                {
                    _Cadre1[_Club.VicePresident] += ",副社長";
                }
                else
                {
                    _Cadre1.Add(_Club.VicePresident, "副社長");
                }
            }

        }
    }
}
