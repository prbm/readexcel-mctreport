using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace ReadExcelFile
{
    class CountryCode
    {
        private String code;
        private String country;
        private String carrier;

        public String Code
        {
            get { return code; }
            set { code = value; }
        }

        public String Country
        {
            get { return country; }
            set { country = value; }
        }

        public String Carrier
        {
            get { return carrier; }
            set { carrier = value; }
        }
    }
}
