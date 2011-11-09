using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace GetCountryCarriers
{
    class CountryCarrier
    {
        private String code;
        private String country;
        private String carrier;

        public CountryCarrier(String code)
        {
            this.Code = code;
        }

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
