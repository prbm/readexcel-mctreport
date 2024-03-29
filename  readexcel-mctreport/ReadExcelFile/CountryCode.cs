﻿using System;
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
            set { 
                // Normalize the country name
                Country c = new Country();
                c.Name = value;
                this.country = c.Name;
            }
        }

        public String Carrier
        {
            get { return carrier; }
            set {
                // normaliza the carrier Name;
                Carrier c = new Carrier();
                c.Name = value;
                carrier = c.Name; 
            }
        }
    }
}
