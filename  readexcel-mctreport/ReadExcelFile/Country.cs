﻿using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Windows.Forms;

namespace ReadExcelFile
{
    class Country
    {
        private String name;

        private void extractCountryName(String name)
        {
            this.name = "NO COUNTRY NAME";

            // abort here if there is no name
            if (name.Trim().Length < 1)
                return;

            name = name.Trim().ToUpper();

            if (name.Contains("MID.") || name.Equals("GUATEMALA") || name.Equals("NICARAGUA"))
                this.name = "CENTRAL AMERICA";
            else if (name.Equals("UNIFIED"))
                this.name = "UNIFIED";
            else if (name.Contains("PT.") || name.Equals("PUERTORICO"))
                this.name = "PUERTO RICO";
            else if (name.Contains("(CHILE") || name.Contains("CHL("))
                this.name = "CHILE";
            else if (name.Equals("PAN(PANAMA)") || name.Contains("(PANAMA"))
                this.name = "PANAMA";
            else if (name.Contains("(CARIB_JAMAICA") || name.Contains("(JAMAICA"))
                this.name = "JAMAICA";
            else if (name.Contains("CRI(") || name.Contains("CRI (") || name.Equals("CRI(COSTA RICA)"))
                this.name = "COSTA RICA";
            else if (name.Contains("DOMENICA") || name.Contains("DOMINICAN"))
                this.name = "DOMINICA";
            else
                this.name = name;
        }
            

        public String Name
        {
            get { return this.name; }
            set
            {
                if (value == null)
                    extractCountryName("");
                else
                    extractCountryName(value.Trim().ToUpper());
            }

        }
    }
}
