﻿using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Windows.Forms;

namespace ReadExcelFile
{
    class CA
    {
        private String carrierName;
        private String country;
        private String subsidiary;
        private String projectStatus;
        private Double numberWorkingPeople;
        private Double totalManMonth;
        private Double mediumManMonth;
        private Double totalWorkingHours;
        private Double mediumWorkingHours;
        private Int32 peopleReportedHours;
        private List<CountryCode> listCountryCodes;

        public CA(){}

        public CA(List<CountryCode> listCountryCodes)
        {
            this.listCountryCodes = listCountryCodes;
        }

        public String CarrierName
        {
            get { return this.carrierName; }
            set { this.carrierName = extractCarrierName(value.ToString()); }
        }

        private String extractCarrierName(String txt)
        {
            String tmp = "NO CARRIER NAME";

            if (txt == null)
                return tmp;

            if (txt.ToUpper().Equals("NO CARRIER NAME"))
                tmp = "OPEN";
            else if (txt.ToUpper().Equals("MOVISTAR"))
                tmp = "TELEFONICA";
            else if (txt.ToUpper().Equals("OPEN MARKET") || txt.ToUpper().Equals("ICE(OPEN)"))
                tmp = "OPEN";
            else if (txt.ToUpper().Equals("ALEGRO_PCS"))
                tmp = "ALEGRO";
            else if (txt.ToUpper().Equals("ICE(COSTA RICA)"))
                tmp = "ICE";
            else if (txt.ToUpper().Equals("C&W"))
                tmp = "CNW";
            else
                tmp = txt.ToUpper().ToString();

            return tmp;
        }

        private String extractCountryName(String txt)
        {
            String tmp = "NO COUNTRY NAME";

            if (txt != null)
                tmp = txt;
            else
                return tmp;

            CountryCode cc = new CountryCode();
            cc = listCountryCodes.Find(delegate(CountryCode c) { return (c.Country.Equals(tmp)); });
            if (cc==null)
            {
                if (tmp.ToUpper().Equals("NO COUNTRY NAME"))
                    tmp = "VENEZUELA";
                else if (tmp.ToUpper().Contains("MID."))
                    tmp = "CENTRAL AMERICA";
                else if (tmp.ToUpper().Equals("UNIFIED"))
                    tmp = "UNIFIED";
                else if (tmp.ToUpper().Contains("PT."))
                    tmp = "PUERTO RICO";
                else if (tmp.ToUpper().Contains("CRI("))
                    tmp = "COSTA RICA";
                else if (tmp.ToUpper().Contains("DOMENICA"))
                    tmp = "DOMINICA";
            }

            return tmp;
        }

        public String Country
        {
            get { return this.country; }
            set { this.country = extractCountryName(value.ToString());}
        }

        public String ProjectStatus
        {
            get { return this.projectStatus; }
            set { this.projectStatus = value; }
        }
        public Double NumberWorkingPeople
        {
            get { return this.numberWorkingPeople; }
            set { this.numberWorkingPeople = Double.Parse(value.ToString()); }
        }

        public Double TotalManMonth
        {
            get { return this.totalManMonth; }
            set { this.totalManMonth = Double.Parse(value.ToString()); }
        }

        public Double MediumManMonth
        {
            get { return this.mediumManMonth; }
            set { this.mediumManMonth = Double.Parse(value.ToString()); }
        }

        public Double TotalWorkingHours
        {
            get { return this.totalWorkingHours; }
            set { this.totalWorkingHours = Double.Parse(value.ToString()); }
        }

        public Double MediumWorkingHours
        {
            get { return this.mediumWorkingHours; }
            set { this.mediumWorkingHours = Double.Parse(value.ToString()); }
        }

        public List<CountryCode> ListCountryCodes
        {
            get { return this.listCountryCodes; }
            set { this.listCountryCodes = value; }
        }

        public Int32 PeopleReportedHours
        {
            get { return this.peopleReportedHours; }
            set { this.peopleReportedHours = value; }
        }

        private String getSubsidiary(String country){
            String subsidiary = null;

            if (country.ToUpper().Equals("CENTRAL AMERICA") || 
               country.ToUpper().Equals("COSTA RICA") ||
               country.ToUpper().Equals("CUBA") ||
               country.ToUpper().Equals("DOMINICA") ||
               country.ToUpper().Equals("ECUADOR") ||
               country.ToUpper().Equals("JAMAICA") ||
               country.ToUpper().Equals("PANAMA") ||
               country.ToUpper().Equals("PARAGUAY") ||
               country.ToUpper().Equals("PUERTO RICO") ||
               country.ToUpper().Equals("VENEZUELA"))
                subsidiary = "LGEPS";
            else if(country.ToUpper().Equals("BRAZIL"))
                subsidiary = "LGESP";
            else if(country.ToUpper().Equals("ARGENTINA") || country.ToUpper().Equals("URUGUAY"))
                subsidiary = "LGEAR";
            else if(country.ToUpper().Equals("MEXICO"))
                subsidiary = "LGEMS";
            else if(country.ToUpper().Equals("CHILE") || country.ToUpper().Equals("BOLIVIA"))
                subsidiary = "LGECL";
            else if(country.ToUpper().Equals("COLOMBIA"))
                subsidiary = "LGECB";
            else if(country.ToUpper().Equals("PERU"))
                subsidiary = "LGEPE";
            else
                subsidiary = "SUB NOT DECLARED";

            return subsidiary;
        }

        public String Subsidiary
        {
            get { return this.subsidiary; }
            set { this.subsidiary = getSubsidiary(value.Trim().ToString()); }
        }

    }
}
