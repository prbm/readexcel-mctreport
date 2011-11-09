using System;
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
            set { this.carrierName = value.ToString(); }
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
                if (tmp.Equals("NO COUNTRY NAME"))
                    tmp = "VENEZUELA";
                else if (tmp.Contains("MID.") || tmp.Equals("UNIFIED"))
                    tmp = "CENTRAL AMERICA";
                else if (tmp.Contains("PT."))
                    tmp = "PUERTO RICO";
                else if (tmp.Contains("CRI("))
                    tmp = "COSTA RICA";
                else if (tmp.Contains("DOMENICA"))
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

    }
}
