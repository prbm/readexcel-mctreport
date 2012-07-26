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
        private String prmsCode;
        private String subsidiary;
        private String projectStatus;
        private Double numberWorkingPeople;
        private Double totalManMonth;
        private Double mediumManMonth;
        private Double totalWorkingHours;
        private Double mediumWorkingHours;
        private Int32 peopleReportedHours;
        private List<CountryCode> listCountryCodes;
        private List<ReportHours> reportedHours = new List<ReportHours>();

        public CA(){}

        public CA(List<CountryCode> listCountryCodes)
        {
            this.listCountryCodes = listCountryCodes;
        }

        public String CarrierName
        {
            get { return this.carrierName; }
            set {
                if (value == null)
                    extractCarrierName("");
                else
                    extractCarrierName(value.ToString()); }
        }

        private void extractCarrierName(String txt)
        {
            this.carrierName = "NO CARRIER NAME";

            // if there is no text, return here
            if (txt.Trim().Length < 1)
                return;
            else
            {
                // if not, get the carrier name
                Carrier c = new Carrier();
                c.Name = txt.Trim().ToUpper();
                this.carrierName = c.Name;
            }
        }

        private void extractCountryName(String txt)
        {
            this.country = "NO COUNTRY NAME";
            Country c = new Country();

            if (txt.Trim().Length < 1)
                c.Name = "";
            else
                c.Name = txt;

            this.country = c.Name;

            CountryCode cc = new CountryCode();
            cc = listCountryCodes.Find(delegate(CountryCode country) { return (country.Country.Equals(this.country)); });
            if (cc==null)
            {
                if (this.country.Equals("NO COUNTRY NAME"))
                    this.country = "VENEZUELA";
                else
                    this.country = c.Name;
            }
        }

        public String Country
        {
            get { return this.country; }
            set {
                if (value == null)
                    extractCountryName("");
                else
                    extractCountryName(value.ToString());
            }
        }

        public void setProjectStatus(String oriStatus)
        {
           this.projectStatus = "EMPTY STATUS";

            if (oriStatus != null)
            {
                oriStatus = oriStatus.Trim().ToUpper();

                if (oriStatus.Equals("COMPLETED") || oriStatus.Equals("OS UPGRADE"))
                    this.projectStatus = "COMPLETED";
                else if (oriStatus.Equals("DROPPED") || oriStatus.Equals("DROP"))
                    this.projectStatus = "DROPPED";
                else if (oriStatus.Equals("HOLD"))
                    this.projectStatus = "HOLD";
                else if (oriStatus.Equals("RC"))
                    this.projectStatus = "ECO";
                else if (oriStatus.Equals("WAIT"))
                    this.projectStatus = "WAIT";
                else if (oriStatus.Equals("RUNNING"))
                    this.projectStatus = "RUNNING";
            }

        }

        public String ProjectStatus
        {
            get { return this.projectStatus; }
            set 
            {
                    setProjectStatus(value);
            }
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

        public void setSubsidiary(String country){
            this.subsidiary = null;

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
                this.subsidiary = "LGESP";
            else if(country.ToUpper().Equals("ARGENTINA") || country.ToUpper().Equals("URUGUAY"))
                this.subsidiary = "LGEAR";
            else if(country.ToUpper().Equals("MEXICO"))
                this.subsidiary = "LGEMS";
            else if(country.ToUpper().Equals("CHILE") || country.ToUpper().Equals("BOLIVIA"))
                this.subsidiary = "LGECL";
            else if(country.ToUpper().Equals("COLOMBIA"))
                this.subsidiary = "LGECB";
            else if(country.ToUpper().Equals("PERU"))
                this.subsidiary = "LGEPE";
            else
                this.subsidiary = "SUB NOT DECLARED";
        }

        public String Subsidiary
        {
            get { return this.subsidiary; }
            set { this.subsidiary = value.Trim().ToString(); }
        }

        public String PRMSCode
        {
            get { return this.prmsCode; }
            set { this.prmsCode = value.Trim().ToString(); }
        }

        public List<ReportHours> ListReportedHours
        {
            get { return reportedHours; }
            set { reportedHours = value; }
        }

        public void setRepHour(double time, String name)
        {
            ReportHours rp = new ReportHours();
            rp.TeamName = name;
            rp.ReportedTime = time;

            if (reportedHours == null)
                reportedHours = new List<ReportHours>();

            if (reportedHours.Find(delegate(ReportHours rep) { return rep.TeamName.ToUpper().Equals(rp.TeamName.ToUpper()); }) != null)
                reportedHours.Remove(reportedHours.Find(delegate(ReportHours rep) { return rep.TeamName.ToUpper().Equals(rp.TeamName.ToUpper()); }));

            reportedHours.Add(rp);

        }
    }
}
