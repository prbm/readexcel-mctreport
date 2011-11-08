using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace ReadExcelFile
{
    class Model
    {
        private String modelCode;
        private List<CA> modelCAs;
        private CA firstCA;
        private Int32 numberModelCAs;

        public Model()
        {
            this.numberModelCAs = 0;
            modelCAs = new List<CA>();
        }

        public String extractModelName(String txt)
        {
            String tmp = "NO MODEL NAME";

            if (txt != null)
                tmp = txt;
            
            // remove undesirable characters in the code portions below
            // remove - character
            if (tmp.Contains("-"))
            {
                String[] t = tmp.Split('-');
                tmp = t[1];
            }

            // remove LG starting characters
            if (tmp.Contains("LG"))
                tmp = tmp.Substring(txt.IndexOf("L") + 2, tmp.Length - 2);

            // remove _ character
            if (tmp.Contains("_"))
                tmp = tmp.Substring(0, tmp.IndexOf("_"));

            // remove characters at the end of the model name
            if (tmp.Length > 4 && !Char.IsNumber(tmp, tmp.Length - 1))
            {
                char[] t = tmp.ToCharArray();
                int c = tmp.Length - 1;

                for (; c >= 0; c--)
                    if (Char.IsNumber(t[c]))
                        break;

                tmp = tmp.Substring(0, ++c);
            }

            if (tmp.Length < 1)
                return "NO MODEL NAME";
            else
                return tmp;
        }

        public String ModelCode
        {
            get { return this.modelCode; }
            set { this.modelCode = extractModelName(value.ToString());}
        }

        public List<CA> ModelCas
        {
            get { return this.modelCAs; }
        }

        public CA ModelCAs
        {
            get
            {
                firstCA = new CA();
                foreach (CA ca in modelCAs)
                {
                    firstCA.ListCountryCodes = ca.ListCountryCodes;
                    firstCA.CarrierName = ca.CarrierName;
                    firstCA.Country = ca.Country;
                    firstCA.MediumManMonth = ca.MediumManMonth;
                    firstCA.ProjectStatus = ca.ProjectStatus;
                    firstCA.PeopleReportedHours = ca.PeopleReportedHours;
                }
                return firstCA;
            }
            set
            {
                if (value.GetType() == typeof(CA))
                {
                    CA ca = new CA(value.ListCountryCodes);

                    ca.CarrierName = value.CarrierName;
                    ca.Country = value.Country;
                    ca.MediumManMonth = value.MediumManMonth;
                    ca.ProjectStatus = value.ProjectStatus;
                    ca.PeopleReportedHours = value.PeopleReportedHours;

                    modelCAs.Add(ca);
                }

            }
        }

    }
}
