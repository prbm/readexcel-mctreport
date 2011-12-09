using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace ReadExcelFile
{
    class Model
    {
        private String modelCode;
        private CA modelCA;

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

        public CA ModelCA
        {
            get
            {
                return this.modelCA;
            }
            set
            {
                if (value.GetType() == typeof(CA))
                {
                    modelCA = new CA(value.ListCountryCodes);
                    modelCA.CarrierName = value.CarrierName;
                    modelCA.Country = value.Country;
                    modelCA.MediumManMonth = value.MediumManMonth;
                    modelCA.ProjectStatus = value.ProjectStatus;
                    modelCA.PeopleReportedHours = value.PeopleReportedHours;
                    modelCA.setSubsidiary(modelCA.Country);
                }

            }
        }

    }
}
