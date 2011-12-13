using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace ReadExcelFile
{
    class Subsidiary
    {
        private Int32 id;
        private String name;
        private String description;

        public Int32 ID
        {
            get { return this.id; }
            set { this.id = value; }
        }

        public String Name
        {
            get { return this.name; }
            set { this.name = value; }
        }

        public String Description
        {
            get { return this.description; }
            set { this.description = value; }
        }

        public String getSubsidiaryCode(String country)
        {
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
            else if (country.ToUpper().Equals("BRAZIL"))
                subsidiary = "LGESP";
            else if (country.ToUpper().Equals("ARGENTINA") || country.ToUpper().Equals("URUGUAY"))
                subsidiary = "LGEAR";
            else if (country.ToUpper().Equals("MEXICO"))
                subsidiary = "LGEMS";
            else if (country.ToUpper().Equals("CHILE") || country.ToUpper().Equals("BOLIVIA"))
                subsidiary = "LGECL";
            else if (country.ToUpper().Equals("COLOMBIA"))
                subsidiary = "LGECB";
            else if (country.ToUpper().Equals("PERU"))
                subsidiary = "LGEPE";
            else
               subsidiary = "SUB NOT DECLARED";

            return subsidiary;
        }
    }
}
