using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace ReadExcelFile
{
    class Carrier
    {
        private String name;

        private void extractCarrierName(String name)
        {
            this.name = "NO CARRIER NAME";
            String[] tmp;

            // includes test variable, not consider in future releases
            String blablabla = "bla";

            // if there is no content, abort it
            if (name.Trim().Length < 1)
                return;

            // remove - character
            if (name.Contains('_'))
            {
                tmp = name.Split('_');
                name = tmp[0];
            }

            name = name.Trim().ToUpper();

            // get the final name of carrier
            if (name.Equals("NO CARRIER NAME"))
                this.name = "OPEN";
            else if (name.Equals("MOVISTAR") || name.Equals("TELEFONICA(COSTA RICA)") || name.Equals("TELEFONICA_MOVILES"))
                this.name = "TELEFONICA";
            else if (name.Equals("VTR(CHILE)"))
                this.name = "VTR";
            else if (name.Equals("VIRGIN MOBILE(CHILE)"))
                //this.name = "VIRGIN MOBILE";
                this.name = "VIRGIN";
            //else if (name.Equals("CT MIAMI(PANAMA)"))
            //    this.name = "CT MIAMI";
            else if (name.Equals("MIAMI OPEN(PANAMA)") || name.Equals("MIAMI OPEN") || name.Equals("CT MIAMI(PANAMA)") || name.Equals("CTM") || name.Equals("CTMIAMI"))
                this.name = "CT MIAMI";
            //else if (name.Equals("MIAMI OPEN(PANAMA)") || name.Equals("CT MIAMI(PANAMA)"))
            //    this.name = "MIAMI OPEN";
            else if (name.Equals("VIVA(DOMINICA)"))
                this.name = "VIVA";
            else if (name.Equals("CLARO(JAMAICA)") || name.Equals("CLARO(COSTA RICA)"))
                this.name = "CLARO";
            else if (name.Equals("OPEN MARKET"))
                this.name = "OPEN";
            else if (name.Equals("ALEGRO_PCS"))
                this.name = "ALEGRO";
            else if (name.Equals("NEXTEL"))
                this.name = "NEXTEL";
            else if (name.Equals("ICE(COSTA RICA)") || name.Equals("ICE(OPEN)"))
                this.name = "ICE";
            else if (name.Equals("C&W") || name.Contains("CNW"))
                this.name = "CNW";
            else if (name.Equals("USACELL"))
                this.name = "IUSACELL";
            else if (name.Equals("OLA"))
                this.name = "TIGO";
            else
                this.name = name;
        }

        public String Name
        {
            get { return this.name; }
            set { 
                if(value==null)
                    extractCarrierName(""); 
                else
                    extractCarrierName(value.Trim().ToUpper().ToString()); 
            }
        }

    }
}
