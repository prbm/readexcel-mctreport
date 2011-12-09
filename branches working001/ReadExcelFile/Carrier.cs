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

            // if there is no content, abort it
            if (name.Trim().Length < 1)
                return;

            // remove - character
            if (name.Contains('_'))
            {
                tmp = name.Split('_');
                name = tmp[0];
            }

            // get the final name of carrier
            if (name.Equals("NO CARRIER NAME"))
                this.name = "OPEN";
            else if (name.Equals("MOVISTAR") || name.Equals("TELEFONICA(COSTA RICA)"))
                this.name = "TELEFONICA";
            else if (name.Equals("VTR(CHILE)"))
                this.name = "VTR";
            else if (name.Equals("CT MIAMI(PANAMA)"))
                this.name = "CT MIAMI";
            else if (name.Equals("VIVA(DOMINICA)"))
                this.name = "VIVA";
            else if (name.Equals("CLARO(JAMAICA)"))
                this.name = "CLARO";
            else if (name.Equals("OPEN MARKET"))
                this.name = "OPEN";
            else if (name.Equals("ALEGRO_PCS"))
                this.name = "ALEGRO";
            else if (name.Equals("ICE(COSTA RICA)") || name.Equals("ICE(OPEN)"))
                this.name = "ICE";
            else if (name.Equals("C&W") || name.Contains("CNW"))
                this.name = "CNW";
            else
                this.name = name.Trim().ToUpper();

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
