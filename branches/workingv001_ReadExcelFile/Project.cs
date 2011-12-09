using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace ReadExcelFile
{
    class Project
    {
        String  name;
        Int32   quantity;
        List<CA> cas;
        List<String> countries;
        
        public Project()
        {
            cas = new List<CA>();
            countries = new List<String>();
        }

        public String Name
        {
            get
            {
                return name;
            }
            set
            {
                name = value;
            }
        }

        public List<CA> CAProjects
        {
            get
            {
                return cas;
            }
        }

        public Int32 Quantity
        {
            get
            {
                return quantity;
            }
            set
            {
                quantity = value;
            }
        }

        public String CAProject
        {
            set
            {
                CA ca = new CA();
                ca.CarrierName = value;

                // avoid adding already existing projects
                if (!cas.Contains(ca))
                    cas.Add(ca);

                // free the object
                ca = null;
            }
        }

        public String Country
        {
            set
            {
                if (!countries.Contains(value))
                {
                    countries.Add(value);
                }
            }
        }
    }


}
