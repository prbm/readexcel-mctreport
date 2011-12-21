using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace ReadExcelFile
{
    class ProjectCourseCarrier
    {
        private Int32 idCarrier;
        private Int32 idCountry;
        private String name;

        public Int32 IdCarrier
        {
            get { return idCarrier; }
            set { idCarrier = value; }
        }

        public Int32 IdCountry
        {
            get { return idCountry; }
            set { idCountry = value; }
        }

        public String Name
        {
            get { return name; }
            set { name = value; }
        }
    }
}
