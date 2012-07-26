using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace ReadExcelFile
{
    class ReportHours
    {
        private String teamName = null;
        private double reportedTime = 0.0;

        public String TeamName
        {
            get { return teamName; }
            set { teamName = value; }
        }

        public double ReportedTime
        {
            get { return reportedTime; }
            set { reportedTime = value; }
        }


    }


}
