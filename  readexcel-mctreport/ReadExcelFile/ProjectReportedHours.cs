using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace ReadExcelFile
{
    class ProjectReportedHours
    {
        private Int32 projectId;
        private Int32 month;
        private Int32 year;
        private Int32 employeeId;
        private Double amountReported;

        public Int32 ProjectId
        {
            get { return projectId; }
            set { projectId = value; }
        }

        public Int32 Month
        {
            get { return month; }
            set { month = value; }
        }

        public Int32 Year
        {
            get { return year; }
            set { year = value; }
        }

        public Int32 EmployeeId
        {
            get { return employeeId; }
            set { employeeId = value; }
        }

        public Double AmountReported
        {
            get { return amountReported; }
            set { amountReported = value; }
        }
    }
}
