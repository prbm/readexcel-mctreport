using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace ReadExcelFile
{
    class ProjectCAStatus
    {
        private int id;
        private Int32 nyear;
        private Int32 nmonth;
        private Int32 status;
        private DateTime dateOfChange;
        private Int32 projectID;

        public ProjectCAStatus()
        {
            id = 0;
            nyear = 0;
            nmonth = 0;
            status = 0;
            dateOfChange = new DateTime();
            projectID = 0;
        }

        public int Id
        {
            get { return id; }
            set { id = value; }
        }

        public Int32 Nyear
        {
            get { return nyear; }
            set { nyear = value; }
        }

        public Int32 Nmonth
        {
            get { return nmonth; }
            set { nmonth = value; }
        }

        public Int32 Status
        {
            get { return status; }
            set { status = value; }
        }

        public DateTime DateOfChange
        {
            get { return dateOfChange; }
            set { dateOfChange = value; }
        }

        public Int32 ProjectID
        {
            get { return projectID; }
            set { projectID = value; }
        }

    }
}
