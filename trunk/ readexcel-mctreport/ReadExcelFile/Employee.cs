using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace ReadExcelFile
{
    class Employee
    {
        private Int32 id;
        private String name;
        private Int32 companyId;
        private DateTime date;
        private DateTime arrival;
        private DateTime left;
        private TimeSpan workingHours;

        public Int32 Id
        {
            get { return id; }
            set { id = value; }
        }

        public String Name
        {
            get{return name;}
            set{name = value;}
        }

        public Int32 CompanyId
        {
            get { return companyId; }
            set { companyId = value; }
        }

        public DateTime Arrival
        {
            get { return arrival; }
            set { arrival = value; }
        }
        public DateTime Left
        {
            get { return left; }
            set { 
                left = value;
                workingHours = left.Subtract(arrival.AddMinutes(60));
            }
        }
        public DateTime Date
        {
            get { return date; }
            set { date = value; }
        }
        public TimeSpan WorkingHours
        {
            get { return workingHours; }
        }

    }
}
