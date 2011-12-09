using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace ReadExcelFile
{
    class Employee
    {
        String name;
        DateTime date;
        DateTime arrival;
        DateTime left;
        TimeSpan workingHours;


        public String Name
        {
            get{return name;}
            set{name = value;}
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
