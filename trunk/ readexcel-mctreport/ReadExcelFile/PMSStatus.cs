using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace ReadExcelFile
{
    class PMSStatus
    {
        private Int32 id;
        private String code;
        private String description;
        private DateTime creationdate;

        public PMSStatus()
        {
            id = 0;
            code = null;
            description = null;
            creationdate = new DateTime();
        }

        public Int32 ID
        {
            get { return this.id; }
            set { this.id = value; }
        }

        public String Code
        {
            get { return this.code; }
            set { this.code = value; }
        }

        public String Description
        {
            get { return this.description; }
            set { this.description = value; }
        }

        public DateTime CreationDate
        {
            get { return this.creationdate; }
            set { this.creationdate = value; }
        }
    }


}
