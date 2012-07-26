using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace ReadExcelFile
{
    class ProjectCA 
    {
        private Int32 id;
        private String projectCode;
        private String carrierName;
        private String countryName;
        private Subsidiary subsidiary;
        private Int32 pdStatusProject;
        private PMSStatus pmsStatus;
        private ProjectCAStatus pcaStatus;
        private List<ProjectReportedHours> pRH;
        internal List<ProjectReportedHours> PRH
        {
            get { return pRH; }
        }

        public ProjectCA()
        {
            id = 0;
            projectCode = null;
            carrierName = null;
            countryName = null;
            subsidiary = new Subsidiary();
            pdStatusProject = 0;
            pcaStatus = new ProjectCAStatus();
            pmsStatus = new PMSStatus();
        }

        public String CarrierName
        {
            get { return carrierName; }
            set { carrierName = value; }
        }

        public String CountryName
        {
            get { return countryName; }
            set { countryName = value; }
        }

        public Subsidiary Subsidiary
        {
            get { return subsidiary; }
            set { subsidiary = value; }
        }

        public Int32 Id
        {
            get { return id; }
            set { id = value; }
        }

        public String ProjectCode
        {
            get { return projectCode; }
            set { projectCode = value; }
        }

        public Int32 PdStatusProject
        {
            get { return pdStatusProject; }
            set { pdStatusProject = value; }
        }


        public PMSStatus PmsStatus
        {
            get { return pmsStatus; }
            set { pmsStatus = value; }
        }


        internal ProjectCAStatus PcaStatus
        {
            get { return pcaStatus; }
            set { pcaStatus = value; }
        }
    }
    
}
