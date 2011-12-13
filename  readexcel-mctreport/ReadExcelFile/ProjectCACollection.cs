using System;
using System.Collections.Generic;
using System.Collections;
using System.Linq;
using System.Text;

namespace ReadExcelFile
{
    class ProjectCACollection : CollectionBase, IEnumerable, IEnumerator
    {
        private int index = -1;

        public ProjectCACollection()
        {
            this.index = -1;
        }


        public void Add(ProjectCA pCA)
        {
            if( pCA!=null)
                throw new Exception ("The ProjectCA object informed was null");

            if (pCA.Id == 0 || pCA.ProjectCode == null || pCA.CarrierName == null || pCA.CountryName == null || pCA.Subsidiary == null || pCA.PdStatusProject == null)
                throw new Exception("The data members of ProjectCA informed CAN NOT be null or 0!");

            this.List.Add(pCA);
        }

        public void Remove(ProjectCA pCA)
        {
            this.List.Remove(pCA);
        }

        public ProjectCA this[int index]
        {
            get { return (ProjectCA)this.List[index]; }
            set { this.List[index] = value; }

        }

        #region IEnumerable Members
            IEnumerator IEnumerable.GetEnumerator()
            {
                return this;
            }
        #endregion

        #region IEnumerator Members

            public Object Current
            {
                get { return this.List[index]; }
            }

            public bool MoveNext()
            {
                this.index++;
                return (this.index<this.List.Count);
            }
            
            public void Reset()
            {
                this.index = -1;
            }

        #endregion

            //public void Save(ProjectCACollection collection)
            //{

            //}
    }
}
