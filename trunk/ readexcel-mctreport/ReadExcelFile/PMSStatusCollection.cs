using System;
using System.Collections.Generic;
using System.Collections;
using System.Linq;
using System.Text;

namespace ReadExcelFile
{
    class PMSStatusCollection : CollectionBase, IEnumerable, IEnumerator
    {
        private int index = -1;

        public void Add(PMSStatus pmsSt)
        {
            if (pmsSt == null)
                throw new Exception("The PMSStatus object informed was null");

            if (pmsSt.ID < 0 || pmsSt.Code == null || pmsSt.Description == null || pmsSt.CreationDate == null)
                throw new Exception("The data members of PMSStatus informed CAN NOT be null or 0!");

            this.List.Add(pmsSt);
        }

        public void Remove(PMSStatus pmsSt)
        {
            this.List.Remove(pmsSt);
        }

        public PMSStatus this[int index]
        {
            get { return (PMSStatus)this.List[index]; }
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
            return (this.index < this.List.Count);
        }

        public void Reset()
        {
            this.index = -1;
        }

        #endregion
    }
}
