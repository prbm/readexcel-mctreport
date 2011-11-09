using System;
using System.Collections.Generic;
using System.Collections;
using System.Linq;
using System.Text;

namespace GetCountryCarriers
{
    class CountryCarriers: IEnumerable, IEnumerator
    {
        private ArrayList CountryCarrierList = new ArrayList();
        private int position = -1;

        public void AddCountryCarrier(CountryCarrier countryCarrier)
        {
            CountryCarrierList.Add(countryCarrier); 
        }

        // Needed due to the usage of IEnumerable
        public IEnumerator GetEnumerator()
        {
            return (IEnumerator)this;
        }

        // Needed due to the usage of IEnumerable
        public bool MoveNext()
        {
            if (position < (CountryCarrierList.Count - 1))
            {
                ++position;
                return true;
            }
            return false;
        }

        public void Reset()
        {
            position = -1;
        }

        public object Current
        {
            get { return CountryCarrierList[position]; }
        }
    }
}
