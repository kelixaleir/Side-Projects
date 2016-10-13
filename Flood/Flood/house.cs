using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Flood
{
    class house
    {
        private string _band;
        private string _name;
        private string _number;
        private string _ss1;
        private string _street;
        private string _town;
        private string _postcode;
        private string _pre;

        public house(string band, string name, string number, string ss1, string street, string town, string postcode, string pre)
        {
            _band = band;
            _name = name;
            _ss1 = ss1;
            _number = number;
            _street = street;
            _town = town;
            _pre = pre;
            _postcode = postcode;

        }

        public string Band
        {
            get
            {
                return _band;
            }

            set
            {
                _band = value;
            }
        }
        public string Pre
        {
            get
            {
                return _pre;
            }

            set
            {
                _pre = value;
            }
        }



        public string Name
        {
            get
            {
                return _name;
            }

            set
            {
                _name = value;
            }
        }
        public string Number
        {
            get
            {
                return _number;
            }

            set
            {
                _number = value;
            }
        }
        public string Street
        {
            get
            {
                return _street;
            }

            set
            {
                _street = value;
            }
        }

        public string Postcode
        {
            get
            {
                return _postcode;
            }

            set
            {
                _postcode = value;
            }
        }
        public string Town
        {
            get
            {
                return _town;
            }

            set
            {
                _town = value;
            }
        }
        public string SS1
        {
            get
            {
                return _ss1;
            }

            set
            {
                _ss1 = value;
            }
        }
    }
}
