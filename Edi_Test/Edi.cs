using System;

namespace Edi_Test
{
    public class Edi
    {
        private string _policynum;
        private string _policynum2;
        private string _poltype;
        private string _messagetype;
        private string _inceptiondate;
        private string _enddate;

        private string _endcode;

        private string _custgender;
        private string _addcustgender;
        private string _alarm;

        private string _firstline;
        private string _secondline;
        private string _thirdline;
        private string _fourthline;
        private string _fifthline;
        private string _sixthline;
        private string _postcode;

        private string _addfirstline;
        private string _addsecondline;
        private string _addthirdline;
        private string _addfourthline;
        private string _addfifthline;
        private string _addsixthline;
        private string _addpostcode;

        private string _riskfirstline;
        private string _risksecondline;
        private string _riskthirdline;
        private string _riskfourthline;
        private string _riskfifthline;
        private string _risksixthline;
        private string _risk_postcode;

        private string _custtitle;
        private string _firstname;
        private string _lastname;
        private string _occupation;
        private string _maritialstatus;
        private string _employeesbusiness;
        private string _employmentcatagory;
        private string _dob;

        private bool _addpolicyholder;
        private bool _concover;
        private bool _buscover;

        private bool _outofhome;
        private bool _inthehome;

        private bool endorsement;

        private string _addcusttitle;
        private string _addfirstname;
        private string _addlastname;
        private string _addoccupation;
        private string _addmaritialstatus;
        private string _addemployeesbusiness;
        private string _addemploymentcatagory;
        private string _adddob;
        private string _relationship;

        private string _walls;
        private string _roof;
        private string _type;
        private string _roofpercent;
        private string _buildtype;
        private string _listed;
        private string _noofbed;
        private string _datebuild;
        private string _movedate;
        private string _businessuse;
        private string _ownership;
        private string _occupancystatus;

        private string _noofchildren;
        private string _noofadults;

        private string _buildcomp;
        private string _buildratedyear;
        private string _buildvol;
        private string _buildpremium;
        private string _buildsum;
        private string _buildad;
        private string _premium;

        private string _contcomp;
        private string _contratedyear;
        private string _contvol;
        private string _contpremium;
        private string _contsum;
        private string _contad;

        private string inhomecount;
        private string outhomecount;

        private string pedalcount;
        private string pedal1;
        private string pedal2;
        private string pedal3;
        private string pedal4;
        private string pedal5;

        private string _unspecvalue;
        private string _unspecpremium;

        private string _outofhomecover;
        private string _outofhomepremium;

        private string _outofhomecover2;
        private string _outofhomepremium2;

        private string _inhomecover;
        private string _inhomepremium;

        private string _inhomesum;
        private string _inhomesum2;

        public string Inhomesum
        {
            get
            {
                return _inhomesum;
            }

            set
            {
                _inhomesum = value;
            }
        }
        public string Inhomesum2
        {
            get
            {
                return _inhomesum2;
            }

            set
            {
                _inhomesum2 = value;
            }
        }


        public string Premium
        {
            get
            {
                return _premium;
            }

            set
            {
                _premium = value;
            }
        }

        public string Inhomepremium
        {
            get
            {
                return _inhomepremium;
            }

            set
            {
                _inhomepremium = value;
            }
        }

        public string Inhomecover
        {
            get
            {
                return _inhomecover;
            }

            set
            {
                _inhomecover = value;
            }
        }

        public string Outhomepremium2
        {
            get
            {
                return _outofhomepremium2;
            }

            set
            {
                _outofhomepremium2 = value;
            }
        }

        public string Outhomepremium
        {
            get
            {
                return _outofhomepremium;
            }

            set
            {
                _outofhomepremium = value;
            }
        }

        public string Outhomecover2
        {
            get
            {
                return _outofhomecover2;
            }

            set
            {
                _outofhomecover2 = value;
            }
        }

        public string Outhomecover
        {
            get
            {
                return _outofhomecover;
            }

            set
            {
                _outofhomecover = value;
            }
        }



        public string Unspecpremium
        {
            get
            {
                return _unspecpremium;
            }

            set
            {
                _unspecpremium = value;
            }
        }

        public string Alarm
        {
            get
            {
                return _alarm;
            }

            set
            {
                _alarm = value;
            }
        }

        public string Unspecvalue
        {
            get
            {
                return _unspecvalue;
            }

            set
            {
                _unspecvalue = value;
            }
        }

        public string Contad
        {
            get
            {
                return _contad;
            }

            set
            {
                _contad = value;
            }
        }

        public string Endcode
        {
            get
            {
                return _endcode;
            }

            set
            {
                _endcode = value;
            }
        }

        public string Contsum
        {
            get
            {
                return _contsum;
            }

            set
            {
                _contsum = value;
            }
        }

        public string Contpremium
        {
            get
            {
                return _contpremium;
            }

            set
            {
                _contpremium = value;
            }
        }

        public string Contvol
        {
            get
            {
                return _contvol;
            }

            set
            {
                _contvol = value;
            }
        }

        public string Contratedyear
        {
            get
            {
                return _contratedyear;
            }

            set
            {
                _contratedyear = value;
            }
        }

        public string Contcomp
        {
            get
            {
                return _contcomp;
            }

            set
            {
                _contcomp = value;
            }
        }

        public string Buildad
        {
            get
            {
                return _buildad;
            }

            set
            {
                _buildad = value;
            }
        }

        public string Buildsum
        {
            get
            {
                return _buildsum;
            }

            set
            {
                _buildsum = value;
            }
        }

        public string Buildpremium
        {
            get
            {
                return _buildpremium;
            }

            set
            {
                _buildpremium = value;
            }
        }

        public string noofchildren
        {
            get
            {
                return _noofchildren;
            }

            set
            {
                _noofchildren = value;
            }
        }


        public string noofadults
        {
            get
            {
                return _noofadults;
            }

            set
            {
                _noofadults = value;
            }
        }

        public string Buildvol
        {
            get
            {
                return _buildvol;
            }

            set
            {
                _buildvol = value;
            }
        }

        public string Buildratedyear
        {
            get
            {
                return _buildratedyear;
            }

            set
            {
                _buildratedyear = value;
            }
        }

        public string Buildcomp
        {
            get
            {
                return _buildcomp;
            }

            set
            {
                _buildcomp = value;
            }
        }

        public string Occupancystatus
        {
            get
            {
                return _occupancystatus;
            }

            set
            {
                _occupancystatus = value;
            }
        }

        public string Ownership
        {
            get
            {
                return _ownership;
            }

            set
            {
                _ownership = value;
            }
        }

        public string Businessuse
        {
            get
            {
                return _businessuse;
            }

            set
            {
                _businessuse = value;
            }
        }

        public string Movedate
        {
            get
            {
                return _movedate;
            }

            set
            {
                _movedate = value;
            }
        }

        public string Datebuild
        {
            get
            {
                return _datebuild;
            }

            set
            {
                _datebuild = value;
            }
        }
        public string addPostcode
        {
            get
            {
                return _addpostcode;
            }

            set
            {
                _addpostcode = value;
            }
        }
        public string Noofbed
        {
            get
            {
                return _noofbed;
            }

            set
            {
                _noofbed = value;
            }
        }

        public string Listed
        {
            get
            {
                return _listed;
            }

            set
            {
                _listed = value;
            }
        }

        public string Buildtype
        {
            get
            {
                return _buildtype;
            }

            set
            {
                _buildtype = value;
            }
        }
      
        public bool Addpolicy
        {
            get
            {
                return _addpolicyholder;
            }
            set
            {
                _addpolicyholder = value;
            }
        }
        public bool Concover
        {
            get
            {
                return _concover;
            }
            set
            {
                _concover = value;
            }
        }

        public bool Endorsement
        {
            get
            {
                return endorsement;
            }
            set
            {
                endorsement = value;
            }
        }

        public bool Buscover
        {
            get
            {
                return _buscover;
            }
            set
            {
                _buscover = value;
            }
        }
        public bool Outofhome
        {
            get
            {
                return _outofhome;
            }
            set
            {
                _outofhome = value;
            }
        }

        public bool Inthehome
        {
            get
            {
                return _inthehome;
            }
            set
            {
                _inthehome = value;
            }
        }

        public string Roofpercent
        {
            get
            {
                return _roofpercent;
            }

            set
            {
                _roofpercent = value;
            }
        }

        public string Type
        {
            get
            {
                return _type;
            }

            set
            {
                _type = value;
            }
        }

        public string Roof
        {
            get
            {
                return _roof;
            }

            set
            {
                _roof = value;
            }
        }
        public string Enddate
        {
            get
            {
                return _enddate;
            }

            set
            {
                _enddate = value;
            }
        }
        public string Inceptiondate
        {
            get
            {
                return _inceptiondate;
            }

            set
            {
                _inceptiondate = value;
            }
        }

        public string Walls
        {
            get
            {
                return _walls;
            }

            set
            {
                _walls = value;
            }
        }

        public string Relationship
        {
            get
            {
                return _relationship;
            }

            set
            {
                _relationship = value;
            }
        }

        public string Adddob
        {
            get
            {
                return _adddob;
            }

            set
            {
                _adddob = value;
            }
        }

        public string Addemploymentcatagory
        {
            get
            {
                return _addemploymentcatagory;
            }

            set
            {
                _addemploymentcatagory = value;
            }
        }

        public string Addemployeesbusiness
        {
            get
            {
                return _addemployeesbusiness;
            }

            set
            {
                _addemployeesbusiness = value;
            }
        }

        public string Addmaritialstatus
        {
            get
            {
                return _addmaritialstatus;
            }

            set
            {
                _addmaritialstatus = value;
            }
        }

        public string Addrelationship
        {
            get
            {
                return _relationship;
            }

            set
            {
                _relationship = value;
            }
        }

        public string Addoccupation
        {
            get
            {
                return _addoccupation;
            }

            set
            {
                _addoccupation = value;
            }
        }

        public string Addlastname
        {
            get
            {
                return _addlastname;
            }

            set
            {
                _addlastname = value;
            }
        }

        public string Addfirstname
        {
            get
            {
                return _addfirstname;
            }

            set
            {
                _addfirstname = value;
            }
        }

        public string Addcusttitle
        {
            get
            {
                return _addcusttitle;
            }

            set
            {
                _addcusttitle = value;
            }
        }

        public bool Addpolicyholder
        {
            get
            {
                return _addpolicyholder;
            }

            set
            {
                _addpolicyholder = value;
            }
        }

        public string Dob
        {
            get
            {
                return _dob;
            }

            set
            {
                _dob = value;
            }
        }

        public string Employmentcatagory
        {
            get
            {
                return _employmentcatagory;
            }

            set
            {
                _employmentcatagory = value;
            }
        }

        public string Employeesbusiness
        {
            get
            {
                return _employeesbusiness;
            }

            set
            {
                _employeesbusiness = value;
            }
        }

        public string Maritialstatus
        {
            get
            {
                return _maritialstatus;
            }

            set
            {
                _maritialstatus = value;
            }
        }

        public string Occupation
        {
            get
            {
                return _occupation;
            }

            set
            {
                _occupation = value;
            }
        }

        public string Lastname
        {
            get
            {
                return _lastname;
            }

            set
            {
                _lastname = value;
            }
        }

        public string Firstname
        {
            get
            {
                return _firstname;
            }

            set
            {
                _firstname = value;
            }
        }

        public string Custtitle
        {
            get
            {
                return _custtitle;
            }

            set
            {
                _custtitle = value;
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

        public string Firstline
        {
            get
            {
                return _firstline;
            }

            set
            {
                _firstline = value;
            }
        }
        public string Secondline
        {
            get
            {
                return _secondline;
            }

            set
            {
                _secondline = value;
            }
        }

        public string Thirdline
        {
            get
            {
                return _thirdline;
            }

            set
            {
                _thirdline = value;
            }
        }

        public string Fourthline
        {
            get
            {
                return _fourthline;
            }

            set
            {
                _fourthline = value;
            }
        }

        public string Fifthline
        {
            get
            {
                return _fifthline;
            }

            set
            {
                _fifthline = value;
            }
        }


        public string Sixthline
        {
            get
            {
                return _sixthline;
            }

            set
            {
                _sixthline = value;
            }
        }

        public string AddFirstline
        {
            get
            {
                return _addfirstline;
            }

            set
            {
                _addfirstline = value;
            }
        }
        public string AddSecondline
        {
            get
            {
                return _addsecondline;
            }

            set
            {
                _addsecondline = value;
            }
        }

        public string AddThirdline
        {
            get
            {
                return _addthirdline;
            }

            set
            {
                _addthirdline = value;
            }
        }

        public string AddFourthline
        {
            get
            {
                return _addfourthline;
            }

            set
            {
                _addfourthline = value;
            }
        }

        public string AddFifthline
        {
            get
            {
                return _addfifthline;
            }

            set
            {
                _addfifthline = value; 
            }
        }


        public string AddSixthline
        {
            get
            {
                return _addsixthline;
            }

            set
            {
                _addsixthline = value;
            }
        }


        public string Messagetype
        {
            get
            {
                return _messagetype;
            }

            set
            {
                _messagetype = value;
            }
        }

        public string Poltype
        {
            get
            {
                return _poltype;
            }

            set
            {
                _poltype = value;
            }
        }

        public string Policynum
        {
            get
            {
                return _policynum;
            }

            set
            {
                _policynum = value;
            }
        }

        public string Policynum2
        {
            get
            {
                return _policynum2;
            }

            set
            {
                _policynum2 = value;
            }
        }

        public string custgender
        {
            get
            {
                return _custgender;
            }

            set
            {
                _custgender = value;
            }
        }

        public string Addcustgender
        {
            get
            {
                return _addcustgender;
            }

            set
            {
                _addcustgender = value;
            }
        }

        public string Inhomecount
        {
            get
            {
                return inhomecount;
            }

            set
            {
                inhomecount = value;
            }
        }

        public string Outhomecount
        {
            get
            {
                return outhomecount;
            }

            set
            {
                outhomecount = value;
            }
        }




        public string Pedalcount
        {
            get
            {
                return pedalcount;
            }

            set
            {
                pedalcount = value;
            }
        }

        public string Pedal1
        {
            get
            {
                return pedal1;
            }

            set
            {
                pedal1 = value;
            }
        }

        public string Pedal2
        {
            get
            {
                return pedal2;
            }

            set
            {
                pedal2 = value;
            }
        }

        public string Pedal3
        {
            get
            {
                return pedal3;
            }

            set
            {
                pedal3 = value;
            }
        }

        public string Pedal4
        {
            get
            {
                return pedal4;
            }

            set
            {
                pedal4 = value;
            }
        }
        public string Pedal5
        {
            get
            {
                return pedal5;
            }

            set
            {
                pedal5 = value;
            }
        }

        public string Inhomecover2
        {
            get
            {
                return _inhomepremium;
            }

            set
            {
                _inhomepremium = value;
            }
        }

        public string RiskFirstline
        {
            get
            {
                return _riskfirstline;
            }

            set
            {
                _riskfirstline = value;
            }
        }
        public string RiskSecondline
        {
            get
            {
                return _risksecondline;
            }

            set
            {
                _risksecondline = value;
            }
        }

        public string RiskThirdline
        {
            get
            {
                return _riskthirdline;
            }

            set
            {
                _riskthirdline = value;
            }
        }

        public string RiskFourthline
        {
            get
            {
                return _riskfourthline;
            }

            set
            {
                _riskfourthline = value;
            }
        }

        public string RiskFifthline
        {
            get
            {
                return _riskfifthline;
            }

            set
            {
                _riskfifthline = value;
            }
        }


        public string RiskSixthline
        {
            get
            {
                return _risksixthline;
            }

            set
            {
                _risksixthline = value;
            }
        }

        public string RiskPostcode
        {
            get
            {
                return _risk_postcode;
            }

            set
            {
                _risk_postcode = value;
            }
        }




        public Edi()
        {
            _policynum = "";
            _policynum2 = "";
            _poltype = "";
            _messagetype = "";
            _inceptiondate = "";
            _enddate = "";

            _noofadults = "";
            _noofchildren = "";

            _firstline = "";
            _secondline = "";
            _thirdline = "";
            _fourthline = "";
            _fifthline = "";
            _sixthline = "";
            _postcode = "";

            _custtitle = "";
            _firstname = "";
            _lastname = "";
            _occupation = "";
            _maritialstatus = "";
            _employeesbusiness = "";
            _employmentcatagory = "";
            _dob = "";

            _addpolicyholder = false;
            _addcusttitle = "";
            _addfirstname = "";
            _addlastname = "";
            _addoccupation = "";
            _addmaritialstatus = "";
            _addemployeesbusiness = "";
            _addemploymentcatagory = "";
            _adddob = "";
            _relationship = "";

            _walls = "";
            _roof = "";
            _type = "";
            _buscover = false;
            _concover = false;
            _roofpercent = "";
            _buildtype = "";
            _listed = "";
            _noofbed = "";
            _datebuild = "";
            _movedate = "";
            _businessuse = "";
            _ownership = "";
            _occupancystatus = "";

            _buildcomp = "";
            _buildratedyear = "";
            _buildvol = "";
            _buildpremium = "";
            _buildsum = "";
            _buildad = "";
            _alarm = "";

            _relationship = "";

            _contcomp = "";
            _contratedyear = "";
            _contvol = "";
            _contpremium = "";
            _contsum = "";
            _contad = "";

            _unspecvalue = "";
            _unspecpremium = "";
            _endcode = "";

            _unspecvalue = "";
            _unspecpremium = "";

            _outofhomecover = "";
            _outofhomepremium = "";

            _outofhomecover2 = "";
            _outofhomepremium2 = "";

            _inhomecover = "";
            _inhomepremium = "";
            pedalcount = "";
            pedal1 = "";
            pedal2 = "";
            pedal3 = "";
            pedal4 = "";
            pedal5 = "";

            inhomecount = "";
            outhomecount = "";

    }


        public Edi(string polname, string polname2)
        {
            _policynum = polname;
            _policynum2 = polname2;
        }

        public override string ToString()
        {
            return _policynum;
        }
        
    }
}
