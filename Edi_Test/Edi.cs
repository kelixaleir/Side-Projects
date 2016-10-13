using System;

public class Edi
{
    private string _header;
    private string _policynum;
    private string _poltype;
    private string _messagetype;
    private string _inceptiondate;
    private string _enddate;

    private string _firstline;
    private string _postcode;

    private string _custtitle;
    private string _firstname;
    private string _lastname;
    private string _occupation;
    private string _maritialstatus;
    private string _employeesbusiness;
    private string _employmentcatagory;
    private string _dob;

    private bool _addpolicyholder;
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

    private string _buildcomp;
    private string _buildratedyear;
    private string _buildvol;
    private string _buildpremium;
    private string _buildsum;
    private string _buildad;

    private string _contcomp;
    private string _contratedyear;
    private string _contvol;
    private string _contpremium;
    private string _contsum;
    private string _contad;

    private string _unspecvalue;
    private string _unspecpremium;

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

    public string Header
    {
        get
        {
            return _header;
        }

        set
        {
            _header = value;
        }
    }

    public Edi()
    {

    }

}
