using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using System.IO;
using System.Data.SqlClient;
using System.Globalization;
using Excel;
using System.Text.RegularExpressions;


namespace Edi_Test
{
    public partial class Form1 : Form
    {

        //List of edi messages
        List<List<string>> edicollection = new List<List<string>>();
        //Edi list of each line in an EDI message
        List<Edi> edilist = new List<Edi>();
        //Sql conncetion string
        SqlConnection myConnection = new SqlConnection("Data Source=10.1.3.17,1433;Network Library=DBMSSOCN; Initial Catalog=TEST_EDIDB;User ID=EDIDBUser;Password=WlhbR2WeUj;");
        bool tabcontrol = true;
        string today;
        int selecteditem = -1;
        string policy;
        string concover;
        string buscover;
        string inhome;
        string outhome;
        string endorsement;


        public Form1()
        {

            InitializeComponent();
            start();
            //Setting all the date fields to a standardised format to allow importing from excel easier
            dateTimePicker1.Format = DateTimePickerFormat.Custom;
            dateTimePicker1.CustomFormat = "ddMMyyyy";
            dateTimePicker1.Value = DateTime.Today;
            dateTimePicker2.Format = DateTimePickerFormat.Custom;
            dateTimePicker2.CustomFormat = "ddMMyyyy";
            dateTimePicker2.Value = dateTimePicker1.Value.AddDays(364);
            _custdob.Format = DateTimePickerFormat.Custom;
            _custdob.CustomFormat = "ddMMyyyy";
            _adddob.Format = DateTimePickerFormat.Custom;
            _adddob.CustomFormat = "ddMMyyyy";
            _datebuilt.Format = DateTimePickerFormat.Custom;
            _datebuilt.CustomFormat = "ddMMyyyy";
            _movdate.Format = DateTimePickerFormat.Custom;
            _movdate.CustomFormat = "ddMMyyyy";
            //Additional policyholders disabed by default
            _addtitle.Enabled = false;
            _addfirst.Enabled = false;
            _addlast.Enabled = false;
            _addocc.Enabled = false;
            _addempcat.Enabled = false;
            _addemp.Enabled = false;
            _adddob.Enabled = false;
            _addmar.Enabled = false;
            _addrelat.Enabled = false;
            //Comboboxes set to 0 
            comboBox6.SelectedIndex = 0;
            comboBox8.SelectedIndex = 0;
            comboBox7.SelectedIndex = 0;


            _custgender.SelectedIndex = 0;
            _addcustgender.SelectedIndex = 0;
            endocheck.Checked = true;

            //Additional policyholders disabed by default
            _addcustadd.Enabled = false;
            _addcustadd2.Enabled = false;
            _addcustadd3.Enabled = false;
            _addcustadd4.Enabled = false;
            _addcustadd5.Enabled = false;
            _addcustadd6.Enabled = false;
            _addcustgender.Enabled = false;
            _addcustpost.Enabled = false;
            edilistbox.DataSource = edilist;


            //String to hold todays date in standardised format
            today = DateTime.Now.ToString("ddMMyyyy");



        }

        private void button1_Click(object sender, EventArgs e)
        {
            try
            {
                string dummyFileName = "Save Here";

                SaveFileDialog sf = new SaveFileDialog();
                // Feed the dummy name to the save dialog
                sf.FileName = dummyFileName;

                if (sf.ShowDialog() == DialogResult.OK)
                {
                    // Now here's our save folder
                    string savePath = Path.GetDirectoryName(sf.FileName);
                    // Do whatever

                    System.IO.Directory.CreateDirectory(savePath);
                    int i = 0;
                    foreach (Edi edi in edilist)
                    {

                        addEdi(edi);
                    }
                    foreach (List<string> edistring in edicollection)
                    {

                        StreamWriter file = new StreamWriter(savePath + "\\test" + i + ".edi");

                        foreach (string line in edistring)
                        {


                            file.WriteLine(line);

                        }
                        file.Close();
                        i++;
                    }
                    edicollection.Clear();



                    MessageBox.Show("Export Completed");
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message);
            }

            // Displays a SaveFileDialog so the user can save the file
            // assigned to Button1.
            /*  SaveFileDialog saveFileDialog1 = new SaveFileDialog();
              saveFileDialog1.Filter = "Txt File|*.txt";
              saveFileDialog1.Title = "Save an Image File";
              saveFileDialog1.ShowDialog();
             */
            // If the file name is not an empty string open it for saving.



        }
        private void addEdi(Edi edi3)
        {
            try
            {
                List<string> edimessage = new List<string>();

                Edi edi = edi3;
                string employeebusiness = FindEmployeeBusiness(edi.Employeesbusiness);
                string maritalstatus = FindMaritalStatus(edi.Maritialstatus);
                string ocuppation = FindOccupations(edi.Occupation);
                string propertytype = FindPropertyType(edi.Buildtype);
                string ownership = FindOwnershipType(edi.Ownership);
                string wall = WallConstruction(edi.Walls);
                string roof = FindRoofConstruction(edi.Roof);
                string occstatus = FindOccupancyStatus(edi.Occupancystatus);
                string empstatus = FindEmploymentType(edi.Employmentcatagory);
                string titlecode = FindTitle(edi.Custtitle);

                string incover1 = Findcover(edi.Inhomecover);
                string incover2 = Findcover(edi.Inhomecover2);

                string outhome = Findcover(edi.Outhomecover);
                string outhome2 = Findcover(edi.Outhomecover2);

                string addrelationship = Findrelationship(edi.Relationship);

                string listed = FindListedBuilding(edi.Listed);

                string alarm = Findalarm(edi.Alarm);

                string addemployeebusiness;
                string addmaritalstatus;
                string addocuppation;
                string addempstatus;
                string addtitlecode;
                string tempAddOCD = "";
                int RCAcount = 0;
                int spcount = 0;
                DateTime tempdate = DateTime.ParseExact(edi.Dob, "ddMMyyyy", CultureInfo.InvariantCulture, DateTimeStyles.None);


                DateTime today = DateTime.Today;
                int age = today.Year - tempdate.Year;
                if (_custdob.Value.Date > today.AddYears(-age)) age--;

                edimessage.Add("STX=ANA:1+BECC:SWINTON BRANCH MACHINE+5013546157324:Integra Insurance Solutions Ltd+151204:073306+0000600849 164+        +HHPR03'");
                edimessage.Add("MHD=180+HHPR03:9'");
                edimessage.Add("CID=Z00101:.+C00186:Integra Ho'");
                edimessage.Add("SFI=17:EDI: :01122015:01122015'");
                edimessage.Add("SFT=1'");

                int count = int.Parse(edi.Messagetype);
                count.ToString().PadLeft(2, '0');
                if (edi.Messagetype != null)
                {


                }
                string policynumberend = "";
                if (edi.Policynum2.Length > 3)
                {
                    policynumberend = edi.Policynum.Substring(0, 5);
                }
                switch (edi.Poltype)
                {
                    case "New Business":
                        edimessage.Add("PLH=" + edi.Policynum + edi.Policynum2 + "+NPE:0:PRO:01:S00007+HO+" + edi.Inceptiondate + ":000100+" + edi.Enddate + ":235959+" + edi.Policynum2 + "+++" + policynumberend + "'");
                        break;
                    case "Adjustment":
                        edimessage.Add("PLH=" + edi.Policynum + edi.Policynum2 + "+NPE:0:ADJ:" + edi.Messagetype.PadLeft(2, '0') + ":S00007+HO+" + edi.Inceptiondate + ":000100+" + edi.Enddate + ":235959+" + edi.Policynum2 + "+++" + policynumberend + "'");
                        break;
                    case "Renewal":
                        edimessage.Add("PLH=" + edi.Policynum + edi.Policynum2 + "+NPE:0:RNC:" + count.ToString().PadLeft(2, '0') + ":S00007+HO+" + edi.Inceptiondate + ":000100+" + edi.Enddate + ":235959+" + edi.Policynum2 + "+++" + policynumberend + "'");
                        break;
                    case "Cancellation":
                        edimessage.Add("PLH=" + edi.Policynum + edi.Policynum2 + "+NPE:0:CAB:" + count.ToString().PadLeft(2, '0') + ":S00007+HO+" + edi.Inceptiondate + ":000100+" + edi.Enddate + ":235959+" + edi.Policynum2 + "+++" + policynumberend + "'");
                        break;
                    case "Lapse":
                        edimessage.Add("PLH=" + edi.Policynum + edi.Policynum2 + "+NPE:0:PLL:" + count.ToString().PadLeft(2, '0') + ":S00007+HO+" + edi.Inceptiondate + ":000100+" + edi.Enddate + ":235959+" + edi.Policynum2 + "+++" + policynumberend + "'");
                        break;
                }
                int temppremium = int.Parse(edi.Premium) * 100;
                edimessage.Add("POL=HH+RN+" + edi.Enddate + ":235900+12+Y+Y++12+:Legal & General Insu:31+.+" + edi.Inceptiondate + "'");
                edimessage.Add("PYD=SF:CA+" + temppremium + "++" + temppremium + "++" + temppremium + "'");
                edimessage.Add("RUF=56+TOTAL'");
                double premium2 = temppremium * 0.900;
                edimessage.Add("LDG=++++" + (int)premium2+ "++03'");
                if (endocheck.Checked == false)
                {
                    edimessage.Add("ENA=2+" + edi.Endcode + "+1'");
                    edimessage.Add("ENT=1'");
                }
                double ipt = temppremium - premium2;
                edimessage.Add("TAX=1+5+1000+" + (int)ipt + "'");
                edimessage.Add("TAT=1'");
                edimessage.Add("RUT=1'");



                edimessage.Add("PID=01+P+" + edi.Firstname + ":" + edi.Lastname + "+" + titlecode.ToString().PadLeft(3, '0') + ":" + edi.Custtitle + "+" + edi.custgender + "+" + edi.Dob + "'");
                edimessage.Add("PLO=" + edi.Firstline + ":" + edi.Secondline + ":" + edi.Thirdline + ":" + edi.Fourthline + ":" + edi.Fifthline + ":" + edi.Sixthline + "+" + edi.Postcode + "+++00:01'");
                edimessage.Add("PRO=" + maritalstatus + "+" + empstatus + ":" + ocuppation.ToString().PadLeft(3, '0') + ":" + edi.Occupation + ":Y:" + employeebusiness.ToString().PadLeft(3, '0') + ":" + edi.Employeesbusiness + "+++" + edi.noofchildren + "'");
                if (edi.Addpolicyholder)
                {
                    addemployeebusiness = FindaddEmployeeBusiness(edi.Addemployeesbusiness);
                    addmaritalstatus = FindaddMaritalStatus(edi.Addmaritialstatus);
                    addocuppation = FindaddOccupations(edi.Addoccupation);
                    addempstatus = FindaddEmploymentType(edi.Addemploymentcatagory);
                    addtitlecode = FindaddTitle(edi.Addcusttitle);


                    edimessage.Add("PID=2+P+" + edi.Addfirstname + ":" + edi.Addlastname + "+" + addtitlecode.ToString().PadLeft(3, '0') + ":" + edi.Addcusttitle + "+" + edi.Addcustgender + "+" + edi.Adddob + "'");
                    edimessage.Add("PLO=" + edi.AddFirstline + ":" + edi.AddSecondline + ":" + edi.AddThirdline + ":" + edi.AddFourthline + ":" + edi.AddFifthline + ":" + edi.AddSixthline + "+" + edi.addPostcode + "+++00:01'");
                    edimessage.Add("PRO=" + addmaritalstatus + "+" + addempstatus + ":" + addocuppation.ToString().PadLeft(3, '0') + ":" + edi.Addoccupation + ":Y:" + addemployeebusiness.ToString().PadLeft(3, '0') + ":" + edi.Addemployeesbusiness + "+++'");
                    edimessage.Add("PTR=2'");

                    tempAddOCD = "OCD=" + edi.Addfirstname + ":" + edi.Addlastname + "+" + addempstatus + ":" + addocuppation.ToString().PadLeft(3, '0') + ":" + edi.Addoccupation + ":Y:" + addemployeebusiness.ToString().PadLeft(3, '0') + ":" + edi.Addemployeesbusiness + "++" + edi.Adddob + ":" + age + "+" + addrelationship + "'";
                }
                else
                {
                    edimessage.Add("PTR=1'");
                }
                edimessage.Add("RDD=01+" + propertytype.ToString().PadLeft(2, '0') + ":" + edi.Type + "+" + ownership.ToString().PadLeft(2, '0') + ":" + edi.Ownership + "+N+99+" + edi.RiskFirstline + ":" + edi.RiskSecondline + ":" + edi.RiskThirdline + ":" + edi.RiskFourthline + ":" + edi.RiskFifthline + ":" + edi.RiskSixthline + "+" + edi.RiskPostcode + "+++" + edi.Noofbed + "++" + listed.PadLeft(2, '0') + "++" + edi.Datebuild + "'");
                edimessage.Add("RDC=" + wall.PadLeft(2, '0') + ":" + roof.PadLeft(2, '0') + "::" + edi.Roofpercent.PadLeft(3, '0') + "+Y:Y+Y+Y+N+N+N'");
                edimessage.Add("RDU=" + occstatus + "+N+Y+Y+N:30+++N++N:N+" + edi.Movedate + "'");
                edimessage.Add("RDS=Y+Y+N+N+Y+" + alarm + "+N++++005'");
                edimessage.Add("OCD=" + edi.Firstname + ":" + edi.Lastname + "+" + empstatus + ":" + ocuppation.ToString().PadLeft(3, '0') + ":" + edi.Occupation + ":Y:" + employeebusiness.ToString().PadLeft(3, '0') + ":" + edi.Employeesbusiness + "++" + edi.Dob + ":" + age + "+P'");

                if (edi.Addpolicyholder)
                {
                    edimessage.Add(tempAddOCD);
                    edimessage.Add("ODT=2'");
                }
                else
                {
                    edimessage.Add("ODT=1'");
                }
                if (edi.Buscover == false)
                {
                    //  if (int.Parse(edi.Buildsum) > 0)
                    //  {
                    RCAcount++;

                    edimessage.Add("RCA=" + RCAcount.ToString().PadLeft(2, '0') + "+B05+1+N:" + edi.Buildsum + "+" + edi.Buildad + "+" + edi.Businessuse + "+Y:" + edi.Buildsum + "+" + edi.Buildvol + "+" + _datebuilt.Value.Date.Year + "+ +" + edi.RiskPostcode + "'");
                    edimessage.Add("RUF=56+TOTAL'");
                    edimessage.Add("LDG=++++0++03'");
                    edimessage.Add("CXS=" + edi.Buildcomp + ":E'");
                    edimessage.Add("CXT=1'");
                    edimessage.Add("TAX=1+5+1000+0'");
                    edimessage.Add("TAT=1'");
                    edimessage.Add("RUT=1'");
                    edimessage.Add("NCB=Y+2:" + edi.Buildratedyear + "+7+N'");

                    //  }
                }
                if (edi.Concover == false)
                {
                    //   if (int.Parse(edi.Contsum) > 0)
                    // {
                    RCAcount++;

                    edimessage.Add("RCA=" + RCAcount.ToString().PadLeft(2, '0') + "+C14+1+N:" + edi.Contsum + "+" + edi.Contad + "+" + edi.Businessuse + "+Y:" + int.Parse(edi.Contsum) + "+" + edi.Contvol + "++ +" + edi.RiskPostcode + "'");
                    edimessage.Add("RUF=56+TOTAL'");
                    edimessage.Add("LDG=++++0++03'");
                    edimessage.Add("CXS=" + edi.Contcomp + ":F'");
                    edimessage.Add("CXT=1'");
                    edimessage.Add("TAX=1+5+1000+0'");
                    edimessage.Add("TAT=1'");
                    edimessage.Add("RUT=1'");
                    edimessage.Add("NCB=Y+2:" + edi.Contratedyear + "+7+N'");

                    //  }


                    switch (int.Parse(edi.Inhomecount))
                    {
                        case 0:
                            break;
                        case 1:

                        //change sum insured
                        edimessage.Add("SPI=1+" + incover1 + ":Within Home 1+" + (int.Parse(edi.Inhomesum)) + ":" + (int.Parse(edi.Inhomesum)) + ":N+" + edi.RiskPostcode + "'");
                        edimessage.Add("RUF=56+TOTAL'");
                        edimessage.Add("LDG=++++0'");
                        edimessage.Add("TAX=1+5+1000+0'");
                        edimessage.Add("TAT=1'");
                        edimessage.Add("RUT=1'");
                        edimessage.Add("SPT=1'");

                            break;
                        case 2:
                            //change sum insured
                            edimessage.Add("SPI=1+" + incover1 + ":Within Home 1+" + (int.Parse(edi.Inhomesum)) + ":" + (int.Parse(edi.Inhomesum)) + ":N+" + edi.RiskPostcode + "'");
                            edimessage.Add("RUF=56+TOTAL'");
                            edimessage.Add("LDG=++++0'");
                            edimessage.Add("TAX=1+5+1000+0'");
                            edimessage.Add("TAT=1'");
                            edimessage.Add("RUT=1'");
                            spcount++;
                            //change sum insured
                            edimessage.Add("SPI=2+" + incover2 + ":Within Home 2+" + (int.Parse(edi.Inhomesum2)) + ":" + (int.Parse(edi.Inhomesum2)) + ":N+" + edi.RiskPostcode + "'");
                            edimessage.Add("RUF=56+TOTAL'");
                            edimessage.Add("LDG=++++0'");
                            edimessage.Add("TAX=1+5+1000+0'");
                            edimessage.Add("TAT=1'");
                            edimessage.Add("RUT=1'");
                            edimessage.Add("SPT=2'");
                            break;
                        default:
                            break;
                    }

                }

                if (int.Parse(edi.Unspecvalue) > 0)
                {
                    RCAcount++;
                    edimessage.Add("RCA=" + RCAcount.ToString().PadLeft(2, '0') + "+U01+1+N:" + edi.Unspecvalue + "++" + edi.Businessuse + "+Y:" + int.Parse(edi.Unspecvalue) + "+++N+" + edi.RiskPostcode + "'");
                    edimessage.Add("RUF=56+TOTAL'");
                    edimessage.Add("LDG=++++0++03'");
                    edimessage.Add("TAX=1+5+1000+0'");
                    edimessage.Add("TAT=1'");
                    edimessage.Add("RUT=1'");
                }

                switch (int.Parse(edi.Outhomecount))
                {
                    case 0:
                        break;
                    case 1:
                        RCAcount++;
                        edimessage.Add("RCA=" + RCAcount.ToString().PadLeft(2, '0') + "+S03+1+Y:" + (int.Parse(edi.Outhomepremium)) + ":1++N+Y+++ +" + edi.RiskPostcode + "'");
                        edimessage.Add("RUF=56+TOTAL'");
                        edimessage.Add("LDG=++++0'");
                        edimessage.Add("TAX=1+5+1000+0'");
                        edimessage.Add("TAT=1'");
                        edimessage.Add("RUT=1'");

                        edimessage.Add("SPI=1+" + outhome + ":Specified Item 1+" + (int.Parse(edi.Outhomepremium)) + ":" + (int.Parse(edi.Outhomepremium)) + ":N+" + edi.RiskPostcode + "'");
                        edimessage.Add("RUF=56+TOTAL'");
                        edimessage.Add("LDG=++++0'");
                        edimessage.Add("TAX=1+5+1000+0'");
                        edimessage.Add("TAT=1'");
                        edimessage.Add("RUT=1'");
                        edimessage.Add("SPT=1'");
                        break;
                    case 2:
                        edimessage.Add("RCA=" + RCAcount.ToString().PadLeft(2, '0') + "+S03+1+Y:" + ((int.Parse(edi.Outhomepremium)) + (int.Parse(edi.Outhomepremium2))) + ":2++N+Y+++ +" + edi.RiskPostcode + "'");
                        edimessage.Add("RUF=56+TOTAL'");
                        edimessage.Add("LDG=++++0'");
                        edimessage.Add("TAX=1+5+1000+0'");
                        edimessage.Add("TAT=1'");
                        edimessage.Add("RUT=1'");

                        edimessage.Add("SPI=1+" + outhome + ":Specified Item 1+" + (int.Parse(edi.Outhomepremium))+ ":" + (int.Parse(edi.Outhomepremium)) + ":N+" + edi.RiskPostcode + "'");
                        edimessage.Add("RUF=56+TOTAL'");
                        edimessage.Add("LDG=++++0'");
                        edimessage.Add("TAX=1+5+1000+0'");
                        edimessage.Add("TAT=1'");
                        edimessage.Add("RUT=1'");

                        edimessage.Add("SPI=2+" + outhome2 + ":Specified Item 1+" + (int.Parse(edi.Outhomepremium2)) + ":" + (int.Parse(edi.Outhomepremium2)) + ":N+" + edi.RiskPostcode + "'");
                        edimessage.Add("RUF=56+TOTAL'");
                        edimessage.Add("LDG=++++0'");
                        edimessage.Add("TAX=1+5+1000+0'");
                        edimessage.Add("TAT=1'");
                        edimessage.Add("RUT=1'");
                        edimessage.Add("SPT=2'");
                        break;
                    default:
                        break;



                }

                switch (int.Parse(edi.Pedalcount))
                {
                    case 0:
                        break;
                    case 1:
                        RCAcount++;
                        edimessage.Add("RCA=" + RCAcount.ToString().PadLeft(2, '0') + "+P02+1+Y:" + int.Parse(edi.Pedal1) + ":1++N+++0000'");
                        edimessage.Add("RUF=56+TOTAL'");
                        edimessage.Add("LDG=++++0'");
                        edimessage.Add("RUT=1'");
                        edimessage.Add("SPI=1+P02:Bicycle 1+" + int.Parse(edi.Pedal1)+ "::N'");
                        edimessage.Add("RUF=56+TOTAL'");
                        edimessage.Add("LDG=++++0'");
                        edimessage.Add("TAX=1+5+1000+0'");
                        edimessage.Add("TAT=1'");
                        edimessage.Add("RUT=1'");
                        edimessage.Add("SPT=1'");

                        break;
                    case 2:
                        RCAcount++;
                        edimessage.Add("RCA=" + RCAcount.ToString().PadLeft(2, '0') + "+P02+1+Y:" + ((int.Parse(edi.Pedal1)) + (int.Parse(edi.Pedal2))) + ":2++N+++0000'");
                        edimessage.Add("RUF=56+TOTAL'");
                        edimessage.Add("LDG=++++0'");
                        edimessage.Add("RUT=1'");
                        edimessage.Add("SPI=1+P02:Bicycle 1+" + int.Parse(edi.Pedal1) + "::N'");
                        edimessage.Add("RUF=56+TOTAL'");
                        edimessage.Add("LDG=++++0'");
                        edimessage.Add("TAX=1+5+1000+0'");
                        edimessage.Add("TAT=1'");
                        edimessage.Add("RUT=1'");
                        edimessage.Add("SPI=2+P02:Bicycle 2+" + int.Parse(edi.Pedal2) + "::N'");
                        edimessage.Add("RUF=56+TOTAL'");
                        edimessage.Add("LDG=++++0'");
                        edimessage.Add("TAX=1+5+1000+0'");
                        edimessage.Add("TAT=1'");
                        edimessage.Add("RUT=1'");
                        edimessage.Add("SPT=2'");
                        break;
                    case 3:
                        RCAcount++;
                        edimessage.Add("RCA=" + RCAcount.ToString().PadLeft(2, '0') + "+P02+1+Y:" + ((int.Parse(edi.Pedal1)) + (int.Parse(edi.Pedal2)) + (int.Parse(edi.Pedal3) )) + ":3++N+++0000'");
                        edimessage.Add("RUF=56+TOTAL'");
                        edimessage.Add("LDG=++++0'");
                        edimessage.Add("RUT=1'");
                        edimessage.Add("SPI=1+P02:Bicycle 1+" + int.Parse(edi.Pedal1) + "::N'");
                        edimessage.Add("RUF=56+TOTAL'");
                        edimessage.Add("LDG=++++0'");
                        edimessage.Add("TAX=1+5+1000+0'");
                        edimessage.Add("TAT=1'");
                        edimessage.Add("RUT=1'");
                        edimessage.Add("SPI=2+P02:Bicycle 2+" + int.Parse(edi.Pedal2) + "::N'");
                        edimessage.Add("RUF=56+TOTAL'");
                        edimessage.Add("LDG=++++0'");
                        edimessage.Add("TAX=1+5+1000+0'");
                        edimessage.Add("TAT=1'");
                        edimessage.Add("RUT=1'");
                        edimessage.Add("SPI=3+P02:Bicycle 3+" + int.Parse(edi.Pedal3) + "::N'");
                        edimessage.Add("RUF=56+TOTAL'");
                        edimessage.Add("LDG=++++0'");
                        edimessage.Add("TAX=1+5+1000+0'");
                        edimessage.Add("TAT=1'");
                        edimessage.Add("RUT=1'");
                        edimessage.Add("SPT=3'");
                        break;
                    case 4:
                        RCAcount++;
                        edimessage.Add("RCA=" + RCAcount.ToString().PadLeft(2, '0') + "+P02+1+Y:" + ((int.Parse(edi.Pedal1)) + (int.Parse(edi.Pedal2)) + (int.Parse(edi.Pedal3)) + (int.Parse(edi.Pedal4))) + ":4++N+++0000'");
                        edimessage.Add("RUF=56+TOTAL'");
                        edimessage.Add("LDG=++++0'");
                        edimessage.Add("RUT=1'");
                        edimessage.Add("SPI=1+P02:Bicycle 1+" + int.Parse(edi.Pedal1) + "::N'");
                        edimessage.Add("RUF=56+TOTAL'");
                        edimessage.Add("LDG=++++0'");
                        edimessage.Add("TAX=1+5+1000+0'");
                        edimessage.Add("TAT=1'");
                        edimessage.Add("RUT=1'");
                        edimessage.Add("SPI=2+P02:Bicycle 2+" + int.Parse(edi.Pedal2) + "::N'");
                        edimessage.Add("RUF=56+TOTAL'");
                        edimessage.Add("LDG=++++0'");
                        edimessage.Add("TAX=1+5+1000+0'");
                        edimessage.Add("TAT=1'");
                        edimessage.Add("RUT=1'");
                        edimessage.Add("SPI=3+P02:Bicycle 3+" + int.Parse(edi.Pedal3) + "::N'");
                        edimessage.Add("RUF=56+TOTAL'");
                        edimessage.Add("LDG=++++0'");
                        edimessage.Add("TAX=1+5+1000+0'");
                        edimessage.Add("TAT=1'");
                        edimessage.Add("RUT=1'");
                        edimessage.Add("SPI=4+P02:Bicycle 4+" + int.Parse(edi.Pedal4) + "::N'");
                        edimessage.Add("RUF=56+TOTAL'");
                        edimessage.Add("LDG=++++0'");
                        edimessage.Add("TAX=1+5+1000+0'");
                        edimessage.Add("TAT=1'");
                        edimessage.Add("RUT=1'");
                        edimessage.Add("SPT=4'");
                        break;
                    case 5:
                        RCAcount++;
                        edimessage.Add("RCA=" + RCAcount.ToString().PadLeft(2, '0') + "+P02+1+Y:" + ((int.Parse(edi.Pedal1)) + (int.Parse(edi.Pedal2) ) + (int.Parse(edi.Pedal3)) + (int.Parse(edi.Pedal4)) + (int.Parse(edi.Pedal5))) + ":5++N+++0000'");
                        edimessage.Add("RUF=56+TOTAL'");
                        edimessage.Add("LDG=++++0'");
                        edimessage.Add("RUT=1'");
                        edimessage.Add("SPI=1+P02:Bicycle 1+" + int.Parse(edi.Pedal1) + "::N'");
                        edimessage.Add("RUF=56+TOTAL'");
                        edimessage.Add("LDG=++++0'");
                        edimessage.Add("TAX=1+5+1000+0'");
                        edimessage.Add("TAT=1'");
                        edimessage.Add("RUT=1'");
                        edimessage.Add("SPI=2+P02:Bicycle 2+" + int.Parse(edi.Pedal2) + "::N'");
                        edimessage.Add("RUF=56+TOTAL'");
                        edimessage.Add("LDG=++++0'");
                        edimessage.Add("TAX=1+5+1000+0'");
                        edimessage.Add("TAT=1'");
                        edimessage.Add("RUT=1'");
                        edimessage.Add("SPI=3+P02:Bicycle 3+" + int.Parse(edi.Pedal3) + "::N'");
                        edimessage.Add("RUF=56+TOTAL'");
                        edimessage.Add("LDG=++++0'");
                        edimessage.Add("TAX=1+5+1000+0'");
                        edimessage.Add("TAT=1'");
                        edimessage.Add("RUT=1'");
                        edimessage.Add("SPI=4+P02:Bicycle 4+" + int.Parse(edi.Pedal4) + "::N'");
                        edimessage.Add("RUF=56+TOTAL'");
                        edimessage.Add("LDG=++++0'");
                        edimessage.Add("TAX=1+5+1000+0'");
                        edimessage.Add("TAT=1'");
                        edimessage.Add("RUT=1'");
                        edimessage.Add("SPI=5+P02:Bicycle 5+" + int.Parse(edi.Pedal5) + "::N'");
                        edimessage.Add("RUF=56+TOTAL'");
                        edimessage.Add("LDG=++++0'");
                        edimessage.Add("TAX=1+5+1000+0'");
                        edimessage.Add("TAT=1'");
                        edimessage.Add("RUT=1'");
                        edimessage.Add("SPT=5'");
                        break;
                    default:
                        Console.WriteLine("Default case");
                        break;
                }

                //PEDAL CYCLES

                /*
    RCA=(RCACOUNT)+P02+1+Y:(TOTAL Sum insured):(Total no of spec items)++N+++0000'
    RUF=56+TOTAL'
    LDG=++++5563000'
    RUT=1'

    SPI=(SPI Count)+P02:Bicycle 1+(Sum insured)::N'
    RUF=56+TOTAL'
    LDG=++++0'
    TAX=1+5+0950+0'
    TAT=1'
    RUT=1'
    SPI=(SPI Count)+P02:Bicyle 2+(Sum insured)::N'
    RUF=56+TOTAL'
    LDG=++++0'
    TAX=1+5+0950+0'
    TAT=1'
    RUT=1'
    SPT=2'



    */

                edimessage.Add("RIT=" + RCAcount + "'");
                edimessage.Add("RDT=1'");
                edimessage.Add("MTR=" + edimessage.Count + "'");
                edimessage.Add("END=1'");
                edicollection.Add(edimessage);

            }





            catch (Exception ex)
            {
                MessageBox.Show(ex.Message);
            }




        }
        private string FindEmployeeBusiness(string empbusiness)
        {

            myConnection.Open();

            SqlCommand cmd = myConnection.CreateCommand();
            cmd.CommandText = "SELECT ABI FROM ['Employers Business$'] WHERE[Employer Business] = '" + empbusiness + "'";

            string result = "";


            using (SqlDataReader reader = cmd.ExecuteReader())
            {
                while (reader.Read())
                {
                    //Send these to your WinForms textboxes
                    result = reader["ABI"].ToString();
                }
            }

            myConnection.Close();
            return result;


        }

        private string Findcover(string cover)
        {

            myConnection.Open();

            SqlCommand cmd = myConnection.CreateCommand();
            cmd.CommandText = "SELECT ABI FROM [Cover$] WHERE[Cover] = '" + cover + "'";

            string result = "";


            using (SqlDataReader reader = cmd.ExecuteReader())
            {
                while (reader.Read())
                {
                    //Send these to your WinForms textboxes
                    result = reader["ABI"].ToString();
                }
            }

            myConnection.Close();
            return result;


        }

        private string FindEmploymentType(string emptype)
        {
            myConnection.Open();

            SqlCommand cmd = myConnection.CreateCommand();
            cmd.CommandText = "SELECT ABI FROM ['Employment Type$'] WHERE[Employee Type] = '" + emptype + "'";

            string result = "";


            using (SqlDataReader reader = cmd.ExecuteReader())
            {
                while (reader.Read())
                {
                    //Send these to your WinForms textboxes
                    result = reader["ABI"].ToString();
                }
            }

            myConnection.Close();
            return result;
        }

        private string Findrelationship(string marrstatus)
        {
            myConnection.Open();

            SqlCommand cmd = myConnection.CreateCommand();
            cmd.CommandText = "SELECT ABI FROM [Relationship$] WHERE[Relationship] = '" + marrstatus + "'";

            string result = "";


            using (SqlDataReader reader = cmd.ExecuteReader())
            {
                while (reader.Read())
                {
                    //Send these to your WinForms textboxes
                    result = reader["ABI"].ToString();
                }
            }

            myConnection.Close();
            return result;
        }

        private string Findalarm(string alarm)
        {
            myConnection.Open();

            SqlCommand cmd = myConnection.CreateCommand();
            cmd.CommandText = "SELECT ABI FROM [Alarm$] WHERE[Alarm] = '" + alarm + "'";

            string result = "";


            using (SqlDataReader reader = cmd.ExecuteReader())
            {
                while (reader.Read())
                {
                    //Send these to your WinForms textboxes
                    result = reader["ABI"].ToString();
                }
            }

            myConnection.Close();
            return result;
        }


        private string FindListedBuilding(string listbuild)
        {
            myConnection.Open();

            SqlCommand cmd = myConnection.CreateCommand();
            cmd.CommandText = "SELECT ABI FROM ['Listed Building$'] WHERE[Listed Building] = '" + listbuild + "'";

            string result = "";


            using (SqlDataReader reader = cmd.ExecuteReader())
            {
                while (reader.Read())
                {
                    //Send these to your WinForms textboxes
                    result = reader["ABI"].ToString();
                }
            }

            myConnection.Close();
            return result;
        }

        private string FindMaritalStatus(string marstat)
        {
            myConnection.Open();

            SqlCommand cmd = myConnection.CreateCommand();
            cmd.CommandText = "SELECT ABI FROM ['Marital Status$'] WHERE[Maritial Status] = '" + marstat + "'";

            string result = "";


            using (SqlDataReader reader = cmd.ExecuteReader())
            {
                while (reader.Read())
                {
                    //Send these to your WinForms textboxes
                    result = reader["ABI"].ToString();
                }
            }

            myConnection.Close();
            return result;
        }

        private string RetrieveMaritalStatus(string marstat)
        {
            myConnection.Open();

            SqlCommand cmd = myConnection.CreateCommand();
            cmd.CommandText = "SELECT [Maritial Status] FROM ['Marital Status$'] WHERE ABI = '" + marstat + "'";

            string result = "";


            using (SqlDataReader reader = cmd.ExecuteReader())
            {
                while (reader.Read())
                {
                    //Send these to your WinForms textboxes
                    result = reader["Maritial Status"].ToString();
                }
            }

            myConnection.Close();
            return result;
        }


        private string RetrieveWalls(string walls)
        {
            myConnection.Open();

            SqlCommand cmd = myConnection.CreateCommand();
            cmd.CommandText = "SELECT [Wall Construction] FROM ['Wall Construction$'] WHERE ABI = '" + walls + "'";

            string result = "";


            using (SqlDataReader reader = cmd.ExecuteReader())
            {
                while (reader.Read())
                {
                    //Send these to your WinForms textboxes
                    result = reader["Wall Construction"].ToString();
                }
            }

            myConnection.Close();
            return result;
        }

        private string RetrieveRoof(string roof)
        {
            myConnection.Open();

            SqlCommand cmd = myConnection.CreateCommand();
            cmd.CommandText = "SELECT [Roof Construction] FROM ['Roof Construction$'] WHERE ABI = '" + roof + "'";

            string result = "";


            using (SqlDataReader reader = cmd.ExecuteReader())
            {
                while (reader.Read())
                {
                    //Send these to your WinForms textboxes
                    result = reader["Roof Construction"].ToString();
                }
            }

            myConnection.Close();
            return result;
        }
        private string RetrieveType(string type)
        {
            myConnection.Open();

            SqlCommand cmd = myConnection.CreateCommand();
            cmd.CommandText = "SELECT [Property Type] FROM ['Property Type$'] WHERE ABI = '" + type + "'";

            string result = "";


            using (SqlDataReader reader = cmd.ExecuteReader())
            {
                while (reader.Read())
                {
                    //Send these to your WinForms textboxes
                    result = reader["Property Type"].ToString();
                }
            }

            myConnection.Close();
            return result;
        }
        private string RetrieveOwnership(string owner)
        {

            myConnection.Open();

            SqlCommand cmd = myConnection.CreateCommand();
            cmd.CommandText = "SELECT [Ownership Type] FROM ['Ownership Type$'] WHERE ABI = '" + owner + "'";

            string result = "";


            using (SqlDataReader reader = cmd.ExecuteReader())
            {
                while (reader.Read())
                {
                    //Send these to your WinForms textboxes
                    result = reader["Ownership Type"].ToString();
                }
            }

            myConnection.Close();
            return result;

        }

        private string RetrieveEmploymentType(string marstat)
        {
            myConnection.Open();

            SqlCommand cmd = myConnection.CreateCommand();
            cmd.CommandText = "SELECT [Employee Type] FROM ['Employment Type$'] WHERE ABI = '" + marstat + "'";

            string result = "";


            using (SqlDataReader reader = cmd.ExecuteReader())
            {
                while (reader.Read())
                {
                    //Send these to your WinForms textboxes
                    result = reader["Employee Type"].ToString();
                }
            }

            myConnection.Close();
            return result;
        }

        private string FindOccupancyStatus(string occstat)
        {
            myConnection.Open();

            SqlCommand cmd = myConnection.CreateCommand();
            cmd.CommandText = "SELECT ABI FROM ['Occupancy status$'] WHERE[Occupancy Status] = '" + occstat + "'";

            string result = "";


            using (SqlDataReader reader = cmd.ExecuteReader())
            {
                while (reader.Read())
                {
                    //Send these to your WinForms textboxes
                    result = reader["ABI"].ToString();
                }
            }

            myConnection.Close();
            return result;
        }

        private string FindOccupations(string occ)
        {
            myConnection.Open();

            SqlCommand cmd = myConnection.CreateCommand();
            cmd.CommandText = "SELECT ABI FROM [Occupations$] WHERE[Occupations] = '" + occ + "'";

            string result = "";


            using (SqlDataReader reader = cmd.ExecuteReader())
            {
                while (reader.Read())
                {
                    //Send these to your WinForms textboxes
                    result = reader["ABI"].ToString();
                }
            }

            myConnection.Close();
            return result;
        }

        private string FindOwnershipType(string owner)
        {
            myConnection.Open();

            SqlCommand cmd = myConnection.CreateCommand();
            cmd.CommandText = "SELECT ABI FROM ['Ownership Type$'] WHERE[Ownership Type] = '" + owner + "'";

            string result = "";


            using (SqlDataReader reader = cmd.ExecuteReader())
            {
                while (reader.Read())
                {
                    //Send these to your WinForms textboxes
                    result = reader["ABI"].ToString();
                }
            }

            myConnection.Close();
            return result;
        }

        private string FindPropertyType(string prop)
        {
            myConnection.Open();

            SqlCommand cmd = myConnection.CreateCommand();
            cmd.CommandText = "SELECT ABI FROM ['Property Type$'] WHERE[Property Type] = '" + prop + "'";

            string result = "";


            using (SqlDataReader reader = cmd.ExecuteReader())
            {
                while (reader.Read())
                {
                    //Send these to your WinForms textboxes
                    result = reader["ABI"].ToString();
                }
            }

            myConnection.Close();
            return result;
        }

        private string FindRoofConstruction(string roofc)
        {
            myConnection.Open();

            SqlCommand cmd = myConnection.CreateCommand();
            cmd.CommandText = "SELECT ABI FROM ['Roof Construction$'] WHERE[Roof Construction] = '" + roofc + "'";

            string result = "";


            using (SqlDataReader reader = cmd.ExecuteReader())
            {
                while (reader.Read())
                {
                    //Send these to your WinForms textboxes
                    result = reader["ABI"].ToString();
                }
            }

            myConnection.Close();
            return result;
        }

        private string FindTitle(string title)
        {
            myConnection.Open();

            SqlCommand cmd = myConnection.CreateCommand();
            cmd.CommandText = "SELECT ABI FROM [Title$] WHERE[Title] = '" + title + "'";

            string result = "";


            using (SqlDataReader reader = cmd.ExecuteReader())
            {
                while (reader.Read())
                {
                    //Send these to your WinForms textboxes
                    result = reader["ABI"].ToString();
                }
            }

            myConnection.Close();
            return result;
        }

        private string WallConstruction(string wall)
        {
            myConnection.Open();

            SqlCommand cmd = myConnection.CreateCommand();
            cmd.CommandText = "SELECT ABI FROM ['Wall Construction$'] WHERE[Wall Construction] = '" + wall + "'";

            string result = "";


            using (SqlDataReader reader = cmd.ExecuteReader())
            {
                while (reader.Read())
                {
                    //Send these to your WinForms textboxes
                    result = reader["ABI"].ToString();
                }
            }

            myConnection.Close();
            return result;
        }
        private void start()
        {
            try
            {
                _editype.SelectedIndex = 0;
                _custitle.SelectedIndex = 0;
                _occu.SelectedIndex = 0;
                _marrstatus.SelectedIndex = 0;
                _addmar.SelectedIndex = 0;
                comboBox2.SelectedIndex = 0;
                _emptype.SelectedIndex = 0;

                _addtitle.SelectedIndex = 0;
                _addocc.SelectedIndex = 0;
                _addmar.SelectedIndex = 0;
                _addemp.SelectedIndex = 0;
                _addempcat.SelectedIndex = 0;
                _addrelat.SelectedIndex = 0;

                _wallcon.SelectedIndex = 0;
                _roofcon.SelectedIndex = 0;
                _proptype.SelectedIndex = 0;
                _listedbuild.SelectedIndex = 0;
                _numbed.SelectedIndex = 0;
                _owner.SelectedIndex = 0;

                _occstatus.SelectedIndex = 0;
                _physec.SelectedIndex = 0;
                _secalarm.SelectedIndex = 0;
                _bususe.SelectedIndex = 0;
                _conAD.SelectedIndex = 0;
                buildAD.SelectedIndex = 0;

                comboBox1.SelectedIndex = 0;
                comboBox3.SelectedIndex = 0;
                comboBox4.SelectedIndex = 0;
                comboBox5.SelectedIndex = 0;
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message);
            }
        }



        private string FindaddEmployeeBusiness(string bus)
        {
            myConnection.Open();

            SqlCommand cmd = myConnection.CreateCommand();
            cmd.CommandText = "SELECT ABI FROM ['Employers Business$'] WHERE[Employer Business] = '" + bus + "'";

            string result = "";


            using (SqlDataReader reader = cmd.ExecuteReader())
            {
                while (reader.Read())
                {
                    //Send these to your WinForms textboxes
                    result = reader["ABI"].ToString();
                }
            }

            myConnection.Close();
            return result;
        }
        private string FindaddOccupations(string occ)
        {
            myConnection.Open();

            SqlCommand cmd = myConnection.CreateCommand();
            cmd.CommandText = "SELECT ABI FROM [Occupations$] WHERE[Occupations] = '" + occ + "'";

            string result = "";


            using (SqlDataReader reader = cmd.ExecuteReader())
            {
                while (reader.Read())
                {
                    //Send these to your WinForms textboxes
                    result = reader["ABI"].ToString();
                }
            }

            myConnection.Close();
            return result;
        }
        private string FindaddMaritalStatus(string mar)
        {
            myConnection.Open();

            SqlCommand cmd = myConnection.CreateCommand();
            cmd.CommandText = "SELECT ABI FROM ['Marital Status$'] WHERE[Maritial Status] = '" + mar + "'";

            string result = "";


            using (SqlDataReader reader = cmd.ExecuteReader())
            {
                while (reader.Read())
                {
                    //Send these to your WinForms textboxes
                    result = reader["ABI"].ToString();
                }
            }

            myConnection.Close();
            return result;
        }

        private string FindaddTitle(string title)
        {
            myConnection.Open();

            SqlCommand cmd = myConnection.CreateCommand();
            cmd.CommandText = "SELECT ABI FROM [Title$] WHERE[Title] = '" + title + "'";

            string result = "";


            using (SqlDataReader reader = cmd.ExecuteReader())
            {
                while (reader.Read())
                {
                    //Send these to your WinForms textboxes
                    result = reader["ABI"].ToString();
                }
            }

            myConnection.Close();
            return result;
        }


        private string FindaddEmploymentType(string emp)
        {
            myConnection.Open();

            SqlCommand cmd = myConnection.CreateCommand();
            cmd.CommandText = "SELECT ABI FROM ['Employment Type$'] WHERE[Employee Type] = '" + emp + "'";

            string result = "";


            using (SqlDataReader reader = cmd.ExecuteReader())
            {
                while (reader.Read())
                {
                    //Send these to your WinForms textboxes
                    result = reader["ABI"].ToString();
                }
            }

            myConnection.Close();
            return result;
        }

        private void dateTimePicker1_ValueChanged(object sender, EventArgs e)
        {

        }

        private void tabPage1_Click(object sender, EventArgs e)
        {

        }

        private void Form1_Load(object sender, EventArgs e)
        {

        }

        private void groupBox1_Enter(object sender, EventArgs e)
        {

        }

        private void label6_Click(object sender, EventArgs e)
        {

        }

        private void _editype_SelectedIndexChanged(object sender, EventArgs e)
        {

        }

        private void comboBox3_SelectedIndexChanged(object sender, EventArgs e)
        {

        }

        private void comboBox8_SelectedIndexChanged(object sender, EventArgs e)
        {

        }

        private void groupBox4_Enter(object sender, EventArgs e)
        {

        }

        private void _customerDetails_Enter(object sender, EventArgs e)
        {

        }

        private void textBox21_TextChanged(object sender, EventArgs e)
        {

        }

        private void checkBox1_CheckedChanged(object sender, EventArgs e)
        {

        }
        private double calculatePremium(string bpremium, string cpremium, string vpremium)
        {


            double b = double.Parse(bpremium) * 100;
            double c = double.Parse(cpremium) * 100;
            double v = double.Parse(vpremium) * 100;

            return b + c + v;
        }

        private void textBox1_TextChanged(object sender, EventArgs e)
        {

        }

        private void _flatroof_ValueChanged(object sender, EventArgs e)
        {

        }

        private void _building_Click(object sender, EventArgs e)
        {

        }

        private void label46_Click(object sender, EventArgs e)
        {

        }

        private void comboBox1_SelectedIndexChanged(object sender, EventArgs e)
        {

        }

        private void groupBox6_Enter(object sender, EventArgs e)
        {

        }
        void BindData()
        {
            edilistbox.DataSource = null;
            edilistbox.DataSource = edilist;
            edilistbox.DisplayMember = "Name";

        }


        private void button3_Click(object sender, EventArgs e)
        {
            //If the selected item in the listbox exists, overwrite this, set the selected number back to -1 (So no item is selected) 
            //and binds the data from the list to the list box
            if (selecteditem > -1)
            {
                saveedi(edilist[selecteditem]);
                edilistbox.SelectedIndex = -1;
                selecteditem = -1;
                BindData();
            }
            //Else if nothing is selected to will crete and save and new EDI and add this to the listbox
            else {

                edilist.Add(saveedi(new Edi(_polname.Text, _polname2.Text)));
                edilistbox.SelectedIndex = -1;
                selecteditem = -1;

                BindData();
            }



        }
        private Edi_Test.Edi saveedi(Edi edi2)
        {

            Edi edi = edi2;
            edi.Policynum = _polname.Text;

            edi.Poltype = _editype.Text;
            edi.Messagetype = _msgcount.Text;
            edi.Inceptiondate = dateTimePicker1.Text;
            edi.Enddate = dateTimePicker2.Text;
            edi.Firstline = _custadd.Text;
            edi.Secondline = _custadd2.Text;
            edi.Thirdline = _custadd3.Text;
            edi.Fourthline = _custadd4.Text;
            edi.Fifthline = _custadd5.Text;
            edi.Sixthline = _custadd6.Text;
            edi.Postcode = _custpost.Text;
            edi.Alarm = _secalarm.Text;
            edi.Premium = _totalprem.Text;

            edi.Endcode = _endcode.Text;

            edi.custgender = _custgender.Text;
            edi.Addcustgender = _addcustgender.Text;

            edi.RiskFirstline = _riskaddress1.Text;
            edi.RiskSecondline = _riskaddress2.Text;
            edi.RiskThirdline = _riskaddress3.Text;
            edi.RiskFourthline = _riskaddress4.Text;
            edi.RiskFifthline = _riskaddress5.Text;
            edi.RiskSixthline = _riskaddress6.Text;

            edi.RiskPostcode = _riskpostcode.Text;


            edi.Custtitle = _custitle.Text;
            edi.Firstname = _fname.Text;
            edi.Lastname = _lname.Text;
            edi.Occupation = _occu.Text;
            edi.Maritialstatus = _marrstatus.Text;
            edi.Addmaritialstatus = _addmar.Text;
            edi.Employeesbusiness = comboBox2.Text;
            edi.Employmentcatagory = _emptype.Text;
            edi.Dob = _custdob.Text;

            edi.noofchildren = _Numberchild.Text;
            edi.noofadults = _NumberAdult.Text;

            edi.Addpolicyholder = checkBox1.Checked;

            edi.Addcusttitle = _addtitle.Text;
            edi.Addfirstname = _addfirst.Text;
            edi.Addlastname = _addlast.Text;
            edi.AddFirstline = _addcustadd.Text;
            edi.AddSecondline = _addcustadd2.Text;
            edi.AddThirdline = _addcustadd3.Text;
            edi.AddFourthline = _addcustadd4.Text;
            edi.AddFifthline = _addcustadd5.Text;
            edi.AddSixthline = _addcustadd6.Text;
            edi.addPostcode = _addcustpost.Text;
            edi.Addoccupation = _addocc.Text;
            edi.Addmaritialstatus = _addmar.Text;
            edi.Addemployeesbusiness = _addemp.Text;
            edi.Addemploymentcatagory = _addempcat.Text;
            edi.Adddob = _adddob.Text;
            edi.Relationship = _addrelat.Text;

            edi.Walls = _wallcon.Text;
            edi.Roof = _roofcon.Text;
            edi.Roofpercent = _flatroof.Value.ToString();
            edi.Buildtype = _proptype.Text;
            edi.Noofbed = _numbed.Text;
            edi.Datebuild = _datebuilt.Value.ToString("ddMMyyyy");
            edi.Movedate = _movdate.Value.ToString("ddMMyyyy");
            edi.Businessuse = _bususe.Text;
            edi.Ownership = _owner.Text;
            edi.Occupancystatus = _occstatus.Text;
            edi.Listed = _listedbuild.Text;

            edi.Buildcomp = buscomp.Text;
            edi.Buildratedyear = textBox8.Text;
            edi.Buildvol = _buildpremium.Text;
            edi.Buildsum = _buildingsum.Text;
            edi.Buildad = buildAD.Text;

            edi.Contcomp = concomp.Text;
            edi.Contratedyear = conyear.Text;
            edi.Contvol = textBox10.Text;
            edi.Contsum = _contentsum.Text;
            edi.Contad = _conAD.Text;

            edi.Unspecpremium = _unsvalue.Text;
            edi.Unspecvalue = _unsvalue.Text;

            edi.Outhomecover = comboBox1.Text;
            edi.Outhomecover2 = comboBox3.Text;
            edi.Outhomepremium = textBox1.Text;
            edi.Outhomepremium2 = textBox2.Text;

            edi.Inhomecover = comboBox4.Text;
            edi.Inhomecover2 = comboBox5.Text;

            edi.Inhomesum = textBox3.Text;
            edi.Inhomesum2 = textBox4.Text;

            edi.Concover = _contcov.Checked;
            edi.Buscover = _buildcov.Checked;

            edi.Pedalcount = comboBox6.Text;
            edi.Outhomecount = comboBox8.Text;
            edi.Inhomecount = comboBox7.Text;
            edi.Pedal1 = ppremium1.Text;
            edi.Pedal2 = ppremium2.Text;
            edi.Pedal3 = ppremium3.Text;
            edi.Pedal4 = ppremium4.Text;
            edi.Pedal5 = ppremium5.Text;

            edi.Endorsement = endocheck.Checked;

            return edi;


        }

        private void changeselecteditem()
        {
            try
            {
                _totalprem.Text = edilist[selecteditem].Premium;
                _polname.Text = edilist[selecteditem].Policynum;
                _polname2.Text = edilist[selecteditem].Policynum2;
                _editype.Text = edilist[selecteditem].Poltype;
                _msgcount.Text = edilist[selecteditem].Messagetype;
                dateTimePicker1.Value = DateTime.ParseExact(edilist[selecteditem].Inceptiondate, "ddMMyyyy", CultureInfo.InvariantCulture, DateTimeStyles.None);
                dateTimePicker2.Value = DateTime.ParseExact(edilist[selecteditem].Enddate, "ddMMyyyy", CultureInfo.InvariantCulture, DateTimeStyles.None);
                _custadd.Text = edilist[selecteditem].Firstline;
                _custadd2.Text = edilist[selecteditem].Secondline;
                _custadd3.Text = edilist[selecteditem].Thirdline;
                _custadd4.Text = edilist[selecteditem].Fourthline;
                _custadd5.Text = edilist[selecteditem].Fifthline;
                _custadd6.Text = edilist[selecteditem].Sixthline;
                _custpost.Text = edilist[selecteditem].Postcode;


                _riskaddress1.Text = edilist[selecteditem].RiskFirstline;
                _riskaddress2.Text = edilist[selecteditem].RiskSecondline;
                _riskaddress3.Text = edilist[selecteditem].RiskThirdline;
                _riskaddress4.Text = edilist[selecteditem].RiskFourthline;
                _riskaddress5.Text = edilist[selecteditem].RiskFifthline;
                _riskaddress6.Text = edilist[selecteditem].RiskSixthline;
                _riskpostcode.Text = edilist[selecteditem].RiskPostcode;


                _secalarm.Text = edilist[selecteditem].Alarm;

                _custitle.Text = edilist[selecteditem].Custtitle;
                _fname.Text = edilist[selecteditem].Firstname;
                _lname.Text = edilist[selecteditem].Lastname;
                _occu.Text = edilist[selecteditem].Occupation;
                _marrstatus.Text = edilist[selecteditem].Maritialstatus;
                comboBox2.Text = edilist[selecteditem].Employeesbusiness;
                _emptype.Text = edilist[selecteditem].Employmentcatagory;
                _custdob.Value = DateTime.ParseExact(edilist[selecteditem].Dob, "ddMMyyyy", CultureInfo.InvariantCulture, DateTimeStyles.None);
                _custgender.Text = edilist[selecteditem].custgender;
                checkBox1.Checked = edilist[selecteditem].Addpolicyholder;

                _endcode.Text = edilist[selecteditem].Endcode;
                if (checkBox1.Checked)
                {
                    _addtitle.Text = edilist[selecteditem].Addcusttitle;
                    _addfirst.Text = edilist[selecteditem].Addfirstname;
                    _addlast.Text = edilist[selecteditem].Addlastname;
                    _addocc.Text = edilist[selecteditem].Addoccupation;
                    _addmar.Text = edilist[selecteditem].Addmaritialstatus;
                    _addemp.Text = edilist[selecteditem].Addemployeesbusiness;
                    _addempcat.Text = edilist[selecteditem].Addemploymentcatagory;
                    _adddob.Value = DateTime.ParseExact(edilist[selecteditem].Adddob, "ddMMyyyy", CultureInfo.InvariantCulture, DateTimeStyles.None);
                    _addrelat.Text = edilist[selecteditem].Relationship;


                    _addcustadd.Text = edilist[selecteditem].AddFirstline;
                    _addcustadd2.Text = edilist[selecteditem].AddSecondline;
                    _addcustadd3.Text = edilist[selecteditem].AddThirdline;
                    _addcustadd4.Text = edilist[selecteditem].AddFourthline;
                    _addcustadd5.Text = edilist[selecteditem].AddFifthline;
                    _addcustadd6.Text = edilist[selecteditem].AddSixthline;
                    _addcustpost.Text = edilist[selecteditem].addPostcode;
                    _addcustgender.Text = edilist[selecteditem].Addcustgender;
                }
                else
                {
                    _addtitle.Text = "";
                    _addfirst.Text = "";
                    _addlast.Text = "";
                    _addocc.Text = "";
                    _addmar.Text = "";
                    _addemp.Text = "";
                    _addempcat.Text = "";
                    _addcustgender.Text = "";

                    _addcustadd.Text = "";
                    _addcustadd2.Text = "";
                    _addcustadd3.Text = "";
                    _addcustadd4.Text = "";
                    _addcustadd5.Text = "";
                    _addcustadd6.Text = "";
                    _addcustpost.Text = "";

                    _addrelat.Text = "";
                }

                _NumberAdult.Text = edilist[selecteditem].noofadults;
                _Numberchild.Text = edilist[selecteditem].noofchildren;

                _wallcon.Text = edilist[selecteditem].Walls;
                _roofcon.Text = edilist[selecteditem].Roof;
                _flatroof.Value = int.Parse(edilist[selecteditem].Roofpercent);
                _proptype.Text = edilist[selecteditem].Buildtype;
                _numbed.Text = edilist[selecteditem].Noofbed;
                if (tabcontrol)
                {
                    _datebuilt.Value = DateTime.ParseExact(edilist[selecteditem].Datebuild, "ddMMyyyy", CultureInfo.InvariantCulture, DateTimeStyles.None);
                    _movdate.Value = DateTime.ParseExact(edilist[selecteditem].Movedate, "ddMMyyyy", CultureInfo.InvariantCulture, DateTimeStyles.None);
                }
                else
                {
                    _datebuilt.Value = DateTime.ParseExact(today, "ddMMyyyy", CultureInfo.InvariantCulture, DateTimeStyles.None);
                    _movdate.Value = DateTime.ParseExact(today, "ddMMyyyy", CultureInfo.InvariantCulture, DateTimeStyles.None);
                }
                _bususe.Text = edilist[selecteditem].Businessuse;
                _owner.Text = edilist[selecteditem].Ownership;
                _occstatus.Text = edilist[selecteditem].Occupancystatus;


                buscomp.Text = edilist[selecteditem].Buildcomp;
                textBox8.Text = edilist[selecteditem].Buildratedyear;
                _buildpremium.Text = edilist[selecteditem].Buildvol;
                _buildingsum.Text = edilist[selecteditem].Buildsum;
                buildAD.Text = edilist[selecteditem].Buildad;

                concomp.Text = edilist[selecteditem].Contcomp;
                conyear.Text = edilist[selecteditem].Contratedyear;
                textBox10.Text = edilist[selecteditem].Contvol;
                _contentsum.Text = edilist[selecteditem].Contsum;
                _conAD.Text = edilist[selecteditem].Contad;

                _unsvalue.Text = edilist[selecteditem].Unspecpremium;
                _unsvalue.Text = edilist[selecteditem].Unspecvalue;

                comboBox1.Text = edilist[selecteditem].Outhomecover;
                comboBox3.Text = edilist[selecteditem].Outhomecover2;
                textBox1.Text = edilist[selecteditem].Outhomepremium;
                textBox2.Text = edilist[selecteditem].Outhomepremium2;

                comboBox4.Text = edilist[selecteditem].Inhomecover;
                comboBox5.Text = edilist[selecteditem].Inhomecover2;

                textBox3.Text = edilist[selecteditem].Inhomesum;
                textBox4.Text = edilist[selecteditem].Inhomesum2;

                _contcov.Checked = edilist[selecteditem].Concover;
                _buildcov.Checked = edilist[selecteditem].Buscover;


                endocheck.Checked = edilist[selecteditem].Endorsement;
                _endcode.Text = edilist[selecteditem].Endcode;

                comboBox8.Text = edilist[selecteditem].Outhomecount;
                comboBox7.Text = edilist[selecteditem].Inhomecount;
                comboBox6.Text = edilist[selecteditem].Pedalcount;
                ppremium1.Text = edilist[selecteditem].Pedal1;
                ppremium2.Text = edilist[selecteditem].Pedal2;
                ppremium3.Text = edilist[selecteditem].Pedal3;
                ppremium4.Text = edilist[selecteditem].Pedal4;
                ppremium5.Text = edilist[selecteditem].Pedal5;

            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message);
            }
        }


        private void edilistbox_SelectedIndexChanged(object sender, EventArgs e)
        {

            if (edilistbox.SelectedIndex > -1)
            {
                selecteditem = edilistbox.SelectedIndex;
                changeselecteditem();
            }




        }

        private DataTable dataset(DataTable table)
        {

            DataTable NewTable = table;

            NewTable.Columns.Add("_policynum");
            NewTable.Columns.Add("_policynum2");
            NewTable.Columns.Add("_poltype");
            NewTable.Columns.Add("_messagetype");
            NewTable.Columns.Add("_inceptiondate");
            NewTable.Columns.Add("_enddate");
            NewTable.Columns.Add("_firstline");
            NewTable.Columns.Add("_secondline");
            NewTable.Columns.Add("_thirdline");
            NewTable.Columns.Add("_fourthline");
            NewTable.Columns.Add("_fifthline");
            NewTable.Columns.Add("_sixthline");
            NewTable.Columns.Add("_postcode");

            NewTable.Columns.Add("_riskfirstline");
            NewTable.Columns.Add("_risksecondline");
            NewTable.Columns.Add("_riskthirdline");
            NewTable.Columns.Add("_riskfourthline");
            NewTable.Columns.Add("_riskfifthline");
            NewTable.Columns.Add("_risksixthline");
            NewTable.Columns.Add("_riskpostcode");

            NewTable.Columns.Add("_custtitle");
            NewTable.Columns.Add("_firstname");
            NewTable.Columns.Add("_lastname");
            NewTable.Columns.Add("_occupation");
            NewTable.Columns.Add("_maritialstatus");
            NewTable.Columns.Add("_employeesbusiness");
            NewTable.Columns.Add("_employmentcatagory");
            NewTable.Columns.Add("_dob");
            NewTable.Columns.Add("_custgender");
            NewTable.Columns.Add("_addpolicyholder");
            NewTable.Columns.Add("_addcusttitle");
            NewTable.Columns.Add("_addfirstname");
            NewTable.Columns.Add("_addlastname");
            NewTable.Columns.Add("_addfirstline");
            NewTable.Columns.Add("_addsecondline");
            NewTable.Columns.Add("_addthirdline");
            NewTable.Columns.Add("_addfourthline");
            NewTable.Columns.Add("_addfifthline");
            NewTable.Columns.Add("_addsixthline");
            NewTable.Columns.Add("_addpostcode");
            NewTable.Columns.Add("_addoccupation");
            NewTable.Columns.Add("_addmaritialstatus");
            NewTable.Columns.Add("_addemployeesbusiness");
            NewTable.Columns.Add("_addemploymentcatagory");
            NewTable.Columns.Add("_adddob");
            NewTable.Columns.Add("_addcustgender");
            NewTable.Columns.Add("_relationship");
            NewTable.Columns.Add("_walls");
            NewTable.Columns.Add("_roof");
            NewTable.Columns.Add("_listed");
            NewTable.Columns.Add("_roofpercent");
            NewTable.Columns.Add("_buildtype");
            NewTable.Columns.Add("_noofbed");
            NewTable.Columns.Add("_datebuild");
            NewTable.Columns.Add("_movedate");
            NewTable.Columns.Add("_businessuse");
            NewTable.Columns.Add("_ownership");
            NewTable.Columns.Add("_occupancystatus");
            NewTable.Columns.Add("_buildcomp");
            NewTable.Columns.Add("_buildratedyear");
            NewTable.Columns.Add("_buildvol");
            NewTable.Columns.Add("_buildsum");
            NewTable.Columns.Add("_buildad");
            NewTable.Columns.Add("_contcomp");
            NewTable.Columns.Add("_contratedyear");
            NewTable.Columns.Add("_contvol");
            NewTable.Columns.Add("_contsum");
            NewTable.Columns.Add("_contad");
            NewTable.Columns.Add("_unspecvalue");
            NewTable.Columns.Add("_premium");
            NewTable.Columns.Add("_outhomecover");
            NewTable.Columns.Add("_outhomecover2");
            NewTable.Columns.Add("_outofhomepremium");
            NewTable.Columns.Add("_outofhomepremium2");
            NewTable.Columns.Add("_inhomecover");
            NewTable.Columns.Add("_inhomecover2");
            NewTable.Columns.Add("_inhomesum");
            NewTable.Columns.Add("_inhomesum2");
            NewTable.Columns.Add("_security");

            NewTable.Columns.Add("_concover");
            NewTable.Columns.Add("_buscover");

            NewTable.Columns.Add("_outhomecount");
            NewTable.Columns.Add("_inhomecount");
            NewTable.Columns.Add("_pedalcount");
            NewTable.Columns.Add("_pedal1");
            NewTable.Columns.Add("_pedal2");
            NewTable.Columns.Add("_pedal3");
            NewTable.Columns.Add("_pedal4");
            NewTable.Columns.Add("_pedal5");
            NewTable.Columns.Add("_endorsement");
            NewTable.Columns.Add("_endorsecode");

            NewTable.TableName = "Sheet1";

            foreach (Edi edi in edilist)
            {
                if (edi.Addpolicyholder == false)
                {
                    policy = "N";
                }
                if (edi.Addpolicyholder == true)
                {
                    policy = "Y";
                }

                if (edi.Concover == false)
                {
                    concover = "N";
                }
                if (edi.Concover == true)
                {
                    concover = "Y";
                }

                if (edi.Buscover == false)
                {
                    buscover = "N";
                }
                if (edi.Buscover == true)
                {
                    buscover = "Y";
                }


                if (edi.Endorsement == false)
                {
                    endorsement = "N";
                }
                if (edi.Endorsement == true)
                {
                    endorsement = "Y";
                }

                table.Rows.Add(edi.Policynum, edi.Policynum2, edi.Poltype, edi.Messagetype, edi.Inceptiondate, edi.Enddate, edi.Firstline, edi.Secondline, edi.Thirdline, edi.Fourthline, edi.Fifthline, edi.Sixthline, edi.Postcode, edi.RiskFirstline, edi.RiskSecondline, edi.RiskThirdline, edi.RiskFourthline, edi.RiskFifthline, edi.RiskSixthline, edi.RiskPostcode, edi.Custtitle, edi.Firstname, edi.Lastname, edi.Occupation, edi.Maritialstatus, edi.Employeesbusiness, edi.Employmentcatagory, edi.Dob, edi.custgender, policy, edi.Addcusttitle, edi.Addfirstname, edi.Addlastname, edi.AddFirstline, edi.AddSecondline, edi.AddThirdline, edi.AddFourthline, edi.AddFifthline, edi.AddSixthline, edi.addPostcode, edi.Addoccupation, edi.Addmaritialstatus, edi.Addemployeesbusiness, edi.Addemploymentcatagory, edi.Adddob, edi.Addcustgender, edi.Relationship, edi.Walls, edi.Roof, edi.Listed, edi.Roofpercent, edi.Buildtype, edi.Noofbed, edi.Datebuild, edi.Movedate, edi.Businessuse, edi.Ownership, edi.Occupancystatus, edi.Buildcomp, edi.Buildratedyear, edi.Buildvol, edi.Buildsum, edi.Buildad, edi.Contcomp, edi.Contratedyear, edi.Contvol, edi.Contsum, edi.Contad, edi.Unspecvalue, edi.Premium, edi.Outhomecover, edi.Outhomecover2, edi.Outhomepremium, edi.Outhomepremium2, edi.Inhomecover, edi.Inhomecover2, edi.Inhomesum, edi.Inhomesum2, edi.Alarm, concover, buscover, edi.Outhomecount, edi.Inhomecount, edi.Pedalcount, edi.Pedal1, edi.Pedal2, edi.Pedal3, edi.Pedal4, edi.Pedal5, endorsement, edi.Endcode);

            }

            return NewTable;

        }

        private void tab1Control_SelectedIndexChanged(object sender, EventArgs e)
        {
            if (CustomerDetails.SelectedTab == CustomerDetails.TabPages["Insurable"])//your specific tabname
            {
                tabcontrol = true;
            }
        }
        protected virtual bool IsFileLocked(OpenFileDialog file)
        {
            FileStream stream = null;

            try
            {
                stream = File.Open(file.FileName, FileMode.Open, FileAccess.Read);
            }
            catch (IOException)
            {
                //the file is unavailable because it is:
                //still being written to
                //or being processed by another thread
                //or does not exist (has already been processed)
                return true;
            }
            finally
            {
                if (stream != null)
                    stream.Close();
            }

            //file is not locked
            return false;
        }
        private void button2_Click(object sender, EventArgs e)
        {
            try
            {

                OpenFileDialog openFileDialog1 = new OpenFileDialog();
                openFileDialog1.Filter = "Excel File|*.xlsx";
                openFileDialog1.Title = "Open an Excel File";
                openFileDialog1.ShowDialog();



                if (IsFileLocked(openFileDialog1))
                {
                    MessageBox.Show("File is currently in use, please close the file");
                }
                else
                {
                    FileStream stream = File.Open(openFileDialog1.FileName, FileMode.Open, FileAccess.Read);

                    //1. Reading from a binary Excel file ('97-2003 format; *.xls)
                    // IExcelDataReader excelReader = ExcelReaderFactory.CreateBinaryReader(stream);

                    //2. Reading from a OpenXml Excel file (2007 format; *.xlsx)
                    IExcelDataReader excelReader = ExcelReaderFactory.CreateOpenXmlReader(stream);

                    //3. DataSet - The result of each spreadsheet will be created in the result.Tables
                    //      DataSet result = excelReader.AsDataSet();

                    //4. DataSet - Create column names from first row
                    excelReader.IsFirstRowAsColumnNames = true;
                    DataSet result = excelReader.AsDataSet();

                    //5. Data Reader methods
                    while (excelReader.Read())
                    {
                        //excelReader.GetInt32(0);

                    }

                    //6. Free resources (IExcelDataReader is IDisposable)
                    excelReader.Close();


                    foreach (DataRow dr in result.Tables["Sheet1"].Rows)
                    {
                        if ((dr["_policynum"].ToString() != ""))
                        {
                            Edi edi = new Edi((dr["_policynum"].ToString()), (dr["_policynum2"].ToString()));

                            edi.Policynum = dr["_policynum"].ToString();

                            edi.Poltype = dr["_poltype"].ToString();
                            edi.Messagetype = dr["_messagetype"].ToString();
                            edi.Inceptiondate = dr["_inceptiondate"].ToString();
                            edi.Enddate = dr["_enddate"].ToString();
                            edi.Firstline = dr["_firstline"].ToString();
                            edi.Secondline = dr["_secondline"].ToString();
                            edi.Thirdline = dr["_thirdline"].ToString();
                            edi.Fourthline = dr["_fourthline"].ToString();
                            edi.Fifthline = dr["_fifthline"].ToString();
                            edi.Sixthline = dr["_sixthline"].ToString();
                            edi.Postcode = dr["_postcode"].ToString();

                            edi.RiskFirstline = dr["_riskfirstline"].ToString();
                            edi.RiskSecondline = dr["_risksecondline"].ToString();
                            edi.RiskThirdline = dr["_riskthirdline"].ToString();
                            edi.RiskFourthline = dr["_riskfourthline"].ToString();
                            edi.RiskFifthline = dr["_riskfifthline"].ToString();
                            edi.RiskSixthline = dr["_risksixthline"].ToString();
                            edi.RiskPostcode = dr["_riskpostcode"].ToString();


                            edi.Custtitle = dr["_custtitle"].ToString();
                            edi.Firstname = dr["_firstname"].ToString();
                            edi.Lastname = dr["_lastname"].ToString();
                            edi.Occupation = dr["_occupation"].ToString();
                            edi.Maritialstatus = dr["_maritialstatus"].ToString();
                            edi.Employeesbusiness = dr["_employeesbusiness"].ToString();
                            edi.Employmentcatagory = dr["_employmentcatagory"].ToString();
                            edi.Dob = dr["_dob"].ToString();
                            edi.custgender = dr["_custgender"].ToString();

                            if (dr["_addpolicyholder"].ToString() == "Y")
                            {
                                edi.Addpolicyholder = true;
                            }
                            else
                            {
                                edi.Addpolicyholder = false;
                            }



                            edi.Addcusttitle = dr["_addcusttitle"].ToString();
                            edi.Addfirstname = dr["_addfirstname"].ToString();
                            edi.Addlastname = dr["_addlastname"].ToString();
                            edi.AddFirstline = dr["_addfirstline"].ToString();
                            edi.AddSecondline = dr["_addsecondline"].ToString();
                            edi.AddThirdline = dr["_addthirdline"].ToString();
                            edi.AddFourthline = dr["_addfourthline"].ToString();
                            edi.AddFifthline = dr["_addfifthline"].ToString();
                            edi.AddSixthline = dr["_addsixthline"].ToString();
                            edi.addPostcode = dr["_addpostcode"].ToString();
                            edi.Addoccupation = dr["_addoccupation"].ToString();
                            edi.Addmaritialstatus = dr["_addmaritialstatus"].ToString();
                            edi.Addemployeesbusiness = dr["_addemployeesbusiness"].ToString();
                            edi.Addemploymentcatagory = dr["_addemploymentcatagory"].ToString();
                            edi.Adddob = dr["_adddob"].ToString();
                            edi.Addcustgender = dr["_addcustgender"].ToString();
                            edi.Relationship = dr["_relationship"].ToString();

                            edi.Walls = dr["_walls"].ToString();
                            edi.Roof = dr["_roof"].ToString();
                            edi.Roofpercent = dr["_roofpercent"].ToString();
                            edi.Listed = dr["_listed"].ToString();
                            edi.Buildtype = dr["_buildtype"].ToString();
                            edi.Noofbed = dr["_noofbed"].ToString();
                            edi.Datebuild = dr["_datebuild"].ToString();
                            edi.Movedate = dr["_movedate"].ToString();
                            edi.Businessuse = dr["_businessuse"].ToString();
                            edi.Ownership = dr["_ownership"].ToString();
                            edi.Occupancystatus = dr["_occupancystatus"].ToString();

                            edi.Buildcomp = dr["_buildcomp"].ToString();
                            edi.Buildratedyear = dr["_buildratedyear"].ToString();
                            edi.Buildvol = dr["_buildvol"].ToString();
                            edi.Buildsum = dr["_buildsum"].ToString();
                            edi.Buildad = dr["_buildad"].ToString();

                            edi.Contcomp = dr["_contcomp"].ToString();
                            edi.Contratedyear = dr["_contratedyear"].ToString();
                            edi.Contvol = dr["_contvol"].ToString();
                            edi.Contsum = dr["_contsum"].ToString();
                            edi.Contad = dr["_contad"].ToString();

                            edi.Premium = dr["_premium"].ToString();
                            edi.Unspecvalue = dr["_unspecvalue"].ToString();

                            edi.Outhomecover = dr["_outhomecover"].ToString();
                            edi.Outhomecover2 = dr["_outhomecover2"].ToString();
                            edi.Outhomepremium = dr["_outofhomepremium"].ToString();
                            edi.Outhomepremium2 = dr["_outofhomepremium2"].ToString();
                            edi.Inhomecover = dr["_inhomecover"].ToString();
                            edi.Inhomecover2 = dr["_inhomecover2"].ToString();

                            edi.Inhomesum = dr["_inhomesum"].ToString();
                            edi.Inhomesum2 = dr["_inhomesum2"].ToString();

                            edi.Alarm = dr["_security"].ToString();


                            edi.Outhomecount = dr["_outhomecount"].ToString();
                            edi.Inhomecount = dr["_inhomecount"].ToString();
                            edi.Pedalcount = dr["_pedalcount"].ToString();
                            edi.Pedal1 = dr["_pedal1"].ToString();
                            edi.Pedal2 = dr["_pedal2"].ToString();
                            edi.Pedal3 = dr["_pedal3"].ToString();
                            edi.Pedal4 = dr["_pedal4"].ToString();
                            edi.Pedal5 = dr["_pedal5"].ToString();
                            edi.Endcode = dr["_endorsecode"].ToString();


                            if (dr["_concover"].ToString() == "Y")
                            {
                                edi.Concover = true;
                            }
                            else
                            {
                                edi.Concover = false;
                            }

                            if (dr["_buscover"].ToString() == "Y")
                            {
                                edi.Buscover = true;
                            }
                            else
                            {
                                edi.Buscover = false;
                            }

                            if (dr["_endorsement"].ToString() == "Y")
                            {
                                edi.Endorsement = true;
                            }
                            else
                            {
                                edi.Endorsement = false;
                            }

                            edilistbox.SelectedIndex = -1;
                            edilist.Add(edi);
                            BindData();
                        }
                    }
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message);
            }

        }

        private void _delete_Click(object sender, EventArgs e)
        {
            if (selecteditem > -1)
            {
                edilist.RemoveAt(selecteditem);
                selecteditem = -1;
                BindData();
            }
        }

        private void button4_Click(object sender, EventArgs e)
        {


            try
            {

                SaveFileDialog exportsave = new SaveFileDialog();
                exportsave.DefaultExt = "xlsx";
                exportsave.Filter = "Excel file (.xlsx)|*.xlsx";



                if (exportsave.ShowDialog() == DialogResult.OK)
                {
                    DataTable table = new DataTable();
                    table = dataset(table);

                    Microsoft.Office.Interop.Excel.Application excelApp = new Microsoft.Office.Interop.Excel.Application();
                    excelApp.Visible = false;
                    excelApp.DisplayAlerts = false;
                    //Create an Excel workbook instance and open it from the predefined location
                    Microsoft.Office.Interop.Excel.Workbook excelWorkBook = excelApp.Workbooks.Add(Type.Missing);

                    //Add a new worksheet to workbook with the Datatable name
                    Microsoft.Office.Interop.Excel.Worksheet excelWorkSheet = (Microsoft.Office.Interop.Excel.Worksheet)excelWorkBook.ActiveSheet;
                    excelWorkSheet.Name = table.TableName;

                    for (int i = 1; i < table.Columns.Count + 1; i++)
                    {
                        excelWorkSheet.Cells[1, i] = table.Columns[i - 1].ColumnName;
                    }

                    for (int j = 0; j < table.Rows.Count; j++)
                    {
                        for (int k = 0; k < table.Columns.Count; k++)
                        {

                            excelWorkSheet.Cells[j + 2, k + 1].Value = table.Rows[j].ItemArray[k].ToString();
                            excelWorkSheet.Cells[j + 2, 5].NumberFormat = "@";
                            excelWorkSheet.Cells[j + 2, 6].NumberFormat = "@";
                            excelWorkSheet.Cells[j + 2, 28].NumberFormat = "@";
                            excelWorkSheet.Cells[j + 2, 45].NumberFormat = "@";
                            excelWorkSheet.Cells[j + 2, 54].NumberFormat = "@";
                            excelWorkSheet.Cells[j + 2, 55].NumberFormat = "@";
                        }
                    }



                    //     excelWorkSheet.Cells[1, 1] = "Sample test data";
                    //  excelWorkSheet.Cells[1, 2] = "Date : " + DateTime.Now.ToShortDateString();

                    excelWorkBook.SaveAs(exportsave.FileName);
                    excelWorkBook.Close();
                    excelApp.Quit();

                    NAR(excelWorkBook);
                    NAR(excelWorkSheet);
                    NAR(excelApp);
                    GC.Collect();
                    GC.WaitForPendingFinalizers();

                    MessageBox.Show("Export Complete");

                }

            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message);
            }

        }

        private void NAR(object o)
        {
            try
            {
                while (System.Runtime.InteropServices.Marshal.ReleaseComObject(o) > 0) ;
            }
            catch { }
            finally
            {
                o = null;
            }
        }

        private void button5_Click(object sender, EventArgs e)
        {
            myConnection.Close();
            try
            {

                OpenFileDialog openFileDialog1 = new OpenFileDialog();
                openFileDialog1.Filter = "EDI|*.edi|Txt|*.txt";
                openFileDialog1.Title = "Open an EDI file";
                openFileDialog1.ShowDialog();



                if (IsFileLocked(openFileDialog1))
                {
                    MessageBox.Show("File is currently in use, please close the file");
                }
                else
                {
                    Edi edi = new Edi();
                    string[] lines = System.IO.File.ReadAllLines(openFileDialog1.FileName);
                    int rcacount = 0;
                    int procount = 0;
                    int plocount = 0;
                    foreach (string line in lines)
                    {


                        if (line.Contains("PLH="))
                        {
                            string policynum = line.Split('=', '+')[1];

                            string[] result = Regex.Matches(policynum, @"[a-zA-Z]+|\d+")
                       .Cast<Match>()
                       .Select(x => x.Value)
                       .ToArray(); //"AAAA" "000343" "BBB" "343"
                            if (result.Length > 4)
                            {
                                edi.Policynum = result[0] + result[1] + result[2] + result[3];
                                edi.Policynum2 = result[4] + result[5];
                            }
                            else {
                                edi.Policynum = result[0] + result[1];
                                edi.Policynum2 = result[2] + result[3];
                            }
                            string poltype = line.Split(':', ':')[2];
                            switch (poltype)
                            {
                                case "PRO":
                                    edi.Poltype = "New Business";
                                    break;
                                case "RNC":
                                    edi.Poltype = "Renewal";
                                    break;
                                case "RCA":
                                    edi.Poltype = "Renewal";
                                    break;
                                case "ADJ":
                                    edi.Poltype = "Adjustment";
                                    break;
                                case "PLL":
                                    edi.Poltype = "Lapse";
                                    break;
                                case "CAB":
                                    edi.Poltype = "Cancellation";
                                    break;


                            }
                            edi.Messagetype = line.Split('+', ':')[4];
                            edi.Inceptiondate = "08122016"; //line.Split('+', ':')[6];


                        }

                        if (line.Contains("POL="))
                        {
                            edi.Enddate = line.Split('+', ':')[2];
                        }
                        if (line.Contains("PRO="))
                        {
                            if (procount == 0)
                            {
                                edi.Occupation = line.Split(':', ':')[2];
                                edi.Employeesbusiness = line.Split(':', '+', '\'')[6];
                                edi.Maritialstatus = RetrieveMaritalStatus(line.Split('=', '+')[1]);
                                edi.Employmentcatagory = RetrieveEmploymentType(line.Split('+', ':')[1]);

                                procount++;
                            }
                            else
                            {
                                edi.Addoccupation = line.Split(':', ':')[2];
                                edi.Addemployeesbusiness = line.Split(':', '+', '\'')[6];
                                edi.Addmaritialstatus = RetrieveMaritalStatus(line.Split('=', '+')[1]);
                                edi.Addemploymentcatagory = RetrieveEmploymentType(line.Split('+', ':')[1]);
                            }
                        }
                        if (line.Contains("RDC="))
                        {

                            edi.Walls = RetrieveWalls(line.Split('=', ':')[1]);
                            edi.Roof = RetrieveRoof(line.Split(':', '+')[1]);
                            //  edi.Roofpercent = line.Split(':', '+')[3];
                            edi.Roofpercent = "0";

                        }
                        if (line.Contains("RDD="))
                        {
                            edi.Ownership = RetrieveOwnership(line.Split('+', ':')[1]);
                            edi.Noofbed = line.Split('+', '+')[9];
                            edi.Datebuild = line.Split('+', '\'')[13];


                            //  edi.Datebuild = ;#
                            try
                            {
                                edi.Buildtype = RetrieveType(line.Split('+')[2]);
                            }
                            catch
                            {
                                myConnection.Close();
                                // edi.Ownership = line.Split('+', ':')[2];
                                edi.Buildtype = RetrieveType(line.Split('+', ':')[3]);
                            }
                        }
                        if ((line.Contains("PID=01")) || (line.Contains("PID=1")))
                        {

                            edi.Custtitle = line.Split(':', '+')[5];
                            edi.Firstname = line.Split(':', '+')[3];
                            edi.Lastname = line.Split('+', ':')[2];
                            edi.Dob = line.Split('+', '\'')[5];
                        }

                        if ((line.Contains("PID=02")) || (line.Contains("PID=2")))
                        {
                            edi.Addpolicyholder = true;
                            edi.Addcusttitle = line.Split(':', '+')[5];
                            edi.Addfirstname = line.Split(':', '+')[3];
                            edi.Addlastname = line.Split('+', ':')[2];
                            edi.Adddob = line.Split('+', '\'')[5];
                        }

                        if (line.Contains("PLO="))
                        {
                            if (plocount == 0)
                            {
                                edi.Firstline = line.Split('=', '+')[1];
                                edi.Postcode = line.Split('+', '+')[1];


                                plocount++;
                            }
                            else
                            {
                                string addresss2 = line.Split('=', ':')[1];
                                string postcode2 = line.Split('+', '+')[1];
                            }

                        }

                        if (line.Contains("RCA=01") || line.Contains("RCA=1"))
                        {
                            edi.Buildsum = line.Split(':', '+')[4];
                            edi.Buildad = line.Split(':', '+')[5];
                            edi.Businessuse = line.Split(':', '+')[6];
                            edi.Buildvol = line.Split('+', '+')[6];
                        }
                        if (line.Contains("LDG="))
                        {
                            if (rcacount == 0)
                            {
                                edi.Buildpremium = line.Split('+', '+')[4];
                                rcacount++;
                            }
                            if (rcacount == 1)
                            {

                            }
                            if (rcacount == 2)
                            {

                            }
                        }


                    }

                    // Edi edi = new Edi((dr["_policynum"].ToString()), (dr["_policynum2"].ToString()));
                    edilistbox.SelectedIndex = -1;
                    edilist.Add(edi);
                    BindData();
                    MessageBox.Show("Successful import");
                }

            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message);

            }
        }

        private void groupBox12_Enter(object sender, EventArgs e)
        {

        }

        private void checkBox1_CheckedChanged_1(object sender, EventArgs e)
        {
            if (checkBox1.Checked)
            {
                _addtitle.Enabled = true;
                _addfirst.Enabled = true;
                _addlast.Enabled = true;
                _addocc.Enabled = true;
                _addempcat.Enabled = true;
                _addemp.Enabled = true;
                _adddob.Enabled = true;
                _addmar.Enabled = true;
                _addrelat.Enabled = true;

                _addcustadd.Enabled = true;
                _addcustadd2.Enabled = true;
                _addcustadd3.Enabled = true;
                _addcustadd4.Enabled = true;
                _addcustadd5.Enabled = true;
                _addcustadd6.Enabled = true;
                _addcustgender.Enabled = true;
                _addcustpost.Enabled = true;
            }
            else {
                _addtitle.Enabled = false;
                _addfirst.Enabled = false;
                _addlast.Enabled = false;
                _addocc.Enabled = false;
                _addempcat.Enabled = false;
                _addemp.Enabled = false;
                _adddob.Enabled = false;
                _addmar.Enabled = false;
                _addrelat.Enabled = false;

                _addcustadd.Enabled = false;
                _addcustadd2.Enabled = false;
                _addcustadd3.Enabled = false;
                _addcustadd4.Enabled = false;
                _addcustadd5.Enabled = false;
                _addcustadd6.Enabled = false;
                _addcustgender.Enabled = false;
                _addcustpost.Enabled = false;

            }
        }

        private void checkBox2_CheckedChanged(object sender, EventArgs e)
        {
            if (_buildcov.Checked)
            {
                buscomp.Enabled = false;
                textBox8.Enabled = false;
                _buildpremium.Enabled = false;
                _buildingsum.Enabled = false;
                buildAD.Enabled = false;
            }
            else
            {
                buscomp.Enabled = true;
                textBox8.Enabled = true;
                _buildpremium.Enabled = true;
                _buildingsum.Enabled = true;
                buildAD.Enabled = true;
            }
        }

        private void label46_Click_1(object sender, EventArgs e)
        {

        }

        private void buildAD_SelectedIndexChanged(object sender, EventArgs e)
        {

        }

        private void bpremium_ValueChanged(object sender, EventArgs e)
        {

        }

        private void label39_Click(object sender, EventArgs e)
        {

        }

        private void textBox8_TextChanged(object sender, EventArgs e)
        {

        }

        private void label19_Click(object sender, EventArgs e)
        {

        }

        private void _contcov_CheckedChanged(object sender, EventArgs e)
        {
            if (_contcov.Checked)
            {
                concomp.Enabled = false;
                conyear.Enabled = false;
                textBox10.Enabled = false;
                _contentsum.Enabled = false;
                _conAD.Enabled = false;
            }

            else
            {

                concomp.Enabled = true;
                conyear.Enabled = true;
                textBox10.Enabled = true;
                _contentsum.Enabled = true;
                _conAD.Enabled = true;

            }
        }



        private void ppremium5_ValueChanged(object sender, EventArgs e)
        {

        }

        private void endocheck_CheckedChanged(object sender, EventArgs e)
        {
            if (endocheck.Checked)
            {
                _endcode.Enabled = false;
            }

            else
            {
                _endcode.Enabled = true;
            }
        }

        private void comboBox6_SelectedIndexChanged(object sender, EventArgs e)
        {
            switch (comboBox6.Text)
            {
                case "0":
                    ppremium1.Enabled = false;
                    ppremium2.Enabled = false;
                    ppremium3.Enabled = false;
                    ppremium4.Enabled = false;
                    ppremium5.Enabled = false;
                    break;
                case "1":
                    ppremium1.Enabled = true;
                    ppremium2.Enabled = false;
                    ppremium3.Enabled = false;
                    ppremium4.Enabled = false;
                    ppremium5.Enabled = false;
                    break;
                case "2":
                    ppremium1.Enabled = true;
                    ppremium2.Enabled = true;
                    ppremium3.Enabled = false;
                    ppremium4.Enabled = false;
                    ppremium5.Enabled = false;
                    break;
                case "3":
                    ppremium1.Enabled = true;
                    ppremium2.Enabled = true;
                    ppremium3.Enabled = true;
                    ppremium4.Enabled = false;
                    ppremium5.Enabled = false;
                    break;
                case "4":
                    ppremium1.Enabled = true;
                    ppremium2.Enabled = true;
                    ppremium3.Enabled = true;
                    ppremium4.Enabled = true;
                    ppremium5.Enabled = false;
                    break;
                case "5":
                    ppremium1.Enabled = true;
                    ppremium2.Enabled = true;
                    ppremium3.Enabled = true;
                    ppremium4.Enabled = true;
                    ppremium5.Enabled = true;
                    break;
                default:
                    Console.WriteLine("Default case");
                    break;
            }

        }

        private void _exportxml_Click(object sender, EventArgs e)
        {
            MessageBox.Show("Feature currently in development");
        }

        private void label4_Click(object sender, EventArgs e)
        {

        }

        private void label23_Click(object sender, EventArgs e)
        {

        }

        private void comboBox8_SelectedIndexChanged_1(object sender, EventArgs e)
        {
            switch (int.Parse(comboBox8.Text))
            {
                case 0:
                    comboBox1.Enabled = false;
                    comboBox3.Enabled = false;
                    textBox1.Enabled = false;
                    textBox2.Enabled = false;
                    break;
                case 1:
                    comboBox1.Enabled = true;
                    comboBox3.Enabled = false;
                    textBox1.Enabled = true;
                    textBox2.Enabled = false;
                    break;
                case 2:
                    comboBox1.Enabled = true;
                    comboBox3.Enabled = true;
                    textBox1.Enabled = true;
                    textBox2.Enabled = true;
                    break;
                default:
                    break;


            }
        }

        private void comboBox7_SelectedIndexChanged(object sender, EventArgs e)
        {
            switch (int.Parse(comboBox7.Text))
            {
                case 0:
                    comboBox5.Enabled = false;
                    comboBox4.Enabled = false;
                    textBox3.Enabled = false;
                    textBox4.Enabled = false;
                    break;
                case 1:
                    comboBox5.Enabled = true;
                    comboBox4.Enabled = false;
                    textBox3.Enabled = true;
                    textBox4.Enabled = false;
                    break;
                case 2:
                    comboBox5.Enabled = true;
                    comboBox4.Enabled = true;
                    textBox3.Enabled = true;
                    textBox4.Enabled = true;
                    break;
                default:
                    break;

            }
        }
    }
}
