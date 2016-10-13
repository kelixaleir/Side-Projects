using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using System.Data.SqlClient;
using System.Globalization;
using System.Collections;
using System.Text.RegularExpressions;


namespace Flood
{
    public partial class Form1 : Form
    {
        //Create a list of policy objects from the policy class
        List<house> Houselist = new List<house>();

        //Sql connection string to server
        SqlConnection myConnection = new SqlConnection("Data Source=10.1.3.17,1433;Network Library=DBMSSOCN; Initial Catalog=TEST_EDIDB;User ID=EDIDBUser;Password=WlhbR2WeUj;");
        //Selected item is defaulted at -1
        int selecteditem = -1;
        public Form1()
        {
            InitializeComponent();
        }

        private void groupBox2_Enter(object sender, EventArgs e)
        {

        }

        void BindData()
        {
      
            //Binds the policycontent list to the listbox on the GUI and shows the Eventtype name (e.g ADJ, CAB, Etc.)
            listBox1.DataSource = null;
            listBox1.DataSource = Houselist;
            listBox1.DisplayMember = "Number";
            Houselist.OrderBy(x => x.Number);


        }


        private DataTable FindID(string cmdtext, string sort)
        {
            
                //Create a new datatable and open the SQL connection
                DataTable dt = new DataTable();
                myConnection.Open();

                //Run SQL query to find information based on whats in textbox1 (Policy Number)
                SqlCommand cmd = myConnection.CreateCommand();
                cmd.CommandText = cmdtext;
                //Fill datatable with this information
                using (var da = new SqlDataAdapter(cmd.CommandText, myConnection))
                {
                    da.Fill(dt);
                }
                //      myConnection.Close();
                //Return the datatable
                DataView dv = dt.DefaultView;
                dv.Sort = sort;
                DataTable sortedDT = dv.ToTable();
                return sortedDT;
            }

        



        private void select()
        {
            //Update the textboxes on the GUI with the policy information that you click on in the list box
            textBox1.Text = Houselist[listBox1.SelectedIndex].Band;
            textBox2.Text = Houselist[listBox1.SelectedIndex].Name + " " + Houselist[listBox1.SelectedIndex].Number + " " + Houselist[listBox1.SelectedIndex].SS1;
            textBox3.Text = Houselist[listBox1.SelectedIndex].Town;
            textBox5.Text = Houselist[listBox1.SelectedIndex].Street;
            textBox6.Text = Houselist[listBox1.SelectedIndex].Pre;
            if (Houselist[listBox1.SelectedIndex].Pre == "1")
               {
                string temppost = Houselist[listBox1.SelectedIndex].Postcode;
                string alphabet = String.Empty;
                string digit = String.Empty;

                Match regexMatch = Regex.Match(temppost, "\\d");
                if (regexMatch.Success) //Found numeric part in the coverage string
                {
                    int digitStartIndex = regexMatch.Index; //Get the index where the numeric digit found
                    alphabet = temppost.Substring(0, digitStartIndex);
                    digit = temppost.Substring(digitStartIndex);
                }
                Calculate(Houselist[listBox1.SelectedIndex].Band, findcountry(alphabet));
            }

            else
            {
                textBox7.Text = "";
                textBox8.Text = "";
                textBox9.Text = "";
            }
        }

        private void Calculate(string band, string country)
        {
            int _combined = 0;
            int _building = 0;
            int _contents = 0;

            switch (country)
            {
                case "England":
                    switch (band)
                    {
                        case "A":
                            _combined = 210;
                            _building = 132;
                            _contents = 78;
                            break;
                        case "B":
                            _combined = 210;
                            _building = 132;
                            _contents = 78;
                            break;
                        case "C":
                            _combined = 246;
                            _building = 148;
                            _contents = 98;
                            break;
                        case "D":
                            _combined = 276;
                            _building = 168;
                            _contents = 108;
                            break;
                        case "E":
                            _combined = 330;
                            _building = 199;
                            _contents = 131;
                            break;
                        case "F":
                            _combined = 408;
                            _building = 260;
                            _contents = 148;
                            break;
                        case "G":
                            _combined = 540;
                            _building = 334;
                            _contents = 206;
                            break;
                        case "H":
                            _combined = 1200;
                            _building = 800;
                            _contents = 400;
                            break;
                    }
                            break;
                case "Northern Ireland":
                    Console.WriteLine(5);
                    break;
                case "Scotland":
                    switch (band)
                    {
                        case "A":
                            _combined = 210;
                            _building = 132;
                            _contents = 78;
                            break;
                        case "B":
                            _combined = 210;
                            _building = 132;
                            _contents = 78;
                            break;
                        case "C":
                            _combined = 246;
                            _building = 148;
                            _contents = 98;
                            break;
                        case "D":
                            _combined = 276;
                            _building = 168;
                            _contents = 108;
                            break;
                        case "E":
                            _combined = 330;
                            _building = 199;
                            _contents = 131;
                            break;
                        case "F":
                            _combined = 408;
                            _building = 260;
                            _contents = 148;
                            break;
                        case "G":
                            _combined = 540;
                            _building = 334;
                            _contents = 206;
                            break;
                        case "H":
                            _combined = 1200;
                            _building = 800;
                            _contents = 400;
                            break;
                    }
                    break;
                case "Wales":
                    switch (band)
                    {
                        case "A":
                            _combined = 210;
                            _building = 132;
                            _contents = 78;
                            break;
                        case "B":
                            _combined = 210;
                            _building = 132;
                            _contents = 78;
                            break;
                        case "C":
                            _combined = 210;
                            _building = 132;
                            _contents = 78;
                            break;
                        case "D":
                            _combined = 246;
                            _building = 148;
                            _contents = 98;
                            break;
                        case "E":
                            _combined = 276;
                            _building = 168;
                            _contents = 108;
                            break;
                        case "F":
                            _combined = 330;
                            _building = 199;
                            _contents = 131;
                            break;
                        case "G":
                            _combined = 408;
                            _building = 260;
                            _contents = 148;
                            break;
                        case "H":
                            _combined = 540;
                            _building = 334;
                            _contents = 206;
                            break;
                        case "I":
                            _combined = 1200;
                            _building = 800;
                            _contents = 400;
                            break;
                    }
                    break;
            }

            textBox7.Text = _combined.ToString();
            textBox8.Text = _building.ToString();
            textBox9.Text = _contents.ToString();
        }

        

        private string findcountry(string post)
        {

                myConnection.Open();

                SqlCommand cmd = myConnection.CreateCommand();
                cmd.CommandText = "SELECT [Country] FROM [dbo].[postcode] WHERE Postcode ='" + post + "'";

                string result = "";


                using (SqlDataReader reader = cmd.ExecuteReader())
                {
                    while (reader.Read())
                    {
                        //Send these to your WinForms textboxes
                        result = reader["Country"].ToString();
                    }
                }

                myConnection.Close();
                return result;


        }
        private void button1_Click(object sender, EventArgs e)
        { try
            {
                if ((comboBox1.Text == "England") || (comboBox1.Text == "Wales"))
                {
                    textBox4.Text.ToUpper();
                    //Close the existing SQL connection when the search button is clicked, this is because when returning the datatable this doesn't close the connection because it cant until after it's returned the information.
                    myConnection.Close();
                    //Clears the list of policies
                    Houselist.Clear();

                    //Runs the SQL queries and imports into datatable and returns to dt variable
                    DataTable dt = FindID("SELECT [NLPG_UPRN],[UARN],[BAND],[NAME],[NUMB],[SS1],[SS2],[SS3],[ST],[TN],[CO],[POSTCO],[BUILT_PRE2009] FROM [dbo].[flood1] WHERE POSTCO = '" + textBox4.Text + "'", "NUMB asc");
                    myConnection.Close();
                    //If the datatable doesn't contain atleast 1 row then it obviously hasn't found anything so this shows a message box
                    if (dt.Rows.Count < 1)
                    {
                        MessageBox.Show("Cannot find postcode");
                    }
                    //If it has found something then for each row, parse the data into temp variables
                    //And create a new policy object with these variables
                    //Then add this to the policylist
                    //Then bind this data to the listbox

                    else {
                        int i = 0;

                        foreach (DataRow row in dt.Rows)
                        {
                            string band = dt.Rows[i][2].ToString();
                            string name = dt.Rows[i][3].ToString();
                            string number = dt.Rows[i][4].ToString();
                            string ss1 = dt.Rows[i][5].ToString();
                            string street = dt.Rows[i][8].ToString();
                            string town = dt.Rows[i][9].ToString();
                            string postcode = dt.Rows[i][11].ToString();
                            string pre = dt.Rows[i][12].ToString();

                            house house = new house(band, name, number, ss1, street, town, postcode, pre);
                            Houselist.Add(house);

                            i++;
                            BindData();
                        }
                        //This shows the first row of the datatable in the GUI
                        textBox1.Text = dt.Rows[0][2].ToString();
                        textBox2.Text = dt.Rows[0][3].ToString() + " " + dt.Rows[0][4].ToString() + " " + dt.Rows[0][5].ToString();
                        textBox3.Text = dt.Rows[0][9].ToString();
                        textBox5.Text = dt.Rows[0][8].ToString();



                    }
                }
                else if ((comboBox1.Text == "Scotland"))
                {
                    textBox4.Text.ToUpper();
                    //Close the existing SQL connection when the search button is clicked, this is because when returning the datatable this doesn't close the connection because it cant until after it's returned the information.
                    myConnection.Close();
                    //Clears the list of policies
                    Houselist.Clear();

                    
                        string[] ssize = textBox4.Text.Split(null);
                    
                   

                    //Runs the SQL queries and imports into datatable and returns to dt variable
                    DataTable dt = FindID("SELECT [ASSESSOR_ID],[PPRN],[UARN],[UPRN],[UA],[ADDRESS_STATUS],[SAON],[PAON],[STREET],[LOCALITY],[TOWN],[ADMIN_AREA],[POST_TOWN],[PCOUT],[PCIN],[BAND],[DATE_CODE],[GARAGE] FROM [dbo].[wales] WHERE PCOUT = '" + ssize[0] + "' AND PCIN ='" + ssize[1] + "'", "PAON asc");
                    myConnection.Close();
                    //If the datatable doesn't contain atleast 1 row then it obviously hasn't found anything so this shows a message box
                    if (dt.Rows.Count < 1)
                    {
                        MessageBox.Show("Cannot find postcode");
                    }
                    //If it has found something then for each row, parse the data into temp variables
                    //And create a new policy object with these variables
                    //Then add this to the policylist
                    //Then bind this data to the listbox

                    else {
                        int i = 0;

                        foreach (DataRow row in dt.Rows)
                        {
                            string band = dt.Rows[i][15].ToString();
                            string name = dt.Rows[i][6].ToString();
                            string number = dt.Rows[i][7].ToString();
                            string ss1 = dt.Rows[i][8].ToString();
                            string street = dt.Rows[i][9].ToString();
                            string town = dt.Rows[i][10].ToString();
                            string postcode = dt.Rows[i][13].ToString() + " " + dt.Rows[i][15].ToString();
                            string pre = dt.Rows[i][16].ToString();

                            house house = new house(band, name, number, ss1, street, town, postcode, pre);
                            Houselist.Add(house);

                            i++;
                            BindData();
                        }
                        //This shows the first row of the datatable in the GUI
                        textBox1.Text = dt.Rows[0][15].ToString();
                        textBox2.Text = dt.Rows[0][6].ToString() + " " + dt.Rows[0][7].ToString() + " " + dt.Rows[0][8].ToString();
                        textBox3.Text = dt.Rows[0][10].ToString();
                        textBox5.Text = dt.Rows[0][9].ToString();



                    }
                }
                else { MessageBox.Show("Please select a country"); }
            }
            catch (Exception ex)
            {
                MessageBox.Show("Please enter a correct postcode, E.G 'KA3 5HS'");
            }
        }

        private void listBox1_SelectedIndexChanged_1(object sender, EventArgs e)
        {
            if (listBox1.SelectedIndex > -1)
            {
                selecteditem = listBox1.SelectedIndex;
                select();
            }
        }

        private void listBox1_Format(object sender, ListControlConvertEventArgs e)
        {
            string name = ((house)e.ListItem).Name;
            string number = ((house)e.ListItem).Number;
            string ss1 = ((house)e.ListItem).SS1;
           
            e.Value = number + " " + name + " " + ss1;
        }

        private void textBox4_TextChanged(object sender, EventArgs e)
        {

        }

        private void label1_Click(object sender, EventArgs e)
        {

        }
    }
}
