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

namespace InboundErrorChecking
{
    public partial class Form1 : Form
    {
        //Create a list of policy objects from the policy class
        List<Policy> Policycontent = new List<Policy>();
        //Sql connection string to server
        SqlConnection myConnection = new SqlConnection("Data Source=10.1.3.17,1433;Network Library=DBMSSOCN; Initial Catalog=EDIDB;User ID=EDIDBUser;Password=WlhbR2WeUj;");
        //Selected item is defaulted at -1
        int selecteditem = -1;

        public Form1()
        {
            InitializeComponent();
            //On start of application run the BindData Function
            BindData();
        }

        private void groupBox1_Enter(object sender, EventArgs e)
        {

        }

        private void Form1_Load(object sender, EventArgs e)
        {

        }

        private DataTable FindID()
        {
            //Create a new datatable and open the SQL connection
            DataTable dt = new DataTable();
            myConnection.Open();

            //Run SQL query to find information based on whats in textbox1 (Policy Number)
            SqlCommand cmd = myConnection.CreateCommand();
            cmd.CommandText = "SELECT [Policy Reference],Agency,[Quote House],[Message Type],[Event Type],RootATSID,[Error Summary] FROM InboundEDI WHERE[Policy Reference]= '" + textBox1.Text +"'";

            /*      using (SqlDataReader reader = cmd.ExecuteReader(CommandBehavior.CloseConnection))
                  {
                      while (reader.Read())
                      {
                          dt.Load(reader);

                              return dt;
                      }
                  }
                  */
                  //Fill datatable with this information
            using (var da = new SqlDataAdapter(cmd.CommandText, myConnection))
            {
                da.Fill(dt);
            }
            //      myConnection.Close();
            //Return the datatable
            return dt;
        }

        private void label1_Click(object sender, EventArgs e)
        {

        }

        private void textBox1_TextChanged(object sender, EventArgs e)
        {

        }

        private void button1_Click(object sender, EventArgs e)
        {
            //Close the existing SQL connection when the search button is clicked, this is because when returning the datatable this doesn't close the connection because it cant until after it's returned the information.
            myConnection.Close();
            //Clears the list of policies
            Policycontent.Clear();

            //Runs the SQL queries and imports into datatable and returns to dt variable
            DataTable dt = FindID();
            //If the datatable doesn't contain atleast 1 row then it obviously hasn't found anything so this shows a message box
            if (dt.Rows.Count < 1)
            {
                MessageBox.Show("No errors for this policy");
            }
            //If it has found something then for each row, parse the data into temp variables
            //And create a new policy object with these variables
            //Then add this to the policylist
            //Then bind this data to the listbox
            else {
                int i = 0;
                foreach( DataRow row in dt.Rows)
                {
                    string pol = dt.Rows[i][0].ToString();
                    string eventt = dt.Rows[i][4].ToString();
                    string roott = dt.Rows[i][5].ToString();
                    string error = dt.Rows[i][6].ToString();

                    Policy policy = new Policy(pol, eventt, roott, error);
                    Policycontent.Add(policy);
                    
                     i++;
                    BindData();
                }
                //This shows the first row of the datatable in the GUI
                textBox3.Text = dt.Rows[0][5].ToString();
                textBox4.Text = dt.Rows[0][4].ToString();
                textBox2.Text = dt.Rows[0][6].ToString();
            }

            }

        private void addtotable()
        {

        }
        void BindData()
        {
            //Binds the policycontent list to the listbox on the GUI and shows the Eventtype name (e.g ADJ, CAB, Etc.)
            listBox1.DataSource = null;
            listBox1.DataSource = Policycontent;
            listBox1.DisplayMember = "Eventtype";


        }

        private void label3_Click(object sender, EventArgs e)
        {

        }

        private void label2_Click(object sender, EventArgs e)
        {

        }

        private void listBox1_SelectedIndexChanged(object sender, EventArgs e)
        {
            //If the selected listbox value is greater that -1 then set the selected index to this value
            if (listBox1.SelectedIndex > -1)
            {
                selecteditem = listBox1.SelectedIndex;
                select();
            }

          
        }

        private void select()
        {
            //Update the textboxes on the GUI with the policy information that you click on in the list box
            textBox3.Text = Policycontent[listBox1.SelectedIndex].Root;
            textBox4.Text = Policycontent[listBox1.SelectedIndex].Eventtype;
            textBox2.Text = Policycontent[listBox1.SelectedIndex].Error;
        }
    }
}
