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
using Microsoft.VisualBasic.FileIO;

namespace RenewalUploader
{
    public partial class Form1 : Form
    {
        //Declare variables for renewal uploader (temp)
        string filename = "";
        string path = "";
        string serverpath = "";
        bool filebool = false;
        //Declare variables for inbound uploader (temp)
        string filename2 = "";
        string path2 = "";
        string serverpath2 = "";
        bool filebool2 = false;


        //The SQL connection string to the CAGv17 server
        SqlConnection myConnection = new SqlConnection("Data Source=10.1.3.17,1433;Network Library=DBMSSOCN; Initial Catalog=EDIDB;User ID=EDIDBUser;Password=WlhbR2WeUj;");
        public Form1()
        {
            InitializeComponent();
        }

        private void button1_Click(object sender, EventArgs e)
        {
            //Opens File dialog box and checks for existing file
            dialog();
            //If the file doesn't exist
            if (filebool)
            {
                //Update text box
                textBox1.Text = "Waiting to Complete";
                //Imports CSV file into a locally stored datatable
                DataTable dt = GetDataTabletFromCSVFile(serverpath);
                //Uploads datatable into sql server using SQL Bulk Copy query
                InsertDataIntoSQLServerUsingSQLBulkCopy(dt);
                //Shows how many rows affected
                textBox2.Text = dt.Rows.Count.ToString();
                //Update text box to show upload is successful
                textBox1.Text = "Uploaded Successfully";
                //Show messagebox when the upload is complete
                MessageBox.Show("This has been successfully Uploaded");
            }

        
        }

        private void dialog()

        {
            //Creates open file box which only accepts CSV files
            using (OpenFileDialog openFileDialog1 = new OpenFileDialog())
            {
                openFileDialog1.Filter = "Excel File|*.CSV";
                openFileDialog1.Title = "Open an Excel File";
                //  openFileDialog1.ShowDialog();

                //If OK is clicked then run next bit
                if (openFileDialog1.ShowDialog() == DialogResult.OK)
                {
                    
                    //Stores information in the temp variables we set earlier
                    filename = openFileDialog1.FileName;
                    path = Path.GetFileName(openFileDialog1.FileName);
                    serverpath = Path.Combine("\\\\CAGV17\\", "Weekly\\", System.IO.Path.GetFileName(openFileDialog1.FileName));

                    //If file exists already then show message box, set filebool to false and return to previous function
                    if (File.Exists(serverpath))
                    {
                        MessageBox.Show("File has already been uploaded");
                        filebool = false;
                        return;

                    }
                    //Else copy the file to the server, set filebool to true and clear the file dialog box from the system memory. 
                    else {
                        System.IO.File.Copy(openFileDialog1.FileName, serverpath, true);
                        filebool = true;
                        openFileDialog1.Dispose();

                    }



                }
            }


        }

        //Same as Dialog function but set to different location, could have done this to 1 function rather than 2 and just insert information into the function and store in local variables but there's no need.
        private void Uploaddialog()

        {
            using (OpenFileDialog openFileDialog2 = new OpenFileDialog())
            {
                openFileDialog2.Filter = "Excel File|*.CSV";
                openFileDialog2.Title = "Open an Excel File";
                //  openFileDialog1.ShowDialog();
                if (openFileDialog2.ShowDialog() == DialogResult.OK)
                {

                    filename2 = openFileDialog2.FileName;
                    path2 = Path.GetFileName(openFileDialog2.FileName);
                    serverpath2 = Path.Combine("\\\\CAGV17\\", "Daily\\", System.IO.Path.GetFileName(openFileDialog2.FileName));
                    if (File.Exists(serverpath2))
                    {
                        MessageBox.Show("File has already been uploaded");
                        filebool2 = false;
                        return;

                    }
                    else {
                        System.IO.File.Copy(openFileDialog2.FileName, serverpath2, true);
                        filebool2 = true;
                        openFileDialog2.Dispose();

                    }



                }
            }


        }

        private void Form1_Load(object sender, EventArgs e)
        {

        }
        //Same as Button1 Click but set to different functions, could have cleaned this up into 1 function rather than 2 but no need.
        private void button2_Click(object sender, EventArgs e)
        {

            Uploaddialog();

            if (filebool2)
            {
                textBox4.Text = "Waiting to Complete";
                DataTable dt = GetDataTabletFromCSVFile(serverpath2);
                InsertDataIntoSQLServerUsingSQLBulkCopy2(dt);
                textBox3.Text = dt.Rows.Count.ToString();
                textBox4.Text = "Uploaded Successfully";
                MessageBox.Show("This has been successfully Uploaded");

            }

           

        }
        //Function to import CSV file into a datatable
        private static DataTable GetDataTabletFromCSVFile(string csv_file_path)
        {
            //Create new datatable
            DataTable csvData = new DataTable();
            try
            {
                //Use VB textfieldparser to read each line of the CSV and add them to the datatable
                using (TextFieldParser csvReader = new TextFieldParser(csv_file_path))
                {
                    csvReader.SetDelimiters(new string[] { "," });
                    csvReader.HasFieldsEnclosedInQuotes = true;
                    string[] colFields = csvReader.ReadFields();
                    foreach (string column in colFields)
                    {
                        DataColumn datecolumn = new DataColumn(column);
                        datecolumn.AllowDBNull = true;
                        csvData.Columns.Add(datecolumn);
                    }
                    while (!csvReader.EndOfData)
                    {
                        string[] fieldData = csvReader.ReadFields();
                        //Making empty value as null
                        for (int i = 0; i < fieldData.Length; i++)
                        {
                            if (fieldData[i] == "")
                            {
                                fieldData[i] = null;
                            }
                        }
                        csvData.Rows.Add(fieldData);
                    }
                }
            }
            //If it fails then show messagebox with error.
            catch (Exception ex)
            {
                MessageBox.Show(ex.ToString());
            }
            return csvData;
        }

        //Function to insert datatable into SQL table
        static void InsertDataIntoSQLServerUsingSQLBulkCopy(DataTable csvFileData)
        {
            try {
                using (SqlConnection dbConnection = new SqlConnection("Data Source=10.1.3.17,1433;Network Library=DBMSSOCN; Initial Catalog=EDIDB;User ID=EDIDBUser;Password=WlhbR2WeUj;"))
                {
                    dbConnection.Open();
                    using (SqlBulkCopy s = new SqlBulkCopy(dbConnection))
                    {
                        s.DestinationTableName = "OutboundEDI";
                        foreach (var column in csvFileData.Columns)
                            s.ColumnMappings.Add(column.ToString(), column.ToString());
                        s.WriteToServer(csvFileData);
                    }
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.ToString());
            }
        }
        //Function to insert datatable into SQL table  (Could have created 1 function to do both of these but no need) - Would just insert a string with destination into the function
        static void InsertDataIntoSQLServerUsingSQLBulkCopy2(DataTable csvFileData)
        {
            try {
                using (SqlConnection dbConnection = new SqlConnection("Data Source=10.1.3.17,1433;Network Library=DBMSSOCN; Initial Catalog=TEST_EDIDB;User ID=EDIDBUser;Password=WlhbR2WeUj;"))
                {
                    dbConnection.Open();
                    using (SqlBulkCopy s = new SqlBulkCopy(dbConnection))
                    {
                        s.DestinationTableName = "InboundEDI";
                        foreach (var column in csvFileData.Columns)
                            s.ColumnMappings.Add(column.ToString(), column.ToString());
                        s.WriteToServer(csvFileData);
                    }
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.ToString());
            }
        }

        private void button3_Click(object sender, EventArgs e)
        {

        }

        private void tabPage1_Click(object sender, EventArgs e)
        {

        }

        private void textBox3_TextChanged(object sender, EventArgs e)
        {

        }

        private void tabPage2_Click(object sender, EventArgs e)
        {

        }
    }
}
