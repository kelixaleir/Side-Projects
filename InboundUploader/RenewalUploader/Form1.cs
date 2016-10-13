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

namespace RenewalUploader
{
    public partial class Form1 : Form
    {
        string filename = "";
        string path = "";
        string serverpath = "";
 
        bool filebool = false;
        SqlConnection myConnection = new SqlConnection("Data Source=10.1.3.17,1433;Network Library=DBMSSOCN; Initial Catalog=EDIDB;User ID=EDIDBUser;Password=WlhbR2WeUj;");
        public Form1()
        {
            InitializeComponent();
        }

        private void button1_Click(object sender, EventArgs e)
        {

            dialog();
            
            if (filebool)
            {
                
                UploadCSV();
            }


        }

        private void dialog()

        {
            using (OpenFileDialog openFileDialog1 = new OpenFileDialog())
            {
                openFileDialog1.Filter = "Excel File|*.CSV";
                openFileDialog1.Title = "Open an Excel File";
                //  openFileDialog1.ShowDialog();
                if (openFileDialog1.ShowDialog() == DialogResult.OK)
                {
                    
                    filename = openFileDialog1.FileName;
                    path = Path.GetFileName(openFileDialog1.FileName);
                    serverpath = Path.Combine("\\\\CAGV17\\", "Weekly\\", System.IO.Path.GetFileName(openFileDialog1.FileName));
                    if (File.Exists(serverpath))
                    {
                        MessageBox.Show("File has already been uploaded");
                        filebool = false;
                        return;

                    }
                    else {
                        System.IO.File.Copy(openFileDialog1.FileName, serverpath, true);
                        filebool = true;
                        openFileDialog1.Dispose();

                    }



                }
            }


        }

        private void UploadCSV()
        {
            

            // FileStream stream = File.Open(openFileDialog1.FileName, FileMode.Open, FileAccess.Read);

            SqlCommand cmd = myConnection.CreateCommand();
            // The SQL command
            cmd.CommandText = "BULK INSERT OutboundEDI FROM '" + "\\\\CAGV17\\" + "Daily\\" + path + "'  WITH (FIELDTERMINATOR = ',',ROWTERMINATOR = '\n')";
            myConnection.Open();
            int result = cmd.ExecuteNonQuery();
            textBox2.Text = result.ToString();
            textBox1.Text = "Uploaded";


            // textBox2.Text = (string)cmd.ExecuteScalar();

            myConnection.Close();
        }

        private void Form1_Load(object sender, EventArgs e)
        {

        }
    }
}
