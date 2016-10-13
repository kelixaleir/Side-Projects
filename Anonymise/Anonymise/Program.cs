using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Text.RegularExpressions;
using System.Threading.Tasks;

namespace Anonymise
{
    class Program
    {
        static void Main(string[] args)
        {
            //Gets file paths of items in "Anonymise" folder
            string[] filePaths = Directory.GetFiles(@"\\cagv22\EDI\Archive Location\Inbound\Anonymise");
            try { 
            //For each filepath
            foreach (string path in filePaths)
            {

                //Temp variable that holds each line of file
                var lines = File.ReadAllLines(path);
                //Temp postcode used to replace
                string temppostcode = "BD1 1AT";

                //
                for (int i = 0; i < lines.Length; i++)
                {
                    //If line cotains "PID"
                    if (lines[i].Contains("PID"))
                    {

                        string temp = lines[i];
                        string[] templines = Regex.Split(lines[i], @"[+]");
                        templines[2] = "DOE:JOHN";
                        templines[5] = "01011991";
                        lines[i] = "";
                        foreach (string line in templines)
                        {

                            if (line == templines[templines.Length - 1])
                            { lines[i] += (line + "'"); }
                            else { lines[i] += line + "+"; }
                        }


                    }

                    if (lines[i].Contains("PLO="))
                    {
                        string temp = lines[i];
                        string[] templines = Regex.Split(lines[i], @"[+]");
                        templines[0] = "PLO=1 ANONYMOUS DRIVE:ANON:ANON";
                        temppostcode = templines[1];
                        templines[1] = "BD1 5BA";

                        if (templines.Length == 3)
                        {
                            templines[2] = "01274700888";
                        }
                        if (templines.Length == 4)
                        {
                            templines[2] = "01274700888";
                            templines[3] = "01274700800";
                        }
                        lines[i] = "";
                        foreach (string line in templines)
                        {

                            if (line == templines[templines.Length - 1])
                            {
                                // char temp = "'";
                                if (line[line.Length - 1] == '\'')
                                {
                                    lines[i] += (line);
                                }
                                else
                                {
                                    lines[i] += (line + "'");
                                }

                            }
                            else { lines[i] += line + "+"; }
                        }


                    }
                    if (lines[i].Contains("OCD="))
                    {
                        string temp = lines[i];
                        string[] templines = Regex.Split(lines[i], @"[+]");


                        templines[0] = "OCD=DOE:JOHN";
                        templines[3] = "01011991";
                        lines[i] = "";
                        foreach (string line in templines)
                        {

                            if (line == templines[templines.Length - 1])
                            {
                                // char temp = "'";
                                if (line[line.Length - 1] == '\'')
                                {
                                    lines[i] += (line);
                                }
                                else
                                {
                                    lines[i] += (line + "'");
                                }

                            }
                            else { lines[i] += line + "+"; }
                        }


                    }
                    if (lines[i].Contains("RCA="))
                    {
                        lines[i] = lines[i].Replace(temppostcode, "BD1 5BA");
                    }

                    if (lines[i].Contains("RDD="))
                    {
                        lines[i] = lines[i].Replace(temppostcode, "BD1 5BA");
                    }



                }

                File.WriteAllLines(path, lines);
            }

        }
        catch(Exception ex)
            {
                System.IO.File.WriteAllText(@"C:\TestFolder\WriteText.txt", ex.ToString());
            }
        }

    }
}
