using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Net;
using System.IO;
using System.Net.FtpClient;
using System.Collections.Generic;
using System.Reflection;
using System.Diagnostics;

namespace EDI_Automation
{

    class Program
    {
        static List<string> saveedi = new List<string>();
        static void Main(string[] args)
        {
            try
            {
                SetMethodRequiresCWD();
                using (new System.Net.FtpClient.FtpClient())
                {
                    string str;
                    FtpWebRequest request1 = (FtpWebRequest)WebRequest.Create("ftp://ftp4.kewill.net/live/archive/edi/in/");
                    request1.Method = "NLST";
                    request1.Credentials = new NetworkCredential("congregational", "h5jVz2NH");
                    Stream responseStream = ((FtpWebResponse)request1.GetResponse()).GetResponseStream();
                    StreamReader reader = new StreamReader(responseStream);
                    List<string> list = new List<string>();
                    new List<string>();
                    while ((str = reader.ReadLine()) != null)
                    {
                        string[] textArray1 = str.Split(null);
                        string item = textArray1[textArray1.Length - 1];
                        list.Add(item);
                    }
                    responseStream.Close();
                    reader.Close();
                    for (int i = 0; i < (list.Count - 1); i++)
                    {
                        string url = "ftp://ftp4.kewill.net/live/archive/edi/in/" + list[i];
                        string path = @"\\cagv22\EDI\Archive Location\Inbound\test\temp\" + list[i];
                        if (File.Exists(path))
                        {
                            File.Delete(path);
                        }
                        if (CheckIfFileExistsOnServer(list[i], "ftp://ftp4.kewill.net/live/archive/edi/in/"))
                        {
                            Checkdatemodified(url, path);
                        }
                    }
                }
            }
            catch (Exception exception)
            {
                File.WriteAllText(@"C:\TestFolder\WriteText.txt", exception.ToString());
            }
            try
            {
                SetMethodRequiresCWD();
                using (new System.Net.FtpClient.FtpClient())
                {
                    string str5;
                    FtpWebRequest request2 = (FtpWebRequest)WebRequest.Create("ftp://ftp4.kewill.net/live/archive/csv/in/");
                    request2.Method = "NLST";
                    request2.Credentials = new NetworkCredential("congregational", "h5jVz2NH");
                    Stream stream = ((FtpWebResponse)request2.GetResponse()).GetResponseStream();
                    StreamReader reader2 = new StreamReader(stream);
                    List<string> list2 = new List<string>();
                    new List<string>();
                    while ((str5 = reader2.ReadLine()) != null)
                    {
                        string[] textArray2 = str5.Split(null);
                        string str6 = textArray2[textArray2.Length - 1];
                        list2.Add(str6);
                    }
                    stream.Close();
                    reader2.Close();
                    for (int j = 0; j < (list2.Count - 1); j++)
                    {
                        string str7 = "ftp://ftp4.kewill.net/live/archive/CSV/in/" + list2[j];
                        string str8 = @"\\cagv22\EDI\Archive Location\Inbound\test\temp\" + list2[j];
                        if (File.Exists(str8))
                        {
                            File.Delete(str8);
                        }
                        if (CheckIfFileExistsOnServer(list2[j], "ftp://ftp4.kewill.net/live/archive/CSV/in/"))
                        {
                            Checkdatemodified(str7, str8);
                        }
                    }
                }
            }
            catch (Exception exception2)
            {
                File.WriteAllText(@"C:\TestFolder\WriteText.txt", exception2.ToString());
            }
            SortEdi();
        }





        static void Checkdatemodified(string url, string savePath)
        {

            FtpWebRequest request = (FtpWebRequest)WebRequest.Create(url);
            request.Method = WebRequestMethods.Ftp.GetDateTimestamp;

            request.Credentials = new NetworkCredential("congregational", "h5jVz2NH");
            request.UseBinary = true;
            DateTime temp;
            using (FtpWebResponse response = (FtpWebResponse)request.GetResponse())
            {
                temp = response.LastModified;
            }

            FtpWebRequest request2 = (FtpWebRequest)WebRequest.Create(url);
            request2.Credentials = new NetworkCredential("congregational", "h5jVz2NH");
            request2.Method = WebRequestMethods.Ftp.DownloadFile;
            request2.UseBinary = true;

            if (temp > DateTime.Now.AddDays(-2))
            {
                saveedi.Add(url);
                using (FtpWebResponse response2 = (FtpWebResponse)request2.GetResponse())
                {
                    //   DateTime temp = response.LastModified;
                    using (Stream rs = response2.GetResponseStream())
                    {
                        using (FileStream ws = new FileStream(savePath, FileMode.Create))
                        {
                            byte[] buffer = new byte[2048];
                            int bytesRead = rs.Read(buffer, 0, buffer.Length);

                            while (bytesRead > 0)
                            {
                                ws.Write(buffer, 0, bytesRead);
                                bytesRead = rs.Read(buffer, 0, buffer.Length);

                            }
                        }
                        //      DateTime temp = Checkdatemodified(url);
                        System.IO.File.SetLastWriteTime(savePath, temp);
                    }

                }
            }
        }




        static void SortEdi()
        {
            string[] filePaths = Directory.GetFiles(@"\\cagv22\EDI\Archive Location\Inbound\temp");

            foreach (string path in filePaths)
            {
                FileInfo info = new FileInfo(path);
                string filename = info.Name;
                string year = info.LastWriteTime.ToString("yyy");
                string tmep = info.LastWriteTime.ToString("yyyMMdd");
                string pathlocation = @"\\cagv22\EDI\Archive Location\Inbound\" + year + "\\" + tmep + "\\" + filename;

                System.IO.Directory.CreateDirectory(@"\\cagv22\EDI\Archive Location\Inbound\" + year + "\\" + tmep);
                if (System.IO.File.Exists(pathlocation))

                {
                    File.Delete(pathlocation);
                }

                File.Move(path, pathlocation);
            }
        }

        static private bool CheckIfFileExistsOnServer(string fileName, string location)
        {
            var request = (FtpWebRequest)WebRequest.Create(location + fileName);
            request.Credentials = new NetworkCredential("congregational", "h5jVz2NH");
            request.Method = WebRequestMethods.Ftp.GetFileSize;

            try
            {
                FtpWebResponse response = (FtpWebResponse)request.GetResponse();
                return true;
            }
            catch (WebException ex)
            {
                FtpWebResponse response = (FtpWebResponse)ex.Response;
                if (response.StatusCode == FtpStatusCode.ActionNotTakenFileUnavailable)
                    return false;
            }
            return false;
        }

        private static void SetMethodRequiresCWD()
        {
            Type requestType = typeof(FtpWebRequest);
            FieldInfo methodInfoField = requestType.GetField("m_MethodInfo", BindingFlags.NonPublic | BindingFlags.Instance);
            Type methodInfoType = methodInfoField.FieldType;


            FieldInfo knownMethodsField = methodInfoType.GetField("KnownMethodInfo", BindingFlags.Static | BindingFlags.NonPublic);
            Array knownMethodsArray = (Array)knownMethodsField.GetValue(null);

            FieldInfo flagsField = methodInfoType.GetField("Flags", BindingFlags.NonPublic | BindingFlags.Instance);

            int MustChangeWorkingDirectoryToPath = 0x100;
            foreach (object knownMethod in knownMethodsArray)
            {
                int flags = (int)flagsField.GetValue(knownMethod);
                flags |= MustChangeWorkingDirectoryToPath;
                flagsField.SetValue(knownMethod, flags);
            }
        }

    }
}
