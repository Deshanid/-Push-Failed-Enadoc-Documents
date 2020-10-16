using System;
using System.Collections.Generic;
using System.Configuration;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace PushFailedEnaDocuments
{
    class Logger
    {

        public static bool PrintError(string _namespace, string _class, string method, string msg)

        {

            try

            {

                string fullPath = "";





                // fullPath = ConfigurationManager.AppSettings["ErrorLogPath"].ToString();





                fullPath = System.AppDomain.CurrentDomain.BaseDirectory.ToString();









                string date = DateTime.Now.Year + "-" + DateTime.Now.Month + "-" + DateTime.Now.Day;



                fullPath = fullPath + "ErrorLogs";

                if (!Directory.Exists(fullPath))

                {

                    Directory.CreateDirectory(fullPath);

                }



                string filePath = fullPath + "\\" + date + "_ErrorLog.txt";

                if (!File.Exists(filePath))

                {

                    TextWriter sw = new StreamWriter(filePath);



                    sw.WriteLine("Layer Name :-" + _namespace);

                    sw.WriteLine("Class Name :-" + _class);

                    sw.WriteLine("Method Name :-" + method);

                    sw.WriteLine("Date Time :-" + DateTime.Now);

                    sw.WriteLine("Error Message :-" + msg);



                    sw.Close();

                }

                else

                {

                    string oldLine = File.ReadAllText(filePath);

                    TextWriter tw = new StreamWriter(filePath);



                    tw.WriteLine(oldLine);



                    tw.WriteLine(tw.NewLine);



                    tw.WriteLine("Layer Name :-" + _namespace);

                    tw.WriteLine("Class Name :-" + _class);

                    tw.WriteLine("Method Name :-" + method);

                    tw.WriteLine("Date Time :-" + DateTime.Now);

                    tw.WriteLine("Error Message :-" + msg);

                    tw.Close();

                }

            }

            catch (Exception)

            {

                return false;

            }

            return true;

        }



        public static bool PrintExecutionLog(string _namespace, string _class, string method, string msg)

        {

            try

            {

                string fullPath = "";

                fullPath = System.AppDomain.CurrentDomain.BaseDirectory.ToString();

                //string fullPath = WebConfigurationManager.AppSettings["ExecutionLogPath"].ToString();
                //System.AppDomain.CurrentDomain.BaseDirectory.ToString();

                string date = DateTime.Now.Year + "-" + DateTime.Now.Month + "-" + DateTime.Now.Day + "-" + DateTime.Now.Hour;

                fullPath = fullPath + "ExecutionLog";

                if (!Directory.Exists(fullPath))

                {

                    Directory.CreateDirectory(fullPath);

                }



                string filePath = fullPath + "\\" + date + "_Log.txt";

                if (!File.Exists(filePath))

                {

                    TextWriter sw = new StreamWriter(filePath);



                    sw.WriteLine("Layer Name :- " + _namespace);

                    sw.WriteLine("Class Name :- " + _class);

                    sw.WriteLine("Method Name :- " + method);

                    sw.WriteLine("Date Time :- " + DateTime.Now);

                    sw.WriteLine("Message :- " + msg);



                    sw.Close();

                }

                else

                {

                    string oldLine = File.ReadAllText(filePath);

                    TextWriter tw = new StreamWriter(filePath);

                    tw.WriteLine(oldLine);

                    tw.WriteLine(tw.NewLine);


                    tw.WriteLine("Layer Name :- " + _namespace);

                    tw.WriteLine("Class Name :- " + _class);

                    tw.WriteLine("Method Name :- " + method);

                    tw.WriteLine("Date Time :- " + DateTime.Now);

                    tw.WriteLine("Message :- " + msg);

                    tw.Close();

                }

            }

            catch (Exception)

            {

                return false;

            }

            return true;

        }
    }
}
