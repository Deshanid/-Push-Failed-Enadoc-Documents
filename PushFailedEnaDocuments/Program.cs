using System;
using System.Collections.Generic;
using System.Linq;
using System.Security;
using System.Text;
using System.Threading.Tasks;
using Microsoft.SharePoint.Client;
using Microsoft.SharePoint.Client.Utilities;
using System.Net;
using RestSharp;
using System.Data;
using System.Configuration;

namespace PushFailedEnaDocuments
{
    class Program
    {
        static void Main(string[] args)
        {
            GetFailedDocuments();
        }
        static void ctx_MixedAuthRequest(object sender, WebRequestEventArgs e)
        {
           try
            {

                //Add the header that tells SharePoint to use Windows authentication.

                e.WebRequestExecutor.RequestHeaders.Add(

                "X-FORMS_BASED_AUTH_ACCEPTED", "f");

            }

            catch (Exception ex)
            {

               // this.Print("ctx_MixedAuthRequest", "Error setting authentication header: " + ex.Message);

                //MessageBox.Show("Error setting authentication header: " + ex.Message);

            }

        }
        public static void GetFailedDocuments()
        {           
            using (ClientContext ctx = new ClientContext(ConfigurationManager.AppSettings["WebUrl"].ToString())) //sharepointurl
            {
                // SecureString securePassword = GetSecureString(Utility.)
                ctx.ExecutingWebRequest += new EventHandler<WebRequestEventArgs>(ctx_MixedAuthRequest);

                //Set the Windows credentials.

                ctx.AuthenticationMode = ClientAuthenticationMode.Default;

                try
                {
                    //Connect to sharepoint
                    string username = ConfigurationManager.AppSettings["UserName"].ToString();
                    string password = ConfigurationManager.AppSettings["Password"].ToString();
                    string domain = ConfigurationManager.AppSettings["Domain"].ToString();
                    ctx.Credentials = new NetworkCredential(username,password,domain); 
                    Console.WriteLine("Successfully connected to Sharepoint");

                    //Build Query.
                    Web web = ctx.Web;
                    List list = web.Lists.GetByTitle(ConfigurationManager.AppSettings["Title"].ToString());
                    var query = new CamlQuery();
                    query.ViewXml = string.Format(@"<View>
                                                    <Query>
                                                    <Where>
                                                    <Eq>
                                                    <FieldRef Name='IsEnadocUploaded' />
                                                    <Value Type='Boolean'>0</Value>
                                                    </Eq>
                                                    </Where>                         
                                                    </Query>
                                                    </View>");
                  

                    ListItemCollection listItems = list.GetItems(query);

                    //Load List
                    ctx.Load(listItems);
                    ctx.ExecuteQuery();
                    var s = listItems.ToList().Select(i => i["MainDocumentId"]).Distinct();

                    //Loop No of IDs
                    foreach (var itm in s)
                    {
                        if (itm != null)
                        {
                            Console.WriteLine(itm);
                            int ID = int.Parse(itm.ToString());


                                try
                                {
                                    var client = new RestClient(ConfigurationManager.AppSettings["ApiUrl"].ToString());//api
                                    var request = new RestRequest(ConfigurationManager.AppSettings["ApiResource"].ToString(), Method.POST);
                                    string mainDocumentID = ID.ToString();
                                    request.AddHeader("Id", mainDocumentID);
                                    var result = client.Execute(request);
                                    if (result.StatusCode == HttpStatusCode.OK)
                                    {
                                        bool text = true;
                                        Print("GetFailedDocuments ", " Document Uploaded. Document ID : " + mainDocumentID, text);
                                        Console.WriteLine("Document Uploaded. Document ID: " + mainDocumentID);
                                        continue;
                                    }
                                    else
                                    {
                                        bool text = true ;
                                        Print("GetFailedDocuments ", " Document Failed from StatusCode. Document ID : " + mainDocumentID, text);
                                        Console.WriteLine("Document Failed.Document ID: " + mainDocumentID);
                                    }
                                }
                                catch(Exception e)
                                {

                                bool text = false;
                                string er = string.Format(" Document Failed. Documet Id : {0} Error : {1}", ID, e);
                                Print("GetFailedDocuments", er , text);
                                Console.WriteLine("Error : " +  e);
                                continue;
                                }
                            
                        }
                        else
                        {
                            bool text = true;
                            Print("GetFailedDocuments", " Document Failed. Document ID : Null " , text);
                            Console.WriteLine("Document Id is Null.");
                        }
                        Console.ReadLine();
                    }

                    bool text1 = true;
                    string logmsg = String.Format("Successfully pushed {0} Documents. ", listItems.Count());
                    Print("GetFailedDocuments", logmsg, text1);
                    Console.WriteLine("Successfully pushed {0} Documents. " , listItems.Count());

                    //key - auto - value - no
                    var run = ConfigurationManager.AppSettings["auto"];

                    Console.ReadLine();
                }
                    
               
                catch (Exception ex)
                {
                    bool text = false;
                    Print("GetFailedDocuments", "Error : " + ex, text);
                    Console.WriteLine(ex);
                    
                }
            }
            return;
        }

        public static void Print(string method, string msg, bool text)
        {

            var _class = typeof(Program);
            var _namespace = _class.Namespace;




            if (text == true)
            {

                Logger.PrintExecutionLog(_namespace, "Program", method, msg);
                return;

            }

            else
            {
                Logger.PrintError(_namespace, "Program", method, msg);
                return;
            }
        }
    }
}
