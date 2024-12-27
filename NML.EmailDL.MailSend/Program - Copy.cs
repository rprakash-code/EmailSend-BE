using Microsoft.SharePoint.Client;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace NML.EmailDL.MailSend
{
    class Program
    {
        static void Main(string[] args)
        {

        }

        private static void SendMailList()
        {
            try
            {


                // string sExceptionMessage = string.Empty;
                //   Helper.TriggerEmail("pr00464627@techmahindra.com; newprakash22389@gmail.com", "", "", "Test", "Test 123", "", out  sExceptionMessage);

                if (!string.IsNullOrEmpty(Constants.LogFilePath) && !string.IsNullOrEmpty(Constants.MailContentListDisplayName) && !string.IsNullOrEmpty(Constants.siteURL.ToString()))
                {
                    Console.WriteLine("Batch Execution Started");
                    Helper.WriteLog(",,,,,," + "BatchExecutionStarted,");
                    string realm = TokenHelper.GetRealmFromTargetUrl(Constants.siteURL);
                    var token = TokenHelper.GetAppOnlyAccessToken(
                                        TokenHelper.SharePointPrincipal,
                                        Constants.siteURL.Authority, realm).AccessToken;


                    if (Helper.GetAllListItems(token))
                    {
                        Helper.WriteLog(",,,,,," + "BatchExecutionCompleted,");
                    }
                    else
                    {
                        Helper.WriteLog(",,,,,," + "ErrorInbatchExecution,");
                    }

                }
                else
                {
                    Console.WriteLine("In Put Parameters are InCorrect");
                    Helper.WriteLog(",,,,,," + "ErrorInbatchExecution,");
                }
                Console.ReadLine();
            }
            catch (Exception ee)
            {
                Console.WriteLine(ee.Message);
                Helper.WriteLog(ee.Message);
            }
        }
    }
}
