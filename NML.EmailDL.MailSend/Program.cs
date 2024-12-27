using Microsoft.SharePoint.Client;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Net;
using System.Text;
using System.Threading.Tasks;

namespace NML.EmailDL.MailSend
{
    class Program
    {

        static void Main(string[] args)
        {
            // Console.WriteLine("1");
            ServicePointManager.SecurityProtocol = SecurityProtocolType.Tls12;
            //string realm = TokenHelper.GetRealmFromTargetUrl(Constants.siteURL);
            //var token = TokenHelper.GetAppOnlyAccessToken(
            //                    TokenHelper.SharePointPrincipal,
            //                    Constants.siteURL.Authority, realm).AccessToken;
            //Helper.GetItemfromMailMissendList("367", "MailMisSendingPreventionRequest", token);
            // ServicePointManager.SecurityProtocol = SecurityProtocolType.Tls12;
            SendMailList();
            // Console.WriteLine("finsih");
            // Console.ReadLine();
        }

        private static void SendMailList()
        {
            try
            {
                if (!string.IsNullOrEmpty(Constants.LogFilePath) && !string.IsNullOrEmpty(Constants.MailContentListDisplayName) && !string.IsNullOrEmpty(Constants.siteURL.ToString()))
                {
                    // Console.WriteLine("2");
                    //  Console.WriteLine("Batch Execution Started");
                    Helper.WriteLog(",,,,,,," + "BatchExecutionStarted,");
                    string realm = TokenHelper.GetRealmFromTargetUrl(Constants.siteURL);
                    var token = TokenHelper.GetAppOnlyAccessToken(
                                        TokenHelper.SharePointPrincipal,
                                        Constants.siteURL.Authority, realm, "").AccessToken;

                    using (var ctx = TokenHelper.GetClientContextWithAccessToken(Constants.siteURL.ToString(), token))
                    {
                        //string sExceptionMessage = string.Empty;
                        //Helper.TriggerEmail("Subash-venkatachalam@mail.nissan.co.jp", "", "", "", "", "SendMailToSupplier",out sExceptionMessage, "https://nissangroup.sharepoint.com/teams/WS_JAO_NML_007_TechMB2EAMO/Shared%20Documents/Forms/AllItems.aspx", ctx);
                        if (Helper.GetAllListItems(ctx, token))
                        {
                            Helper.WriteLog(",,,,,,," + "BatchExecutionCompleted,");
                        }
                        else
                        {
                            Helper.WriteLog(",,,,,,," + "ErrorInbatchExecution,");
                        }
                    }
                }
                else
                {
                    Console.WriteLine("In Put Parameters are InCorrect");
                    Helper.WriteLog(",,,,,,," + "ErrorInbatchExecution,");
                }
                Console.WriteLine("Batch Execution Completed");
            }
            catch (Exception ee)
            {
                Console.WriteLine(ee.Message);
                Helper.WriteLog(ee.Message);
            }
        }
    }
}
