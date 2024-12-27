using Microsoft.IdentityModel.Clients.ActiveDirectory;
using System;
using System.IO;
using System.Reflection;
using System.Collections.Generic;
using System.Configuration;
using System.Linq;
using System.Xml;
using System.Text;
using System.Threading.Tasks;

using System.Net;
using System.Net.Mail;
using Newtonsoft.Json.Linq;
using Microsoft.SharePoint.Client;
using System.Text.RegularExpressions;
using System.Security;
using Group = Microsoft.SharePoint.Client.Group;
using System.Collections;
using Newtonsoft.Json;
using System.Web.Script.Serialization;
using System.Web;

namespace NML.EmailDL.MailSend
{
    class Helper
    {
        public static string SPOPermissionExportString = "";
        public static int incr = 1;
        public static Dictionary<string, string> SharedUserList = new Dictionary<string, string>();
        public static Dictionary<string, string> DirectUsersList = new Dictionary<string, string>();
        public static string GetAccessToken()
        {
            string accessToken = string.Empty;
            try
            {
                string AzureWebApplication_ClientID = ConfigurationManager.AppSettings["AzureWebApplication_ClientID"].ToString();
                string AzureWebApplication_Key = ConfigurationManager.AppSettings["AzureWebApplication_Key"].ToString();
                string tenantName = ConfigurationManager.AppSettings["TenantName"].ToString();

                string authString = "https://login.microsoftonline.com/" + tenantName;
                AuthenticationContext authenticationContext = new AuthenticationContext(authString, false);
                // Config for OAuth client credentials  
                string clientId = ConfigurationManager.AppSettings["AzureWebApplication_ClientID"].ToString();
                ClientCredential clientCred = new ClientCredential(clientId, AzureWebApplication_Key);
                string resource = "https://graph.microsoft.com";
                accessToken = authenticationContext.AcquireTokenAsync(resource, clientCred).Result.AccessToken;
            }
            catch (Exception ex)
            {
                //Logs("Error at GetAccessToken : " + ex.Message);
            }
            return accessToken;
        }

        public static void WriteLog_EmailContent(string LogString)
        {
            try
            {
                StreamWriter log;
                string strFileName = Constants.LogFilePath + "Log_EmailContent_" + DateTime.Now.ToString("yyyyMMdd") + ".log";
                if (!System.IO.File.Exists(strFileName))
                {
                    log = new StreamWriter(strFileName, false, System.Text.Encoding.GetEncoding("UTF-8"));
                }
                else
                {
                    log = System.IO.File.AppendText(strFileName);
                }
                log.WriteLine(DateTime.Now.ToString("[yyyy/MM/dd : HH:mm:ss]") + "," + LogString);
                log.Close();
            }
            catch (Exception ex)
            {
                Console.WriteLine("Error in Log Update for EmailContentLog: " + ex.Message + "......." + ex.InnerException);
            }
        }
        public static void WriteLog(string LogString)
        {
            try
            {
                StreamWriter log;
                string strFileName = Constants.LogFilePath + "Log_DLMailSend_" + DateTime.Now.ToString("yyyyMMdd") + ".CSV";
                if (!System.IO.File.Exists(strFileName))
                {
                    log = new StreamWriter(strFileName, false, System.Text.Encoding.GetEncoding("UTF-8"));
                    log.WriteLine("DateTime,ListName,ItemID,ToEmails,CCEmails,BccEmails,RequestType,ActionType,Status,Comments");
                }
                else
                {
                    log = System.IO.File.AppendText(strFileName);
                }
                log.WriteLine(DateTime.Now.ToString("[yyyy/MM/dd : HH:mm:ss]") + "," + LogString);
                log.Close();
            }
            catch (Exception ex)
            {
                Console.WriteLine("Error in Log Update: " + ex.Message + "......." + ex.InnerException);
            }
        }

        public static long GetMaxIDValue(string siteURL, string listName, string filterCondition, string accessToken)
        {
            long itemCount = 0;
            string json = null;
            try
            {
                string APIURL = String.Concat(siteURL.TrimEnd('/'), "/_api/web/lists/getbytitle('", listName, "')/items?");

                if (string.IsNullOrEmpty(filterCondition))
                {
                    APIURL += "$select=ID&$orderby=ID%20desc&$top=1";
                }

                else
                {
                    APIURL += "$filter=(" + filterCondition + ")&$select =ID&$orderby=ID%20desc&$top=1";
                }

                HttpWebRequest GetMaxIDRequest = HttpWebRequest.Create(APIURL) as HttpWebRequest;
                GetMaxIDRequest.Method = "GET";//"POST";// "DELETE"; //
                GetMaxIDRequest.ContentType = "application/json; odata=verbose";
                GetMaxIDRequest.Accept = "application/json; odata=verbose";
                GetMaxIDRequest.ContentLength = 0;
                GetMaxIDRequest.Headers.Add("Authorization", string.Concat("Bearer ", accessToken));
                using (HttpWebResponse response = GetMaxIDRequest.GetResponse() as HttpWebResponse)
                {
                    Encoding encode = Encoding.GetEncoding("utf-8");
                    StreamReader reader = new StreamReader(response.GetResponseStream(), encode);
                    json = reader.ReadToEnd();
                    dynamic obj = JObject.Parse(json);
                    itemCount = Convert.ToInt32(obj.d.results[0].ID.ToString());
                }


                return itemCount;
            }
            catch (Exception e)
            {
                return itemCount;
            }
        }

        public static bool GetAllListItems(ClientContext ctx, string token)
        {
            string siteURL = Constants.siteURL.ToString();
            string listName = Constants.MailContentListDisplayName;
            string filterCondition = "(Status eq '" + Constants.FilterItemStatus + "')";

            List<object> objlist = new List<object>();
            int RestAPIItemThreshold = 5000;
            long totalCount = 0;
            try
            {
                long itemCount = GetMaxIDValue(siteURL, listName, "", token);
                long startID = Convert.ToInt32(Constants.ItemStartId);
                long endID = RestAPIItemThreshold;
                int loop = 1;

                if (itemCount > RestAPIItemThreshold)
                {
                    loop = (int)((itemCount % RestAPIItemThreshold == 0) ? itemCount / RestAPIItemThreshold : itemCount / RestAPIItemThreshold + 1);
                }
                for (int i = 0; i < loop; i++)
                {
                    string apiURL = siteURL.TrimEnd('/');
                    apiURL = String.Concat(apiURL, "/_api/web/");
                    apiURL = String.Concat(apiURL, "lists/getByTitle('", listName, "')/items");

                    string filterQuery = string.Empty;

                    if (string.IsNullOrEmpty(filterCondition))
                    {
                        filterQuery = String.Concat("?$filter=", "(ID ge ", startID, ")and (ID le ", endID, ")");
                    }
                    else
                    {
                        filterQuery = String.Concat("?$filter=", "(", filterCondition, ") and (ID gt ", startID, ")and (ID le ", endID, ")");
                    }

                    apiURL = string.Concat(apiURL, filterQuery, "&$top=", RestAPIItemThreshold);

                    HttpWebRequest GetAllListItemsRequest = HttpWebRequest.Create(apiURL) as HttpWebRequest;
                    GetAllListItemsRequest.Method = "GET";//"POST";// "DELETE"; //
                    GetAllListItemsRequest.ContentType = "application/json; odata=verbose";
                    GetAllListItemsRequest.Accept = "application/json; odata=verbose";
                    GetAllListItemsRequest.ContentLength = 0;
                    GetAllListItemsRequest.Headers.Add("Authorization", string.Concat("Bearer ", token));
                    using (HttpWebResponse response = GetAllListItemsRequest.GetResponse() as HttpWebResponse)
                    {
                        Encoding encode = Encoding.GetEncoding("utf-8");
                        StreamReader reader = new StreamReader(response.GetResponseStream(), encode);
                        string json = reader.ReadToEnd();
                        dynamic obj = JObject.Parse(json);
                        itemCount = Convert.ToInt32(((Newtonsoft.Json.Linq.JContainer)(obj.d.results)).Count);
                        JObject emailContents = JObject.Parse(json);
                        var totalEmailContent = emailContents["d"]["results"];
                        Console.WriteLine("Total Items: " + totalEmailContent.Count());
                        Helper.WriteLog($"{listName}" + "," + $"{""}" + "," + $"{""}" + "," + $"{""}" + "," + $"{""}" + "," + $"{""}" + "," + $"{""}" + "," + $"{"ListProcessing"}" + "," + $"{""}");

                        foreach (var currItem in totalEmailContent)
                        {
                            SharedUserList = new Dictionary<string, string>();
                            DirectUsersList = new Dictionary<string, string>();

                            var emailStatus = currItem["Status"].ToString();
                            var emailitemId = currItem["ID"].ToString();
                            var emailTitle = currItem["Title"].ToString();
                            var emailToAddress = currItem["ToAddress"].ToString();
                            var emailCCAddress = currItem["CCAddress"].ToString();
                            var emailBCCAddress = currItem["BCCAddress"].ToString();
                            var emailSubject = currItem["EmailSubject"].ToString();
                            var emailContent = currItem["EmailContent"].ToString();
                            var itemListName = currItem["ListName"].ToString();
                            var itemRequestType = currItem["RequestType"].ToString();
                            //  var SupplierMailItemId = currItem["SupplierMailItemId"].ToString();
                            var ItemId = currItem["ItemId"].ToString();
                            var FlowRunId = currItem["FlowRunId"].ToString();
                            var requestorName = currItem["RequestorName"].ToString();
                            var requestorMail = currItem["RequestorEmail"].ToString();

                            Console.WriteLine("Processing For ID:" + emailitemId + ", Title: " + emailTitle);
                            Helper.WriteLog($"{listName}" + "," + $"{emailitemId}" + "," + $"{emailToAddress}" + "," + $"{emailCCAddress}" + "," + $"{emailBCCAddress}" + "," + $"{itemRequestType}" + "," + $"{"ItemProcessing"}" + "," + $"{"Success"}" + "," + $"{"Title: " + emailTitle}");

                            if (emailStatus.ToLower() == Constants.FilterItemStatus.ToLower())
                            {
                                bool CheckSupplierMailStatus = true;

                                //First check for Spplier email sentStatus if the request status is Approver
                                if (itemRequestType.ToLower() == Constants.ApproverMailRequestType.ToLower())
                                {
                                    CheckSupplierMailStatus = Helper.GetSupplieMailStatus(emailitemId, ItemId, token);
                                }

                                if (CheckSupplierMailStatus)
                                {
                                    Helper.WriteLog($"{listName}" + "," + $"{emailitemId}" + "," + $"{emailToAddress}" + "," + $"{emailCCAddress}" + "," + $"{emailBCCAddress}" + "," + $"{itemRequestType}" + "," + $"{"CheckSupplierMailStatus"}" + "," + $"{"Success"}" + "," + $"{"SupplierItem:" + ItemId + " ;SentMailToSupplier"}");
                                    var sMailAnchorTag = "<a href=\"\">click here</a>";
                                    var sMailAnchorTag2 = "<a href=\"\">click here </a>";
                                    string sAttachmentLink = "";
                                    string triggerEmailStatus = "";
                                    string sExceptionMessage = string.Empty;

                                    var directLinkSendMsg = "";
                                    var sharedLinkSendMsg = "";

                                    //EmailContent Control Change 3-Jan-2022
                                    // if ((emailContent.ToLower().Contains(sMailAnchorTag) || emailContent.ToLower().Contains(sMailAnchorTag2)) && itemRequestType.ToLower() == Constants.SuppliermailRequestType.ToLower())
                                    if (itemRequestType.ToLower() == Constants.SuppliermailRequestType.ToLower())
                                    {
                                        sAttachmentLink = GetItemfromMailMissendList(ItemId, itemListName, token);
                                        if (itemListName == Constants.Group1ListName)
                                        {
                                            sAttachmentLink = sAttachmentLink.Replace(Constants.Group1Teams, Constants.Group1Site);
                                        }
                                        else if (itemListName == Constants.Group2ListName)
                                        {
                                            sAttachmentLink = sAttachmentLink.Replace(Constants.Group2Teams, Constants.Group2Site);
                                        }
                                        else if (itemListName == Constants.Group3ListName)
                                        {
                                            sAttachmentLink = sAttachmentLink.Replace(Constants.Group3Teams, Constants.Group3Site);
                                        }
                                        else if (itemListName == Constants.Group4ListName)
                                        {
                                            sAttachmentLink = sAttachmentLink.Replace(Constants.Group4Teams, Constants.Group4Site);
                                        }
                                        Helper.WriteLog($"{listName}" + "," + $"{emailitemId}" + "," + $"{emailToAddress}" + "," + $"{emailCCAddress}" + "," + $"{emailBCCAddress}" + "," + $"{itemRequestType}" + "," + $"{"AttachmentCheck"}" + "," + $"{"Success"}" + "," + $"{"sAttachmentLink:" + sAttachmentLink + ""}");

                                        Constants.DirectLinkUsers = "";
                                        Constants.SharedLinkUsers = "";

                                        // to get the total number of href/documents in the mainContent
                                        Regex nn = new Regex("(?:href)=[\"|']?(.*?)[\"|'|>]+", RegexOptions.Singleline | RegexOptions.CultureInvariant | RegexOptions.IgnoreCase);
                                        if (nn.IsMatch(sAttachmentLink))
                                        {
                                            Console.WriteLine("total Document links : " + nn.Matches(sAttachmentLink).Count);
                                            if (nn.Matches(sAttachmentLink).Count > 0)
                                            {
                                                foreach (Match match in nn.Matches(sAttachmentLink))
                                                {
                                                    string currDocument = match.Groups[1].Value.ToString();
                                                    currDocument = currDocument.Replace("href=\"", "").Replace("\"", "").Replace("\"\\", "\"").Trim();
                                                    //Console.WriteLine(currDocument);
                                                    Helper.AccessGrantForCCEmailAddress(emailCCAddress, currDocument, itemListName, emailitemId);
                                                    Helper.SegregateToEmailASDirectShared(emailToAddress, currDocument, itemListName, emailitemId);
                                                }

                                            }
                                            else
                                            {
                                                Console.WriteLine("There is no documents for this MailItem : " + emailTitle);
                                                Helper.WriteLog($"{listName}" + "," + $"{emailitemId}" + "," + $"{emailToAddress}" + "," + $"{emailCCAddress}" + "," + $"{emailBCCAddress}" + "," + $"{itemRequestType}" + "," + $"{"AttachmentCheck"}" + "," + $"{"Success"}" + "," + $"{"There is no documents for :" + emailTitle + ""}");
                                            }
                                        }

                                        if (DirectUsersList.Count > 0)
                                        {
                                            Console.WriteLine("Processing for Direct link Users");
                                            var DirectDocURL = "";
                                            var directLinksEmailFormat = "";
                                            foreach (KeyValuePair<string, string> DLLinks in DirectUsersList)
                                            {
                                                DirectDocURL = DLLinks.Key.ToString();
                                                if (string.IsNullOrEmpty(Constants.DirectLinkUsers))
                                                {
                                                    Constants.DirectLinkUsers += DLLinks.Value.ToString();
                                                }
                                                else
                                                {
                                                    var DirectUsersTrim = DLLinks.Value.ToString().TrimEnd(';');
                                                    var DirectUsers = DirectUsersTrim.Split(new string[] { ";" }, StringSplitOptions.None);
                                                    foreach (var currDirectUserEmail in DirectUsers)
                                                    {
                                                        if (!Constants.DirectLinkUsers.ToLower().Contains(currDirectUserEmail.ToString().ToLower()))
                                                        {
                                                            Constants.DirectLinkUsers += ";" + currDirectUserEmail.ToString();
                                                        }
                                                        else
                                                        {
                                                            //Console.WriteLine("already added directUsers to DirectLinkUsers: " + currDirectUserEmail);
                                                        }
                                                    }

                                                }
                                                string sfilename2 = HttpUtility.UrlDecode(GetFilenameFromUrl(DirectDocURL.ToString()));
                                                directLinksEmailFormat += "<a href=\"" + DirectDocURL.ToString() + "\">" + sfilename2 + "</a><br>";
                                            }


                                            var directemailContent = emailContent;
                                            directemailContent = directemailContent.Replace(sMailAnchorTag, directLinksEmailFormat);
                                            directemailContent = directemailContent.Replace(sMailAnchorTag2, directLinksEmailFormat);
                                            //Helper.WriteLog_EmailContent("Started Email Trigger to DirectLinkUsers\n-------------------------------------------------------------");                                            
                                            directLinkSendMsg = Helper.TriggerEmail(Constants.DirectLinkUsers, emailCCAddress, emailBCCAddress, emailSubject, directemailContent, itemRequestType, out sExceptionMessage, directLinksEmailFormat, requestorName, requestorMail, ctx);
                                            //Helper.WriteLog_EmailContent("Completed Email Trigger to DirectLinkUsers\n-----------------------------------------------------------");                                            

                                            if (directLinkSendMsg.ToLower().Contains("allmailssent"))
                                            {
                                                triggerEmailStatus += "DirectLinkMailCompleted,";
                                                Helper.WriteLog($"{listName}" + "," + $"{emailitemId}" + "," + $"{Constants.DirectLinkUsers}" + "," + $"{emailCCAddress}" + "," + $"{emailBCCAddress}" + "," + $"{itemRequestType}" + "," + $"{"TriggerMailToDirectLink"}" + "," + $"{"Success"}" + "," + $"{"MailSent :" + directLinksEmailFormat + ""}");
                                            }
                                            else
                                            {
                                                triggerEmailStatus += "DirectLinkMail:" + directLinkSendMsg + ",";
                                                Helper.WriteLog($"{listName}" + "," + $"{emailitemId}" + "," + $"{Constants.DirectLinkUsers}" + "," + $"{emailCCAddress}" + "," + $"{emailBCCAddress}" + "," + $"{itemRequestType}" + "," + $"{"TriggerMailToDirectLink"}" + "," + $"{"Error"}" + "," + $"{directLinksEmailFormat + ""}");
                                            }
                                        }
                                        else
                                        {
                                            triggerEmailStatus += "NoDirectUsersFound,";
                                            Helper.WriteLog($"{listName}" + "," + $"{emailitemId}" + "," + $"{Constants.DirectLinkUsers}" + "," + $"{emailCCAddress}" + "," + $"{emailBCCAddress}" + "," + $"{itemRequestType}" + "," + $"{"TriggerMailToDirectLink"}" + "," + $"{"Sucess"}" + "," + $"{"NoDirectUsersFound"}");
                                        }

                                        if (SharedUserList.Count > 0)
                                        {
                                            Console.WriteLine("Processing for Shared link Users");
                                            var SharedDocURL = "";
                                            var sharedLinksEmailFormat = "";
                                            foreach (KeyValuePair<string, string> SharedLinks in SharedUserList)
                                            {
                                                SharedDocURL = SharedLinks.Key.ToString();
                                                var sFileName = "click here";
                                                if (string.IsNullOrEmpty(Constants.SharedLinkUsers))
                                                {
                                                    var SharedUsersTrim = SharedLinks.Value.ToString().TrimEnd(';');
                                                    var sSharedUsers = SharedUsersTrim.Split(new string[] { ";" }, StringSplitOptions.None);
                                                    var SharedUsers = sSharedUsers[0].ToString().Split(new string[] { "##@@FileName@@##" }, StringSplitOptions.None)[0];
                                                    sFileName = SharedUsersTrim.Split(new string[] { "##@@FileName@@##" }, StringSplitOptions.None)[1];

                                                    Constants.SharedLinkUsers += SharedUsers;
                                                }
                                                else
                                                {
                                                    var SharedUsersTrim = SharedLinks.Value.ToString().TrimEnd(';');
                                                    var SharedUsers = SharedUsersTrim.Split(new string[] { ";" }, StringSplitOptions.None);


                                                    foreach (var scurrSharedUserEmail in SharedUsers)
                                                    {
                                                        var currSharedUserEmail = scurrSharedUserEmail.Split(new string[] { "##@@FileName@@##" }, StringSplitOptions.None)[0];
                                                        sFileName = SharedUsersTrim.Split(new string[] { "##@@FileName@@##" }, StringSplitOptions.None)[1];

                                                        if (!Constants.SharedLinkUsers.ToLower().Contains(currSharedUserEmail.ToString().ToLower()))
                                                        {
                                                            Constants.SharedLinkUsers += ";" + SharedLinks.Value.ToString();
                                                        }
                                                        else
                                                        {
                                                            //Console.WriteLine("already added Shared Users to SahredLinkUsers: " + currSharedUserEmail);
                                                        }
                                                    }

                                                }

                                                sharedLinksEmailFormat += "<a href=\"" + SharedDocURL.ToString() + "\">" + sFileName + "</a><br>";
                                            }
                                            var SharedemailContent = emailContent.Replace(sMailAnchorTag, sharedLinksEmailFormat);
                                            SharedemailContent = SharedemailContent.Replace(sMailAnchorTag2, sharedLinksEmailFormat);
                                            //Helper.WriteLog_EmailContent("Started Email Trigger to SharedLinkUsers\n---------------------------------------------------------------");                                                                                        
                                            sharedLinkSendMsg = Helper.TriggerEmail(Constants.SharedLinkUsers, emailCCAddress, emailBCCAddress, emailSubject, SharedemailContent, itemRequestType, out sExceptionMessage, sharedLinksEmailFormat, requestorName, requestorMail, ctx);
                                            //Helper.WriteLog_EmailContent("Completed Email Trigger to SharedLinkUsers\n---------------------------------------------------------------");                                            

                                            if (sharedLinkSendMsg.ToLower().Contains("allmailssent"))
                                            {
                                                triggerEmailStatus += "SharedLinkMailCompleted";
                                                Helper.WriteLog($"{listName}" + "," + $"{emailitemId}" + "," + $"{Constants.SharedLinkUsers}" + "," + $"{emailCCAddress}" + "," + $"{emailBCCAddress}" + "," + $"{itemRequestType}" + "," + $"{"TriggerMailToSharedLink"}" + "," + $"{"Success"}" + "," + $"{"MailSent :" + sharedLinksEmailFormat + ""}");
                                            }
                                            else
                                            {
                                                triggerEmailStatus += "SharedLinkMail:" + sharedLinkSendMsg + ",";
                                                Helper.WriteLog($"{listName}" + "," + $"{emailitemId}" + "," + $"{Constants.SharedLinkUsers}" + "," + $"{emailCCAddress}" + "," + $"{emailBCCAddress}" + "," + $"{itemRequestType}" + "," + $"{"TriggerMailToSharedLink"}" + "," + $"{"Error"}" + "," + $"{sharedLinksEmailFormat + ""}");
                                            }
                                        }
                                        else
                                        {
                                            triggerEmailStatus += "NoSharedUsersFound,";
                                            Helper.WriteLog($"{listName}" + "," + $"{emailitemId}" + "," + $"{Constants.SharedLinkUsers}" + "," + $"{emailCCAddress}" + "," + $"{emailBCCAddress}" + "," + $"{itemRequestType}" + "," + $"{"TriggerMailToSharedLink"}" + "," + $"{"Sucess"}" + "," + $"{"NoSharedUsersFound"}");
                                        }

                                    }

                                    else
                                    {
                                        Helper.WriteLog_EmailContent("Started Email Trigger to Users\n---------------------------------------------------------------");
                                        triggerEmailStatus = Helper.TriggerEmail(emailToAddress, emailCCAddress, emailBCCAddress, emailSubject, emailContent, itemRequestType, out sExceptionMessage, sAttachmentLink, "", "", ctx);
                                        Helper.WriteLog_EmailContent("Completed Email Trigger to Users\n---------------------------------------------------------------");
                                    }

                                    //to update the Email sent Staus, sent mail contents of sharedLink, Direct lInk in Send Mail list.
                                    string SeperateMailSentStatus = Helper.UpdateDirectSharedMailSentStatusInList(listName, emailitemId, directLinkSendMsg, sharedLinkSendMsg, token);

                                    if (SeperateMailSentStatus.ToLower() == "success")
                                    {
                                        Helper.WriteLog($"{listName}" + "," + $"{emailitemId}" + "," + $"{Constants.SharedLinkUsers}" + "," + $"{emailCCAddress}" + "," + $"{emailBCCAddress}" + "," + $"{itemRequestType}" + "," + $"{"DirectSharedColumnUpdate"}" + "," + $"{"Success"}" + "," + $"{""}");
                                    }
                                    else
                                    {
                                        Helper.WriteLog($"{listName}" + "," + $"{emailitemId}" + "," + $"{Constants.SharedLinkUsers}" + "," + $"{emailCCAddress}" + "," + $"{emailBCCAddress}" + "," + $"{itemRequestType}" + "," + $"{"DirectSharedColumnUpdate"}" + "," + $"{"Error"}" + "," + $"{SeperateMailSentStatus}");
                                    }
                                    var MailMisSendReqStatus = "";//NoDirectUsersFound
                                    if (triggerEmailStatus.ToLower().Contains("success;allmailssent") || (triggerEmailStatus.ToLower().Contains("directlinkmailcompleted,sharedlinkmailcompleted")) || (triggerEmailStatus.ToLower().Contains("directlinkmailcompleted,nosharedusersfound")) || (triggerEmailStatus.ToLower().Contains("nodirectusersfound,sharedlinkmailcompleted")))
                                    {
                                        string updateStatus = Helper.UpdateListItemSendMailList(Constants.MailContentListDisplayName, emailitemId, "EmailSentSuccessFully", emailContent, "Completed", token);
                                        if (updateStatus.ToLower() == "success")
                                        {
                                            Helper.WriteLog($"{listName}" + "," + $"{emailitemId}" + "," + $"{emailToAddress}" + "," + $"{emailCCAddress}" + "," + $"{emailBCCAddress}" + "," + $"{itemRequestType}" + "," + $"{"SendmailItemStatusChange"}" + "," + $"{"Success"}" + "," + $"{""}");
                                            MailMisSendReqStatus = UpdateListItemMailMissendRequest(itemListName, ItemId, Constants.MailSentByAMO, token);
                                        }
                                        else
                                        {
                                            Helper.WriteLog($"{listName}" + "," + $"{emailitemId}" + "," + $"{emailToAddress}" + "," + $"{emailCCAddress}" + "," + $"{emailBCCAddress}" + "," + $"{itemRequestType}" + "," + $"{"SendmailItemStatusChange"}" + "," + $"{"Error"}" + "," + $"{updateStatus}");
                                        }
                                    }
                                    else
                                    {
                                        string updateStatus = Helper.UpdateListItemSendMailList(Constants.MailContentListDisplayName, emailitemId, triggerEmailStatus, emailContent, "MailFailure", token);
                                        if (updateStatus.ToLower() == "success")
                                        {
                                            Helper.WriteLog($"{listName}" + "," + $"{emailitemId}" + "," + $"{emailToAddress}" + "," + $"{emailCCAddress}" + "," + $"{emailBCCAddress}" + "," + $"{itemRequestType}" + "," + $"{"SendmailItemStatusChange"}" + "," + $"{"Error"}" + "," + $"{"FailureInMailSend: Updated the sendmailStatus as MailFailure"}");
                                            MailMisSendReqStatus = UpdateListItemMailMissendRequest(itemListName, ItemId, Constants.MailSentByAMO, token);
                                        }
                                        else
                                        {
                                            Helper.WriteLog(itemListName + "," + emailitemId + "," + emailToAddress + "," + emailCCAddress + "," + emailBCCAddress + "," + itemRequestType + ",Error," + "FailedToUpdate: " + updateStatus);
                                        }
                                    }

                                    if (MailMisSendReqStatus.ToLower() == "success")
                                    {
                                        Helper.WriteLog($"{listName}" + "," + $"{emailitemId}" + "," + $"{emailToAddress}" + "," + $"{emailCCAddress}" + "," + $"{emailBCCAddress}" + "," + $"{itemRequestType}" + "," + $"{"UpdateListItemRequestStatus"}" + "," + $"{"Success"}" + "," + $"{""}");
                                    }
                                    else
                                    {
                                        Helper.WriteLog($"{listName}" + "," + $"{emailitemId}" + "," + $"{emailToAddress}" + "," + $"{emailCCAddress}" + "," + $"{emailBCCAddress}" + "," + $"{itemRequestType}" + "," + $"{"UpdateListItemRequestStatus"}" + "," + $"{"Error"}" + "," + $"{MailMisSendReqStatus}");
                                    }
                                }
                                else
                                {
                                    Helper.WriteLog($"{listName}" + "," + $"{emailitemId}" + "," + $"{emailToAddress}" + "," + $"{emailCCAddress}" + "," + $"{emailBCCAddress}" + "," + $"{itemRequestType}" + "," + $"{"CheckSupplierMailStatus"}" + "," + $"{"Error"}" + "," + $"{"SupplierItem:" + ItemId + " ;NotSentMailToSupplier"}");
                                }

                            }
                            else
                            {
                                Helper.WriteLog($"{listName}" + "," + $"{emailitemId}" + "," + $"{emailToAddress}" + "," + $"{emailCCAddress}" + "," + $"{emailBCCAddress}" + "," + $"{itemRequestType}" + "," + $"{"ItemStatusCheck"}" + "," + $"{"Error"}" + "," + $"{"ItemStatus is :" + emailStatus.ToLower() + "instead of: " + Constants.FilterItemStatus.ToLower()}");
                            }
                            Console.WriteLine("Completed:" + emailitemId + ", Title: " + emailTitle);
                        }

                        totalCount = totalCount + itemCount;
                    }
                    startID += RestAPIItemThreshold;
                    endID += RestAPIItemThreshold;
                }
                return true;
            }

            catch (Exception ex)
            {
                Helper.WriteLog("GetAllListItems - " + ex.Message + ": " + ex.StackTrace);
                return false;
            }
        }

        public static string GetFilenameFromUrl(string url)
        {
            return String.IsNullOrEmpty(url.Trim()) || !url.Contains(".") ? string.Empty : Path.GetFileName(new Uri(url).AbsolutePath);
        }

        public static bool GetSupplieMailStatus(string sendEmailItemID, string mailItemID, string accessToken)
        {
            bool sResult = false;
            string filterQuery = string.Empty;
            filterQuery = String.Concat("?$filter=", "((ItemId eq '", mailItemID, "') and (Status eq '", "Completed", "') or (Status eq '", "MailFailure", "'))");
            string APIURL = String.Concat(Constants.siteURL.ToString().TrimEnd('/'), "/_api/web/lists/getbytitle('", Constants.MailContentListDisplayName, "')/items");
            APIURL = string.Concat(APIURL, filterQuery);
            try
            {
                HttpWebRequest GetAllListItemsRequest = HttpWebRequest.Create(APIURL) as HttpWebRequest;
                GetAllListItemsRequest.Method = "GET";//"POST";// "DELETE"; //
                GetAllListItemsRequest.ContentType = "application/json; odata=verbose";
                GetAllListItemsRequest.Accept = "application/json; odata=verbose";
                GetAllListItemsRequest.ContentLength = 0;
                GetAllListItemsRequest.Headers.Add("Authorization", string.Concat("Bearer ", accessToken));
                using (HttpWebResponse response = GetAllListItemsRequest.GetResponse() as HttpWebResponse)
                {
                    Encoding encode = Encoding.GetEncoding("utf-8");
                    StreamReader reader = new StreamReader(response.GetResponseStream(), encode);
                    string json = reader.ReadToEnd();

                    JObject emailContents = JObject.Parse(json);
                    var supplierEmailItems = emailContents["d"]["results"];
                    foreach (var supplierEmailItem in supplierEmailItems)
                    {
                        var emailStatus = supplierEmailItem["Status"].ToString();
                        var itemRequestType = supplierEmailItem["RequestType"].ToString();

                        if (itemRequestType.ToLower() == Constants.SuppliermailRequestType.ToLower())
                        {
                            if (emailStatus.ToLower() == Constants.CompletedItemStatus.ToLower() || emailStatus.ToLower() == Constants.MailFailureItemStatus.ToLower())
                            {
                                sResult = true;
                            }
                        }
                    }
                }
                return sResult;
            }
            catch (Exception ex)
            {
                Helper.WriteLog($"{""}" + "," + $"{sendEmailItemID}" + "," + $"{""}" + "," + $"{""}" + "," + $"{""}" + "," + $"{""}" + "," + $"{"CheckSupplierMailStatus"}" + "," + $"{"Error"}" + "," + $"{"SupplierItem:" + mailItemID + " ;" + ex.Message + ""}");
                return false;
            }
        }

        public static string GetItemfromMailMissendList(string sMailMissendItemId, string sMailMissendListName, string accessToken)
        {
            bool sResult = false;
            var sAttachmentLink = "";
            string filterQuery = string.Empty;
            filterQuery = String.Concat("?$filter=", "(Id eq '", sMailMissendItemId, "')");
            string APIURL = String.Concat(Constants.siteURL.ToString().TrimEnd('/'), "/_api/web/lists/getbytitle('", sMailMissendListName, "')/items");
            APIURL = string.Concat(APIURL, filterQuery);
            try
            {
                HttpWebRequest GetAllListItemsRequest = HttpWebRequest.Create(APIURL) as HttpWebRequest;
                GetAllListItemsRequest.Method = "GET";//"POST";// "DELETE"; //
                GetAllListItemsRequest.ContentType = "application/json; odata=verbose";
                GetAllListItemsRequest.Accept = "application/json; odata=verbose";
                GetAllListItemsRequest.ContentLength = 0;
                GetAllListItemsRequest.Headers.Add("Authorization", string.Concat("Bearer ", accessToken));
                using (HttpWebResponse response = GetAllListItemsRequest.GetResponse() as HttpWebResponse)
                {
                    Encoding encode = Encoding.GetEncoding("utf-8");
                    StreamReader reader = new StreamReader(response.GetResponseStream(), encode);
                    string json = reader.ReadToEnd();

                    JObject emailContents = JObject.Parse(json);
                    var supplierEmailItems = emailContents["d"]["results"];
                    foreach (var supplierEmailItem in supplierEmailItems)
                    {
                        sAttachmentLink = supplierEmailItem["AttachmentLink"].ToString();
                    }
                }
                return sAttachmentLink;
            }
            catch (Exception ex)
            {
                return string.Empty;
            }
        }


        public static string UpdateListItemSendMailList(string listName, string itemID, string batchComments, string sEmailContent, string Status, string accessToken)
        {
            string sResult = "";
            string APIURL = String.Concat(Constants.siteURL.ToString().TrimEnd('/'), "/_api/web/lists/getbytitle('", listName, "')/items(" + itemID + ")");
            try
            {
                var batchComment_EmailContent = batchComments + "\n\nEmailContent:\n" + sEmailContent;
                StringBuilder uPdateData = new StringBuilder();
                uPdateData.Append(",'" + "Status" + "' : '" + JsonConvert.SerializeObject(Status).ToString().Trim('"') + "'");
                uPdateData.Append(",'" + "BatchComments" + "' : '" + JsonConvert.SerializeObject(batchComments).ToString().Trim('"') + "'");
                //uPdateData.Append(",'" + "EmailContent" + "' : '" + JsonConvert.SerializeObject(sEmailContent).ToString().Trim('"') + "'");

                uPdateData.Append("}");

                string stringData = String.Concat("{'__metadata': { 'type': 'SP.Data.", listName, "ListItem", "'}");
                HttpWebRequest endpointRequest = (HttpWebRequest)HttpWebRequest.Create(APIURL);
                endpointRequest.Headers.Add("Authorization", string.Concat("Bearer ", accessToken));
                endpointRequest.Method = "POST";
                endpointRequest.Accept = "application/json; odata=verbose";
                endpointRequest.ContentType = "application/json; odata=verbose";
                endpointRequest.Headers.Add("X-HTTP-Method", "MERGE");
                endpointRequest.Headers.Add("IF-MATCH", "*");
                stringData = String.Concat(stringData, uPdateData.ToString());
                endpointRequest.ContentLength = Encoding.UTF8.GetByteCount(stringData);
                StreamWriter writer = new StreamWriter(endpointRequest.GetRequestStream());
                writer.Write(stringData);
                writer.Flush();
                using (WebResponse wresp = endpointRequest.GetResponse())
                {
                    sResult = "Success";
                }
                return sResult;
            }
            catch (Exception ee)
            {
                sResult = ee.Message;
                return sResult;
            }

        }

        public static string UpdateDirectSharedMailSentStatusInList(string listName, string itemID, string DirectLinkMailContent_Status, string SharedLinkMailContent_Status, string accessToken)
        {
            string sResult = "";
            string APIURL = String.Concat(Constants.siteURL.ToString().TrimEnd('/'), "/_api/web/lists/getbytitle('", listName, "')/items(" + itemID + ")");
            try
            {
                StringBuilder uPdateData = new StringBuilder();
                uPdateData.Append(",'" + "DirectLinkUsers" + "' : '" + JsonConvert.SerializeObject(DirectLinkMailContent_Status).ToString().Trim('"') + "'");
                uPdateData.Append(",'" + "SharedLinkUsers" + "' : '" + JsonConvert.SerializeObject(SharedLinkMailContent_Status).ToString().Trim('"') + "'");
                uPdateData.Append("}");

                string stringData = String.Concat("{'__metadata': { 'type': 'SP.Data.", listName, "ListItem", "'}");
                HttpWebRequest endpointRequest = (HttpWebRequest)HttpWebRequest.Create(APIURL);
                endpointRequest.Headers.Add("Authorization", string.Concat("Bearer ", accessToken));
                endpointRequest.Method = "POST";
                endpointRequest.Accept = "application/json; odata=verbose";
                endpointRequest.ContentType = "application/json; odata=verbose";
                endpointRequest.Headers.Add("X-HTTP-Method", "MERGE");
                endpointRequest.Headers.Add("IF-MATCH", "*");
                stringData = String.Concat(stringData, uPdateData.ToString());
                endpointRequest.ContentLength = Encoding.UTF8.GetByteCount(stringData);
                StreamWriter writer = new StreamWriter(endpointRequest.GetRequestStream());
                writer.Write(stringData);
                writer.Flush();
                using (WebResponse wresp = endpointRequest.GetResponse())
                {
                    sResult = "Success";
                }
                return sResult;
            }
            catch (Exception ee)
            {
                sResult = ee.Message;
                return sResult;
            }

        }

        //public static string UpdateDirectSharedMailSentStatusInList(string listName, string itemID, string DirectLinkMailContent_Status, string SharedLinkMailContent_Status, string accessToken)
        //{
        //    string sResult = "";
        //    string APIURL = String.Concat(Constants.siteURL.ToString().TrimEnd('/'), "/_api/web/lists/getbytitle('", listName, "')/items(" + itemID + ")");
        //    try
        //    {
        //        StringBuilder uPdateData = new StringBuilder();
        //        uPdateData.Append(",'" + "DirectLinkUsers" + "' : '" + JsonConvert.SerializeObject(DirectLinkMailContent_Status).ToString().Trim('"') + "'");                
        //        uPdateData.Append(",'" + "SharedLinkUsers" + "' : '" + JsonConvert.SerializeObject(SharedLinkMailContent_Status).ToString().Trim('"') + "'");                                                 
        //        uPdateData.Append("}");

        //        string stringData = String.Concat("{'__metadata': { 'type': 'SP.Data.", listName, "ListItem", "'}");
        //        HttpWebRequest endpointRequest = (HttpWebRequest)HttpWebRequest.Create(APIURL);
        //        endpointRequest.Headers.Add("Authorization", string.Concat("Bearer ", accessToken));
        //        endpointRequest.Method = "POST";
        //        endpointRequest.Accept = "application/json; odata=verbose";
        //        endpointRequest.ContentType = "application/json; odata=verbose";
        //        endpointRequest.Headers.Add("X-HTTP-Method", "MERGE");
        //        endpointRequest.Headers.Add("IF-MATCH", "*");
        //        stringData = String.Concat(stringData, uPdateData.ToString());
        //        endpointRequest.ContentLength = Encoding.UTF8.GetByteCount(stringData);
        //        //StreamWriter writer = new StreamWriter(endpointRequest.GetRequestStream());
        //        //writer.Write(stringData);
        //        //writer.Flush();
        //        using (WebResponse wresp = endpointRequest.GetResponse())
        //        {
        //            sResult = "Success";
        //        }
        //        return sResult;
        //    }
        //    catch (Exception ee)
        //    {
        //        sResult = ee.Message;
        //        return sResult;
        //    }

        //}
        public static string SegregateToEmailASDirectShared(string toNotifyAddress, string sAttachmentLinks, string itemListName, string emailitemId)
        {
            var emailSeperateComments = "";
            try
            {
                var ToAddress = toNotifyAddress.TrimEnd(';');
                var toNotify = ToAddress.Split(new string[] { ";" }, StringSplitOptions.None);

                for (var t = 0; t < toNotify.Length; t++)
                {
                    if (toNotify[t].TrimStart().TrimEnd() != "" && toNotify[t].TrimStart().TrimEnd() != "undefined")
                    {
                        if (IsValidEmail(toNotify[t].TrimStart().TrimEnd()))
                        {
                            var currUserEmailAddress = toNotify[t].TrimStart().TrimEnd();
                            var currLink = Helper.GetCurrentEmailSharingResult(emailitemId, toNotify[t].TrimStart().TrimEnd(), sAttachmentLinks, itemListName);
                            var shareDocumentStatus = Helper.AddDocumentLinksToDictionary(emailitemId, currLink, sAttachmentLinks, currUserEmailAddress);

                            if (shareDocumentStatus == "success")
                            {
                                //Console.WriteLine("SharedUser: " + currUserEmailAddress);
                            }
                            else
                            {
                                //Console.WriteLine("Error : " + currUserEmailAddress + " : "+ shareDocumentStatus);
                                Helper.WriteLog($"{itemListName}" + "," + $"{emailitemId}" + "," + $"{""}" + "," + $"{""}" + "," + $"{""}" + "," + $"{""}" + "," + $"{"EmailSegregate"}" + "," + $"{"Error"}" + "," + $"{ currUserEmailAddress + ":" + shareDocumentStatus + ""}");
                            }
                        }
                        else
                        {
                            emailSeperateComments += "IsNotEmailFomrat:" + toNotify[t].TrimStart().TrimEnd() + ",";
                            Helper.WriteLog($"{itemListName}" + "," + $"{emailitemId}" + "," + $"{""}" + "," + $"{""}" + "," + $"{""}" + "," + $"{""}" + "," + $"{"EmailSegregate"}" + "," + $"{"Error"}" + "," + $"{"IsNotEmailFomrat:" + toNotify[t].TrimStart().TrimEnd() + ""}");
                        }
                    }
                }

                return emailSeperateComments;
            }
            catch (Exception ex)
            {
                Helper.WriteLog($"{itemListName}" + "," + $"{emailitemId}" + "," + $"{""}" + "," + $"{""}" + "," + $"{""}" + "," + $"{""}" + "," + $"{"EmailSegregate"}" + "," + $"{"Error"}" + "," + $"{ex.Message}");
                //Console.WriteLine("Error in Email SegregateToEmailASDirectShared");
                return "";
            }
        }

        public static string AccessGrantForCCEmailAddress(string ccNotifyAddress, string sAttachmentLinks, string itemListName, string emailitemId)
        {
            var emailSeperateComments = "";
            try
            {
                var ccAddress = ccNotifyAddress.TrimEnd(';');
                var ccNotify = ccAddress.Split(new string[] { ";" }, StringSplitOptions.None);



                for (var t = 0; t < ccNotify.Length; t++)
                {
                    if (ccNotify[t].TrimStart().TrimEnd() != "" && ccNotify[t].TrimStart().TrimEnd() != "undefined")
                    {
                        if (IsValidEmail(ccNotify[t].TrimStart().TrimEnd()))
                        {
                            var currUserEmailAddress = ccNotify[t].TrimStart().TrimEnd();
                            var currLink = Helper.GetCurrentEmailSharingResult(emailitemId, ccNotify[t].TrimStart().TrimEnd(), sAttachmentLinks, itemListName);
                            //var shareDocumentStatus = Helper.AddDocumentLinksToDictionary(emailitemId, currLink, sAttachmentLinks, currUserEmailAddress);



                            if (currLink.StatusCode.ToString() == "CompletedSuccessfully")
                            {
                                if (currLink.InvitedUsers != null)
                                {
                                    var SharedUrl = currLink.InvitedUsers[0].InvitationLink;



                                    Console.WriteLine(SharedUrl);
                                    var toRemoveEmail = "?email=" + currUserEmailAddress.Replace("@", "%40");
                                    var formattedURL = SharedUrl.Replace(toRemoveEmail.ToString(), "");
                                    Helper.WriteLog($"{itemListName}" + "," + $"{emailitemId}" + "," + $"{""}" + "," + $"{""}" + "," + $"{""}" + "," + $"{""}" + "," + $"{"AccessToCCEmail"}" + "," + $"{"Success"}" + "," + $"{"SharedLink: " + ccNotify[t].TrimStart().TrimEnd() + "; " + formattedURL + ""}");
                                }
                                else if (currLink.UniquelyPermissionedUsers.Count > 0)
                                {
                                    Helper.WriteLog($"{itemListName}" + "," + $"{emailitemId}" + "," + $"{""}" + "," + $"{""}" + "," + $"{""}" + "," + $"{""}" + "," + $"{"AccessToCCEmail"}" + "," + $"{"Success"}" + "," + $"{"DirectLink: " + ccNotify[t].TrimStart().TrimEnd() + "; " + sAttachmentLinks + ""}");
                                }
                                else if (currLink.UsersAddedToGroup != null)
                                {
                                    // to check if the user has permisison via group
                                    Helper.WriteLog($"{itemListName}" + "," + $"{emailitemId}" + "," + $"{""}" + "," + $"{""}" + "," + $"{""}" + "," + $"{""}" + "," + $"{"AccessToCCEmail"}" + "," + $"{"Success"}" + "," + $"{"MayBeAddedInGroup: " + ccNotify[t].TrimStart().TrimEnd() + ""}");
                                }
                            }
                            else if (currLink.StatusCode.ToString() == "AccessDenied")
                            {
                                Console.WriteLine("Error in ShareResult -  AccessDenied: Parameters are not correct");
                                Helper.WriteLog($"{itemListName}" + "," + $"{emailitemId}" + "," + $"{""}" + "," + $"{""}" + "," + $"{""}" + "," + $"{""}" + "," + $"{"AccessToCCEmail"}" + "," + $"{"Error"}" + "," + $"{"AccessDenied: " + ccNotify[t].TrimStart().TrimEnd() + currLink.StatusCode.ToString() + ""}");
                            }
                            else
                            {
                                //to do something
                            }
                        }
                        else
                        {
                            //DateTime,ListName,ItemID,ToEmails,CCEmails,BccEmails,RequestType,ActionType,Status,Comments
                            emailSeperateComments += "IsNotEmailFomrat:" + ccNotify[t].TrimStart().TrimEnd() + ",";
                            Helper.WriteLog($"{itemListName}" + "," + $"{emailitemId}" + "," + $"{""}" + "," + $"{""}" + "," + $"{""}" + "," + $"{""}" + "," + $"{"AccessToCCEmail"}" + "," + $"{"Error"}" + "," + $"{"IsNotEmailFomrat:" + ccNotify[t].TrimStart().TrimEnd() + ""}");
                        }
                    }
                }



                return emailSeperateComments;
            }
            catch (Exception ex)
            {
                Helper.WriteLog($"{itemListName}" + "," + $"{emailitemId}" + "," + $"{""}" + "," + $"{""}" + "," + $"{""}" + "," + $"{""}" + "," + $"{"AccessToCCEmail"}" + "," + $"{"Error"}" + "," + $"{ex.Message}");
                Console.WriteLine("Error in Access grant to CC email" + ex.Message);
                return "";
            }
        }
        public static string AddDocumentLinksToDictionary(string emailitemId, SharingResult currLink, string sAttachmentLinks, string currUserEmailAddress)
        {
            try
            {
                if (currLink.StatusCode.ToString() == "CompletedSuccessfully")
                {
                    // for sharedLink Users
                    if (currLink.InvitedUsers != null)
                    {
                        var SharedUrl = currLink.InvitedUsers[0].InvitationLink;
                        var sFileUrl = currLink.Url;
                        string sfilename2 = HttpUtility.UrlDecode(GetFilenameFromUrl(sFileUrl.ToString()));

                        Console.WriteLine(SharedUrl);
                        var toRemoveEmail = "?email=" + currUserEmailAddress.Replace("@", "%40");
                        var formattedURL = SharedUrl.Replace(toRemoveEmail.ToString(), "");

                        if (!SharedUserList.ContainsKey(formattedURL))
                        {
                            SharedUserList.Add(formattedURL, currUserEmailAddress + "##@@FileName@@##" + sfilename2);//adding new items to KeyValue pair
                        }
                        else if (SharedUserList.ContainsKey(formattedURL))
                        {
                            var currentEmails = SharedUserList[formattedURL];
                            if (currentEmails.Contains(currUserEmailAddress))
                            {
                                // to do something if the email address is already added  in the Key of thisdocument
                            }
                            else
                            {
                                currentEmails += ";" + currUserEmailAddress;
                                SharedUserList[formattedURL] = currentEmails + "##@@FileName@@##" + sfilename2; // Updating the Ditionary 
                            }
                        }
                    }
                    // for directLink Users
                    else if (currLink.UniquelyPermissionedUsers.Count > 0)
                    {
                        foreach (var uniqueuser in currLink.UniquelyPermissionedUsers)
                        {
                            var userEmail = uniqueuser.Email;
                            if (uniqueuser.Email.ToLower() == currUserEmailAddress.ToLower())
                            {
                                bool isUserKnown = currLink.UniquelyPermissionedUsers[0].IsUserKnown;
                                var allowedUsers = currLink.UniquelyPermissionedUsers[0].AllowedRoles;

                                if (!DirectUsersList.ContainsKey(sAttachmentLinks))
                                {
                                    DirectUsersList.Add(sAttachmentLinks, currUserEmailAddress);//adding new items to KeyValue pair
                                }
                                else if (DirectUsersList.ContainsKey(sAttachmentLinks))
                                {
                                    var currentEmails = DirectUsersList[sAttachmentLinks];
                                    if (currentEmails.Contains(currUserEmailAddress))
                                    {
                                        // to do something if the email address is already added  in the Key of thisdocument
                                    }
                                    else
                                    {
                                        currentEmails += ";" + currUserEmailAddress;
                                        DirectUsersList[sAttachmentLinks] = currentEmails; // Updating the Ditionary 
                                    }
                                }
                            }
                        }

                    }
                    else if (currLink.UsersAddedToGroup != null)
                    {
                        // to check if the user has permisison via group

                    }
                }
                else if (currLink.StatusCode.ToString() == "AccessDenied")
                {
                    Console.WriteLine("Error in ShareResult -  AccessDenied: Parameters are not correct");
                }
                else
                {
                    //to do something
                }
                return "success";
            }
            catch (Exception ex)
            {
                Console.WriteLine("Error in Item Update to Dictonary");
                return ex.Message;
            }
        }

        public static Uri GetSiteUriFromListName(string itemListName)
        {
            try
            {
                Uri currSiteURL;
                if (itemListName.ToLower() == Constants.Group1ListName.ToLower()) { currSiteURL = new Uri(Constants.Group1Site.ToString()); }
                else if (itemListName.ToLower() == Constants.Group2ListName.ToLower()) { currSiteURL = new Uri(Constants.Group2Site.ToString()); }
                else if (itemListName.ToLower() == Constants.Group3ListName.ToLower()) { currSiteURL = new Uri(Constants.Group3Site.ToString()); }
                else if (itemListName.ToLower().Contains(Constants.Group4ListName.ToLower())) { currSiteURL = new Uri(Constants.Group4Site.ToString()); }
                else
                {
                    currSiteURL = new Uri(Constants.siteURL.ToString());
                }
                return currSiteURL;
            }
            catch (Exception ex)
            {
                Console.WriteLine("Error in Uri retrievel" + ex.Message);
                Helper.WriteLog($"{itemListName}" + "," + $"{""}" + "," + $"{""}" + "," + $"{""}" + "," + $"{""}" + "," + $"{""}" + "," + $"{"GetSiteUriFromListName"}" + "," + $"{"Error"}" + "," + $"{ex.Message}");
                return null;
            }
        }

        public static SharingResult GetCurrentEmailSharingResult(string emailitemId, string currEmail, string sAttachmentLinks, string itemListName)
        {
            try
            {
                Uri currSiteURL = GetSiteUriFromListName(itemListName);

                string realm = TokenHelper.GetRealmFromTargetUrl(currSiteURL);
                var token = TokenHelper.GetAppOnlyAccessToken(
                                    TokenHelper.SharePointPrincipal,
                                    currSiteURL.Authority, realm, itemListName).AccessToken;
                using (var ctx = TokenHelper.GetClientContextWithAccessToken(currSiteURL.ToString(), token))
                {
                    SharingResult result = ctx.Web.ShareDocument(sAttachmentLinks, currEmail, ExternalSharingDocumentOption.View, false);
                    ctx.Load(result);
                    ctx.ExecuteQuery();
                    return result;
                }
            }
            catch (Exception ex)
            {
                Console.WriteLine("Error in GetCurrentEmailSharingResult" + currEmail + " : " + ex.Message);
                Helper.WriteLog($"{itemListName}" + "," + $"{emailitemId}" + "," + $"{""}" + "," + $"{""}" + "," + $"{""}" + "," + $"{""}" + "," + $"{"SharingResult"}" + "," + $"{"Error"}" + "," + $"{ex.Message}");
                return null;
            }
        }

        //funciton to send the  email
        public static string TriggerEmail(string emailToAddress, string emailCCAddress, string emailBCCAddress, string emailSubject, string emailContent, string itemRequestType, out string sExceptionMessage, string sAttachmentLinks, string sRequestorName, string sRequestorEmail, ClientContext ctx)
        {
            string notify = "";
            sExceptionMessage = "";
            try
            {

                string smtpServerName = Constants.SMTPserver;
                MailAddress fromMail = new MailAddress(Constants.FromAddress); //Mail Form Address

                MailMessage msg = new MailMessage(); // Build mail Body message
                msg.From = fromMail;

                if (!string.IsNullOrEmpty(emailToAddress))
                {
                    var ToAddress = emailToAddress.TrimEnd(';');
                    var toNotify = ToAddress.Split(new string[] { ";" }, StringSplitOptions.None);
                    if (toNotify.Length > 0)
                    {

                        if (!string.IsNullOrEmpty(emailCCAddress))
                        {
                            var CCAddress = emailCCAddress.TrimEnd(';');
                            var ccAddress = CCAddress.Split(new string[] { ";" }, StringSplitOptions.None);
                            if (ccAddress.Length > 0)
                            {
                                for (var c = 0; c < ccAddress.Length; c++)
                                {
                                    if (ccAddress[c].TrimStart().TrimEnd() != "" && ccAddress[c].TrimStart().TrimEnd() != "undefined")
                                    {
                                        if (IsValidEmail(ccAddress[c].TrimStart().TrimEnd()))
                                        {
                                            msg.CC.Add(ccAddress[c].TrimStart().TrimEnd());
                                        }
                                        else
                                        {
                                            notify += "IsNotEmailFomrat:" + toNotify[c].TrimStart().TrimEnd() + ",";
                                        }
                                    }
                                }
                            }
                        }


                        if (!string.IsNullOrEmpty(emailBCCAddress))
                        {
                            var BCCAddress = emailBCCAddress.TrimEnd(';');
                            var bccAddress = BCCAddress.Split(new string[] { ";" }, StringSplitOptions.None);
                            if (bccAddress.Length > 0)
                            {
                                for (var b = 0; b < bccAddress.Length; b++)
                                {
                                    if (bccAddress[b].TrimStart().TrimEnd() != "" && bccAddress[b].TrimStart().TrimEnd() != "undefined")
                                    {
                                        if (IsValidEmail(bccAddress[b].TrimStart().TrimEnd()))
                                        {
                                            msg.Bcc.Add(bccAddress[b].TrimStart().TrimEnd());
                                        }
                                        else
                                        {
                                            notify += "IsNotEmailFomrat:" + toNotify[b] + ",";
                                        }
                                    }
                                }
                            }
                        }

                        for (var t = 0; t < toNotify.Length; t++)
                        {
                            if (toNotify[t].TrimStart().TrimEnd() != "" && toNotify[t].TrimStart().TrimEnd() != "undefined")
                            {
                                if (IsValidEmail(toNotify[t].TrimStart().TrimEnd()))
                                {
                                    msg.To.Add(toNotify[t].TrimStart().TrimEnd());
                                }
                                else
                                {
                                    notify += "IsNotEmailFomrat:" + toNotify[t].TrimStart().TrimEnd() + ",";
                                }
                            }
                        }

                        msg.Subject = emailSubject;
                        msg.IsBodyHtml = true;

                        StringBuilder buidMessage = new StringBuilder();

                        if (itemRequestType.ToLower() == Constants.SuppliermailRequestType.ToLower())
                        {
                            //This content is to send mail to supplier
                            buidMessage.AppendLine(emailContent + "<br><br>");
                            buidMessage.AppendLine("Attached files / 添付ファイル <br>" + sAttachmentLinks);

                        }
                        else
                        {
                            //This content is to send notfication Mail
                            buidMessage.AppendLine(emailContent);
                        }

                        //EmailContent Control Change 3-Jan-2022
                        // buidMessage.AppendLine(emailContent.Replace("&lt;", "<").Replace("&gt;", ">").Replace("&quot;", @""""));

                        if (itemRequestType.ToLower() == Constants.SuppliermailRequestType.ToLower())
                        {
                            var supportTeam = "<u>" + Constants.SupportTeamMailID.ToString() + "</u>";
                            buidMessage.AppendLine("<br><br>");
                            buidMessage.AppendLine("<b>Note:</b>" + "" + "<br><br>");

                            buidMessage.AppendLine("<a href='#dvenglishversion' >English Version is below.</a><br><br>");

                            buidMessage.AppendLine("<div  id='dvjapaneseversion'>");
                            buidMessage.AppendLine("<u>Microsoft O365アカウントをお持ちでない場合は</u>, Microsoft から'<b>SharePoint Online</b>' O365 パスコードが発行されます。" + "<br>");
                            buidMessage.AppendLine("上記メールがSPAM/Junk Folderなどに入っている可能性もあるので、ご確認ください。" + "<br>");
                            buidMessage.AppendLine("なお、対象データは2週間後に削除される予定です。" + "<br> ");
                            buidMessage.AppendLine("もし何かシステム上の不具合等ございましたら、Support ");
                            buidMessage.AppendLine("Team( " + supportTeam + " )までご連絡ください。" + "" + "<br/><br/> ");
                            buidMessage.AppendLine("このメールは、ご返信いただいても回答することができません。<br><br>");
                            buidMessage.AppendLine("</div>");

                           //buidMessage.AppendLine("<a  id='dvenglishversion'  name = 'dvenglishversion'>English</a><br>");
                            buidMessage.AppendLine("<u>If you do not have a Microsoft O365 account</u>, you will receive a '<b>SharePoint Online</b>' O365 Passcode from Microsoft." + "<br>");
                            buidMessage.AppendLine("Please check your <b>SPAM folder</b> as the above e-mail could be in there. The corresponding files will be deleted automatically after 2 Weeks." + "<br>");
                            buidMessage.AppendLine("If you have any issues please contact the Support Team <a  id='dvenglishversion'  name = 'dvenglishversion'" + supportTeam + "</a><br/><br/>");

                            buidMessage.AppendLine("We cannot reply from this e-mail address.<br>");
                           // buidMessage.AppendLine("</div>");
                        }
                        msg.Body = buidMessage.ToString();
                        //to send Mail
                        SmtpClient client = new SmtpClient();
                        client.UseDefaultCredentials = false;
                        client.Host = smtpServerName;
                        client.DeliveryMethod = SmtpDeliveryMethod.Network;
                        try
                        {
                            Console.WriteLine("Sending Email...");
                            client.Send(msg);
                            //Helper.WriteLog_EmailContent("MailSentTo: \n" + emailToAddress);
                            //Helper.WriteLog_EmailContent("MailContent: \n" + buidMessage+ "\n----------------------------------------------------------------------");                        
                            Console.WriteLine("Mail Sent...");
                            notify += "MailSentTo:" + emailToAddress.TrimEnd(';');
                            notify += "Success;AllMailsSent";
                        }
                        catch (SmtpFailedRecipientsException ex)
                        {
                            StringBuilder sbFailedRecipient = new StringBuilder();
                            for (int i = 0; i < ex.InnerExceptions.Length; i++)
                            {
                                SmtpStatusCode status = ex.InnerExceptions[i].StatusCode;
                                if (status == SmtpStatusCode.MailboxBusy ||
                                    status == SmtpStatusCode.MailboxUnavailable)
                                {
                                    Console.WriteLine("Delivery failed - retrying in 5 seconds.");
                                }
                                else
                                {
                                    notify += "FailedtoDeliverMail:" + ex.InnerExceptions[i].FailedRecipient;
                                    Console.WriteLine("Failed to deliver message to {0}",
                                        ex.InnerExceptions[i].FailedRecipient);
                                    sbFailedRecipient.AppendLine(ex.InnerExceptions[i].FailedRecipient + ";");
                                }

                                notify += "FailedtoDeliverMail:" + ex.InnerExceptions[i].FailedRecipient + ";";
                                Console.WriteLine("Failed to deliver message to {0}",
                                       ex.InnerExceptions[i].FailedRecipient);
                                sbFailedRecipient.AppendLine(ex.InnerExceptions[i].FailedRecipient);
                            }

                            ex.StatusCode = SmtpStatusCode.MailboxUnavailable;
                            sExceptionMessage = SmtpStatusCode.MailboxUnavailable.ToString() + sbFailedRecipient.ToString();
                            Console.WriteLine("Mail Error..." + sExceptionMessage);
                            notify += "MailFailure";

                        }

                    }
                }
                else
                {
                    notify += "ToAddressIsEmpty";
                    sExceptionMessage = "ToAddressIsEmpty";
                }
                return notify;
            }
            catch (Exception ex)
            {
                notify += ex.Message;
                Console.WriteLine(ex.Message + ": " + ex.StackTrace);
                sExceptionMessage = ex.Message;
                return notify;
            }
        }

        public static string UpdateListItemMailMissendRequest(string listName, string itemID, string MailSentStatus, string accessToken)
        {
            string sResult = "";
            string APIURL = String.Concat(Constants.siteURL.ToString().TrimEnd('/'), "/_api/web/lists/getbytitle('", listName, "')/items(" + itemID + ")");
            try
            {
                StringBuilder uPdateData = new StringBuilder();
                uPdateData.Append(",'" + "MailSentStatus" + "' : '" + JsonConvert.SerializeObject(MailSentStatus).ToString().Trim('"') + "'");

                uPdateData.Append("}");

                string stringData = String.Concat("{'__metadata': { 'type': 'SP.Data.", listName, "ListItem", "'}");
                HttpWebRequest endpointRequest = (HttpWebRequest)HttpWebRequest.Create(APIURL);
                endpointRequest.Headers.Add("Authorization", string.Concat("Bearer ", accessToken));
                endpointRequest.Method = "POST";
                endpointRequest.Accept = "application/json; odata=verbose";
                endpointRequest.ContentType = "application/json; odata=verbose";
                endpointRequest.Headers.Add("X-HTTP-Method", "MERGE");
                endpointRequest.Headers.Add("IF-MATCH", "*");
                stringData = String.Concat(stringData, uPdateData.ToString());
                endpointRequest.ContentLength = Encoding.UTF8.GetByteCount(stringData);
                StreamWriter writer = new StreamWriter(endpointRequest.GetRequestStream());
                writer.Write(stringData);
                writer.Flush();
                using (WebResponse wresp = endpointRequest.GetResponse())
                {
                    sResult = "Success";
                }
                return sResult;
            }
            catch (Exception ee)
            {
                sResult = ee.Message;
                return sResult;
            }

        }

        public static bool IsValidEmail(string email)
        {
            try
            {
                var addr = new System.Net.Mail.MailAddress(email);
                return true;
            }
            catch
            {
                return false;
            }
        }

    }
}