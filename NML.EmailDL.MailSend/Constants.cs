using System;
using System.Configuration;
using System.IO;
using System.Net;
using System.Net.Mail;
using System.ComponentModel;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Microsoft.SharePoint.Client;
using OfficeDevPnP.Core;
using System.Globalization;

namespace NML.EmailDL.MailSend
{
    class Constants
    {

        public static string LogFilePath = ConfigurationManager.AppSettings["LogFilePath"].ToString();

        // SMTP configuration variables
        public static string SMTPserver = ConfigurationManager.AppSettings["SMTPserver"].ToString();

        public static string FromAddress = ConfigurationManager.AppSettings["FromAddress"].ToString();

        // Site & List Details
        public static Uri siteURL = new Uri(ConfigurationManager.AppSettings["DestinationSiteURL"].ToString());

        public static string MailContentListDisplayName = ConfigurationManager.AppSettings["MailContentListDisplayName"].ToString();

        public static string FilterItemStatus = ConfigurationManager.AppSettings["FilterItemStatus"].ToString();

        public static string ItemStartId = ConfigurationManager.AppSettings["ItemStartId"].ToString();

        public static string CompletedItemStatus = ConfigurationManager.AppSettings["CompletedItemStatus"].ToString();

        public static string MailFailureItemStatus = ConfigurationManager.AppSettings["MailFailureItemStatus"].ToString();

        public static string SuppliermailRequestType = ConfigurationManager.AppSettings["SupplierMailRequestType"].ToString();

        public static string ApproverMailRequestType = ConfigurationManager.AppSettings["ApproverMailRequestType"].ToString();

        public static string ArchiveListName = ConfigurationManager.AppSettings["ArchiveListName"].ToString();

        public static string Group1Site = ConfigurationManager.AppSettings["Group1Site"].ToString();
        public static string Group2Site = ConfigurationManager.AppSettings["Group2Site"].ToString();
        public static string Group3Site = ConfigurationManager.AppSettings["Group3Site"].ToString();
        public static string Group4Site = ConfigurationManager.AppSettings["Group4Site"].ToString();

        public static string Group1ListName = ConfigurationManager.AppSettings["Group1ListName"].ToString();
        public static string Group2ListName = ConfigurationManager.AppSettings["Group2ListName"].ToString();
        public static string Group3ListName = ConfigurationManager.AppSettings["Group3ListName"].ToString();
        public static string Group4ListName = ConfigurationManager.AppSettings["Group4ListName"].ToString();

        public static string Group1Teams = ConfigurationManager.AppSettings["Group1Teams"].ToString();
        public static string Group2Teams = ConfigurationManager.AppSettings["Group2Teams"].ToString();
        public static string Group3Teams = ConfigurationManager.AppSettings["Group3Teams"].ToString();
        public static string Group4Teams = ConfigurationManager.AppSettings["Group4Teams"].ToString();



        public static string FileDeleteDefaultPeriod = ConfigurationManager.AppSettings["FileDeleteDefaultPeriod"].ToString();

        public static string SupportTeamMailID = ConfigurationManager.AppSettings["SupportTeamMailID"].ToString();
        public static string MailSentByAMO = "Mail Sent by Batch";

        public static string DirectLinkUsers = "";
        public static string SharedLinkUsers = "";
        public static string[] SahredURL;

 

    }
}
