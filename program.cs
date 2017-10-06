using System;
using System.Configuration;
using System.IO;
using System.Net;
using System.Text;
using System.Security;
using Microsoft.SharePoint.Client;
using SC.API.ComInterop;
using SC.API.ComInterop.Models;
namespace SharePointExporter
{
    class Program
    {
        
        static void Main(string[] args)
        {
            // Load from App.Config
            string sharpCloudStoryID = ConfigurationManager.AppSettings["sharpCloudStoryID"];
            string sharpCloudUsername = ConfigurationManager.AppSettings["sharpCloudUsername"];
            string sharpCloudPassword = ConfigurationManager.AppSettings["sharpCloudPassword"];
            string sharePointUsername = ConfigurationManager.AppSettings["sharePointUsername"];
            string sharePointPassword = ConfigurationManager.AppSettings["sharePointPassword"];
            string sharePointSite = ConfigurationManager.AppSettings["sharePointSite"];
            string sharePointList = ConfigurationManager.AppSettings["sharePointList"];
            // Login into sharpcloud
            var sc = new SharpCloudApi(sharpCloudUsername, sharpCloudPassword);
            var story = sc.LoadStory(sharpCloudStoryID);
            var securePassword = new SecureString();
            foreach (var character in sharePointPassword)
            {
                securePassword.AppendChar(character);
            }
            securePassword.MakeReadOnly();
            // Login into Sharepoint
            var sharePointCredentials = new SharePointOnlineCredentials(sharePointUsername, securePassword);
            ClientContext context = new ClientContext(sharePointSite);
            context.Credentials = sharePointCredentials;
            List list = context.Web.Lists.GetByTitle(sharePointList);
            // Setup query
            CamlQuery query = new CamlQuery();
            query.ViewXml = "<View/>";
            ListItemCollection items = list.GetItems(query);
            // Loads list
            context.Load(list);
            context.Load(items);
            context.ExecuteQuery();
            // Goes through List
            foreach(var item in items)
            {
                Item sharpItem = story.Item_AddNew(item["Title"]);
            }

            story.Save();
        }
    }
}
