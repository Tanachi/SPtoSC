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
            string sharpCloudUsername = ConfigurationManager.AppSettings["sharpCloudUsername"];
            string sharpCloudPassword = ConfigurationManager.AppSettings["sharpCloudPassword"];
            string sharePointUsername = ConfigurationManager.AppSettings["sharePointUsername"];
            string sharePointPassword = ConfigurationManager.AppSettings["sharePointPassword"];
            string sharePointSite = ConfigurationManager.AppSettings["sharePointSite"];
            string sharePointList = ConfigurationManager.AppSettings["sharePointList"];
            string storyTSB = ConfigurationManager.AppSettings["storyTSB"];
            string storyTSSB = ConfigurationManager.AppSettings["storyTSSB"];
            string storyPSB = ConfigurationManager.AppSettings["storyPSB"];
            string storyCSB = ConfigurationManager.AppSettings["storyCSB"];
            string storySSB = ConfigurationManager.AppSettings["storySSB"];

            // Login into sharpcloud
            var sc = new SharpCloudApi(sharpCloudUsername, sharpCloudPassword);
            Story TSB = sc.LoadStory(storyTSB);
            Story TSSB = sc.LoadStory(storyTSSB);
            Story PSB = sc.LoadStory(storyPSB);
            Story CSB = sc.LoadStory(storyCSB);
            Story SSB = sc.LoadStory(storySSB);
            // Login into Sharepoint
            var securePassword = new SecureString();
            foreach (var character in sharePointPassword)
            {
                securePassword.AppendChar(character);
            }
            securePassword.MakeReadOnly();
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
            // Adds Attribute to story if none exist
            if(SSB.Attribute_FindByName("Priority") == null)
            {
                addAttribute(SSB);
                addAttribute(TSB);
                addAttribute(TSSB);
                addAttribute(PSB);
                addAttribute(CSB);
            }
            // Goes through List Items
            foreach(var item in items)
            {
            // Gets branch and inserts into the branch story
                var fields = item.FieldValuesAsText;
                string branch = "";
                if(item["Branch"] != null)
                    branch = item["Branch"].ToString();
                switch (branch)
                {
                    case "TSB":
                        if(TSB.Item_FindByName(item["Title"].ToString()) == null)
                            addItem(TSB, item);
                        break;
                    case "TSSB":
                        if(TSSB.Item_FindByName(item["Title"].ToString()) == null)
                            addItem(TSSB, item);
                        break;
                    case "CSB":
                        if(CSB.Item_FindByName(item["Title"].ToString()) == null)
                            addItem(CSB, item);
                        break;
                    case "PSB":
                        if(PSB.Item_FindByName(item["Title"].ToString()) == null)
                            addItem(PSB, item);
                        break;
                    case "SSB":
                        if(SSB.Item_FindByName(item["Title"].ToString()) == null)
                            addItem(SSB, item);
                        break;
                }
            }
            TSB.Save();
            TSSB.Save();
            PSB.Save();
            CSB.Save();
            SSB.Save();
        }
        // Adds Attribute to story
        static void addAttribute(Story story)
        {
            story.Attribute_Add("Percent Complete", SC.API.ComInterop.Models.Attribute.AttributeType.Numeric);
            story.Attribute_Add("Project Business Value", SC.API.ComInterop.Models.Attribute.AttributeType.Text);
            story.Attribute_Add("Project Lead", SC.API.ComInterop.Models.Attribute.AttributeType.Text);
            story.Attribute_Add("Project Team", SC.API.ComInterop.Models.Attribute.AttributeType.Text);
            story.Attribute_Add("RAG Status", SC.API.ComInterop.Models.Attribute.AttributeType.List);
            story.Attribute_Add("Status Comments", SC.API.ComInterop.Models.Attribute.AttributeType.Text);
            story.Attribute_Add("Project Dependencies/Assumptions/Risks", SC.API.ComInterop.Models.Attribute.AttributeType.Text);
            story.Attribute_Add("Appropriated Budget", SC.API.ComInterop.Models.Attribute.AttributeType.Numeric);
            story.Attribute_Add("Total Spent to Date", SC.API.ComInterop.Models.Attribute.AttributeType.Numeric);
            story.Attribute_Add("New Requested Budget", SC.API.ComInterop.Models.Attribute.AttributeType.Numeric);
            story.Attribute_Add("Financial Comments", SC.API.ComInterop.Models.Attribute.AttributeType.Text);
            story.Attribute_Add("Value Stream/LOB", SC.API.ComInterop.Models.Attribute.AttributeType.List);
            story.Attribute_Add("Priority", SC.API.ComInterop.Models.Attribute.AttributeType.Numeric);
            story.Attribute_Add("Due Date", SC.API.ComInterop.Models.Attribute.AttributeType.Date);
        }
        // Adds Attribute data to item
        static void addItem(Story story, ListItem item)
        {
            Item storyItem = story.Item_AddNew(item["Title"].ToString());
            if(item["Project_x0020_Team"] != null)
                storyItem.SetAttributeValue(story.Attribute_FindByName("Project Team"), item["Project_x0020_Team"].ToString());
            if(item["Project_x0020_Lead"] != null)
                storyItem.SetAttributeValue(story.Attribute_FindByName("Project Lead"), item["Project_x0020_Lead"].ToString());
            if(item["Percent_x0020_Complete"] != null)
                storyItem.SetAttributeValue(story.Attribute_FindByName("Percent Complete"), double.Parse(item["Percent_x0020_Complete"].ToString()));
            if(item["Appropriated_x0020_Budget"] != null)
                storyItem.SetAttributeValue(story.Attribute_FindByName("Appropriated Budget"), double.Parse(item["Appropriated_x0020_Budget"].ToString()));
            if(item["New_x0020_Requested_x0020_Budget"] != null)
                storyItem.SetAttributeValue(story.Attribute_FindByName("New Requested Budget"), double.Parse(item["New_x0020_Requested_x0020_Budget"].ToString()));
            if(item["Priority"] != null)
                storyItem.SetAttributeValue(story.Attribute_FindByName("Priority"), double.Parse(item["Priority"].ToString()));
            if(item["Total_x0020_Spent_x0020_to_x0020"] != null)
                storyItem.SetAttributeValue(story.Attribute_FindByName("Total Spent to Date"), double.Parse(item["Total_x0020_Spent_x0020_to_x0020"].ToString()));
            if(item["Project_x0020_Business_x0020_Val"] != null)
                storyItem.SetAttributeValue(story.Attribute_FindByName("Project Business Value"), item["Project_x0020_Business_x0020_Val"].ToString());
            storyItem.SetAttributeValue(story.Attribute_FindByName("RAG Status"), item["Status"].ToString());
            if(item["Status_x0020_Comments"] != null)
                storyItem.SetAttributeValue(story.Attribute_FindByName("Status Comments"), item["Status_x0020_Comments"].ToString());
            if(item["Financial_x0020_Comments"] != null)
                storyItem.SetAttributeValue(story.Attribute_FindByName("Financial Comments"), item["Financial_x0020_Comments"].ToString());
            if(item["Project_x0020_Dependencies_x002f"] !=null)
                storyItem.SetAttributeValue(story.Attribute_FindByName("Project Dependencies/Assumptions/Risks"), item["Project_x0020_Dependencies_x002f"].ToString());
            if(item["Category"] != null)
                storyItem.SetAttributeValue(story.Attribute_FindByName("Value Stream/LOB"), item["Category"].ToString());
            if(item["Notes"] != null)
                storyItem.Description = item["Notes"].ToString();
            if(item["Start_x0020_Date"] !=null)
                storyItem.StartDate = DateTime.Parse(item["Start_x0020_Date"].ToString());
            if(item["Due_x0020_Date"] != null)
                storyItem.SetAttributeValue(story.Attribute_FindByName("Due Date"), (item["Due_x0020_Date"].ToString()));
        }
    }
}

