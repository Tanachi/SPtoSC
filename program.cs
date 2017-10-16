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
            string[] attr = { "Project Lead|Text", "Project Team|Text", "External ID|Numeric","Priority|Numberic",
            "Value Stream/LOB|List","RAG Status|List","Percent Complete|Numeric", "Due Date|Date","New Requested Budget|Numeric",
            "Appropriated Budget|Numeric","Project Business Value|Text","Project Dependencies/Assumptions/Risks|Text",
            "Status Comments|Text","Total Spent to Date|Numberic","Financial Comments|Text"};
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
            addAttribute(SSB, attr);
            addAttribute(TSB, attr);
            addAttribute(TSSB, attr);
            addAttribute(PSB, attr);
            addAttribute(CSB, attr);
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
                            addItem(TSB, item, attr);
                        break;
                    case "TSSB":
                        if(TSSB.Item_FindByName(item["Title"].ToString()) == null)
                            addItem(TSSB, item, attr);
                        break;
                    case "CSB":
                        if(CSB.Item_FindByName(item["Title"].ToString()) == null)
                            addItem(CSB, item, attr);
                        break;
                    case "PSB":
                        if(PSB.Item_FindByName(item["Title"].ToString()) == null)
                            addItem(PSB, item, attr);
                        break;
                    case "SSB":
                        if(SSB.Item_FindByName(item["Title"].ToString()) == null)
                            addItem(SSB, item, attr);
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
        static void addAttribute(Story story, string[] attr)
        {
            foreach (var att in attr)
            {
                string[] split = att.Split('|');
                var name = split[0];
                var type = split[1];
                if (story.Attribute_FindByName(name) == null)
                {
                    if (type == "Text")
                    {
                        story.Attribute_Add(name, SC.API.ComInterop.Models.Attribute.AttributeType.Text);
                    }
                    else if (type == "Numeric")
                    {
                        story.Attribute_Add(name, SC.API.ComInterop.Models.Attribute.AttributeType.Numeric);
                    }
                    else if (type == "List")
                    {
                        story.Attribute_Add(name, SC.API.ComInterop.Models.Attribute.AttributeType.List);
                    }
                    else if (type == "Date")
                    {
                        story.Attribute_Add(name, SC.API.ComInterop.Models.Attribute.AttributeType.Date);
                    }

                }
            }
        }
        // Adds Attribute data to item
        static void addItem(Story story, ListItem item, string[] attr)
        { 
            Item storyItem = story.Item_AddNew(item["Title"].ToString());
            string[] shareAtt = { "Project_x0020_Lead", "Project_x0020_Team", "External_x0020_ID", "Priority","Category","Status",
                "Percent_x0020_Complete", "Due_x0020_Date","New_x0020_Requested_x0020_Budget","Appropriated_x0020_Budget",
                "Project_x0020_Business_x0020_Val","Project_x0020_Dependencies_x002f","Status_x0020_Comments","Total_x0020_Spent_x0020_to_x0020",
            "Financial_x0020_Comments" };

            for(var i = 0; i < attr.Length - 1; i++)
            {
                string[] split = attr[i].Split('|');
                var name = split[0];
                var type = split[1];
                if (item[shareAtt[i]] != null)
                {
                    if (type == "Text" || type == "List")
                    {
                        storyItem.SetAttributeValue(story.Attribute_FindByName(name), item[shareAtt[i]].ToString());
                    }
                    else if (type == "Numeric")
                    {
                        storyItem.SetAttributeValue(story.Attribute_FindByName(name), double.Parse(item[shareAtt[i]].ToString()));
                    }

                    else if (type == "Date")
                    {
                        storyItem.SetAttributeValue(story.Attribute_FindByName(name), DateTime.Parse(item[shareAtt[i]].ToString()));
                    }
                }
            }
            if(item["Notes"] != null)
                storyItem.Description = item["Notes"].ToString();
            if(item["Start_x0020_Date"] !=null)
                storyItem.StartDate = DateTime.Parse(item["Start_x0020_Date"].ToString());
            if(item["Tags_x002e_ETS_x0020_Initiatives"] != null)
            {
                string[] tagSplit = item["Tags_x002e_ETS_x0020_Initiatives"].ToString().Split(',');
                foreach(var tag in tagSplit)
                {
                    if(story.ItemTag_FindByNameAndGroup(tag, "ETS Initiatives") == null)
                    {
                        story.ItemTag_AddNew(tag, "", "ETS Initiatives");
                        storyItem.Tag_AddNew(story.ItemTag_FindByName(tag));
                    }
                }
            }
            if (item["Tags_x002e_Governor_x0020_Priori"] != null)
            {
                string[] tagSplit = item["Tags_x002e_Governor_x0020_Priori"].ToString().Split(',');
                foreach (var tag in tagSplit)
                {
                    if (story.ItemTag_FindByNameAndGroup(tag, "Governor Priorities") == null)
                    {
                        story.ItemTag_AddNew(tag, "", "Governor Priorities");
                        storyItem.Tag_AddNew(story.ItemTag_FindByName(tag));
                    }
                }
            }
            if (item["Tags_x002e_ETS_x0020_Priorities"] != null)
            {
                string[] tagSplit = item["Tags_x002e_ETS_x0020_Priorities"].ToString().Split(',');
                foreach (var tag in tagSplit)
                {
                    if (story.ItemTag_FindByNameAndGroup(tag, "ETS Priorities") == null)
                    {
                        story.ItemTag_AddNew(tag, "", "ETS Priorities");

                        storyItem.Tag_AddNew(story.ItemTag_FindByName(tag));
                    }
                }
            }
        }
    }
}

