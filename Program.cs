using System;
using System.Collections.Generic;
using System.Linq;
using System.Security;
using System.Text;
using System.Threading;
using System.Threading.Tasks;
using Microsoft.SharePoint.Client;
using OfficeDevPnP.Core.Framework.Provisioning.Connectors;
using OfficeDevPnP.Core.Framework.Provisioning.Model;
using OfficeDevPnP.Core.Framework.Provisioning.ObjectHandlers;
using OfficeDevPnP.Core.Framework.Provisioning.Providers.Xml;
using OfficeDevPnP.Core.Framework.Provisioning.Extensibility;
using OfficeDevPnP.Core.Diagnostics;
using System.Text.Json;
using System.Security.AccessControl;
using System.Xml;
using System.IO;
using System.Configuration;
using Microsoft.ProjectServer.Client;
using OfficeDevPnP.Core;
using System.Runtime.InteropServices.ComTypes;
using static System.Net.Mime.MediaTypeNames;
using Newtonsoft.Json.Linq;
using SPList = Microsoft.SharePoint.Client.List;

using Newtonsoft.Json;
using OfficeOpenXml;

namespace SPSiteProvosioning
{
    //main root site template - PnPProvisioning_main_root.xml
    //Main site template - PnPProvisioning_SM.xml
    //fix lookup template - PnPProvisioning_Fix_Lookup.xml
    //dynamic list template - PnPProvisioning_DynamicList.xml
    class Program
    {
        static string username;// = "developer1@hochhuth-consulting.de";
        static string password;// = "Fub62326";
        static string TimesheetListID;
        static string InputlstNames;
        static string timesheetYes;


        static void Main(string[] args)
        {

            //TestSample();
            //provision the root site or site collection
            //ProvisionRootSite("https://smalsusinfolabs.sharepoint.com/sites/HHHHDesign");

            //provision subsites
            siteProvisioing(); // for first time site creation.
            //GetProvisioningTemplate("https://smalsusinfolabs.sharepoint.com/sites/HHHHQA/SP");

            ConsoleKeyInfo ch;
            Console.WriteLine("Press the Escape (Esc) key to quit: \n");
            ch = Console.ReadKey();
            if (ch.Key == ConsoleKey.Escape)
            {
                Environment.Exit(0);
            }
        }

        private static void siteProvisioing()
        {
            ConsoleColor defaultForeground = Console.ForegroundColor;

            // Collect information

            //string siteURL = GetInput("Enter site location ", false, defaultForeground);
            //string siteName = GetInput("Enter Subsite Name ", false, defaultForeground);
            Console.WriteLine("Enter site location :");
            string siteURL = Console.ReadLine();
            Console.WriteLine("Enter Subsite Name: ");
            string siteName = Console.ReadLine();
            Console.WriteLine("If Timesheet Functionality required, enter 'Yes' otherwise enter 'No': ");
            timesheetYes = Console.ReadLine();

            //Console.WriteLine("Enter Task Lists Name with comma seperated (List1,List2) :");
            //InputlstNames = Console.ReadLine();
            //username = GetInput("Enter your user name ", false, defaultForeground);
            //password = GetInput("Enter your password ", true, defaultForeground);
            //SecureString pwd = new SecureString();
            //foreach (char c in password.ToCharArray()) pwd.AppendChar(c);
            //AddDynamicTaskLists(siteURL);
            ProvisionSubSite(siteName, siteURL);

            ConsoleKeyInfo ch;
            Console.WriteLine("Press the Escape (Esc) key to quit: \n");
            ch = Console.ReadKey();
            if (ch.Key == ConsoleKey.Escape)
            {
                Environment.Exit(0);
            }
        }

        static void ProvisionSubSite(string siteName, string siteURL)
        {
            Console.WriteLine("Start Provisioing : {0:hh.mm.ss}", DateTime.Now);
            try
            {
                var authManager = new OfficeDevPnP.Core.AuthenticationManager();
                // This method calls a pop up window with the login page and it also prompts  
                // for the multi factor authentication code.  
                ClientContext ctx = authManager.GetWebLoginClientContext(siteURL);

                // This method calls a pop up window with the login page and it also prompts  
                // for the multi factor authentication code.  
                string webUrl = siteURL + '/' + siteName;
                //using (ClientContext ctx = authManager.GetWebLoginClientContext(siteURL))
                //{

                //using (ClientContext ctx = new ClientContext(siteURL))
                //{
                //string username = "developer1@hochhuth-consulting.de";
                //string password = "Fub62326";
                //SecureString securePassword = GetSecureString(password);
                //ctx.Credentials = new SharePointOnlineCredentials(username, securePassword);


                //ctx.Dispose();

                Web Mainweb = ctx.Web;
                ctx.Load(Mainweb, w => w.Title);
                ctx.ExecuteQuery();
                Console.WriteLine("You have connected to {0} site", Mainweb.Title);

                string[] splitted = siteURL.ToLower().Split(new[] { "sites/" }, StringSplitOptions.None)
                           .Select(value => value.Trim())
                           .ToArray();

                bool isSubWebOfSubWeb = splitted.Length > 1 && splitted[1].Split('/').Length > 1 ? true : false;

                bool IsExist = true;

                // If the site is subsite of a subsite, then use OpenWeb as show below in comment
                if (isSubWebOfSubWeb)
                {

                    //IsExist = ctx.Site.OpenWeb(siteURL).WebExists(siteName);
                    var web = ctx.Web;
                    ctx.Load(web, w => w.Webs);
                    ctx.ExecuteQuery(); // use a simple linq query to get any sub webs with the URL we want to check 
                    var subWeb = (from w in web.Webs where w.Url == webUrl select w).SingleOrDefault();
                    if (subWeb != null)
                    {
                        // if found true 
                        IsExist = true;
                    }
                    else
                    {
                        IsExist = false;
                    }

                }
                else
                {
                    IsExist = ctx.Site.RootWeb.WebExists(siteName);
                }

                // Check If Subsite existing

                if (IsExist)
                {
                    Console.WriteLine("Site is Already Exists");
                }
                else
                {

                    Console.WriteLine("Site is not Exists");
                    Console.WriteLine("Creating new site.....");
                    WebCreationInformation webCreationInformation = new WebCreationInformation()
                    {
                        WebTemplate = "STS#3",
                        Title = siteName,
                        Url = siteName,
                        Language = 1033,
                        UseSamePermissionsAsParentSite = true
                    };
                    Web newWeb = ctx.Web.Webs.Add(webCreationInformation);
                    ctx.Load(newWeb);
                    ctx.ExecuteQuery();

                    //Apply provision on SubSite                    
                    Console.WriteLine("target URL - " + webUrl);

                    if (timesheetYes.ToLower() == "yes")
                    {
                        ApplyTemplateOnWeb(webUrl, "PnPProvisioning_SM.xml");

                        //Fix lookup Columns
                        Console.WriteLine("Validating the lookup in Lists...");
                        ApplyTemplateOnWeb(webUrl, "PnPProvisioning_Fix_Lookup.xml");
                    }
                    else
                    {
                        ApplyTemplateOnWeb(webUrl, "PnPProvisioning_wt_Timesheet.xml");

                        //Fix lookup Columns
                        Console.WriteLine("Validating the lookup in Lists...");
                       // ApplyTemplateOnWeb(webUrl, "PnPProvisioning_Fix_Lookup_wt_Timesheet.xml");
                    }


                    //Console.WriteLine("Updateing Parent Id in SmartMetadata List");
                    //SetParentforSmartMetaData(webUrl, "SmartMetadata", "SmartMetaData.json");

                    //Console.WriteLine("Updateing Parent Id in Top Navigation List");
                    //SetParentforSmartMetaData(webUrl, "TopNavigation", "TopNavigation.json");
                    Console.WriteLine("--------------------------------------------------------");
                    /*
                    string[] tLists = InputlstNames.Trim().Split(',');
                    foreach (string lstName in tLists)
                    {
                        if (lstName.Trim() != string.Empty)
                        {
                            Console.WriteLine("List Provisioining : " + lstName.Trim());
                            MakeDynamicListTemplate(webUrl, "PnPProvisioning_DynamicList.xml", lstName.Trim());
                            Console.WriteLine("--------------------------------------------------------");
                        }

                    }
                    */

                    AddDynamicTaskLists(ctx, webUrl);

                    //ApplyTemplate(webUrl, "PnPProvisioningSP_Groups");
                    //}

                }
            }
            catch (Exception ex)
            {
                Console.WriteLine("Something went wrong. " + ex.Message);
            }

            Console.WriteLine("End Provisioing : {0:hh.mm.ss}", DateTime.Now);
        }

        private static void ApplyTemplateOnWeb(string targetWebUrl, string fileName)
        {
            var authManager = new OfficeDevPnP.Core.AuthenticationManager();
            // This method calls a pop up window with the login page and it also prompts  
            // for the multi factor authentication code.  
            ClientContext ctx = authManager.GetWebLoginClientContext(targetWebUrl);
            // The obtained ClientContext object can be used to connect to the SharePoint site.  
            Web web = ctx.Web;
            ctx.Load(web, w => w.Title);
            ctx.ExecuteQuery();
            Console.WriteLine("You have connected to {0} site!!", web.Title);
            //using (ClientContext context = new ClientContext(targetWebUrl))
            // {              
            /*
            SecureString securePassword = GetSecureString(password);
            context.Credentials = new SharePointOnlineCredentials(username, securePassword);
            context.RequestTimeout = System.Threading.Timeout.Infinite;

            Web web = context.Web;
            context.Load(web, w => w.Title);
            context.ExecuteQueryRetry();
            */
            // Configure the XML file system provider
            XMLTemplateProvider provider = new XMLFileSystemTemplateProvider("..\\..\\pnpTemplate", "");
            ProvisioningTemplate template = provider.GetTemplate(fileName);

            ProvisioningTemplateApplyingInformation ptai
                = new ProvisioningTemplateApplyingInformation();
            ptai.ProgressDelegate = delegate (String message, Int32 progress, Int32 total)
            {
                Console.WriteLine("{0:00}/{1:00} – {2}", progress, total, message);
            };

            // Associate file connector for assets
            FileSystemConnector connector = new FileSystemConnector("..\\..\\pnpTemplate", "");
            template.Connector = connector;
            //template.Security.SiteGroups.Add(
            web.ApplyProvisioningTemplate(template, ptai);
            //}
        }

        public static void AddDynamicTaskLists(ClientContext parentCtx, string webURL)
        {
            using (ClientContext ctx = parentCtx.Clone(new Uri(webURL)))
            {
                Web web = ctx.Web;
                ListCollection lists = web.Lists;
                ctx.Load(lists, allLists => allLists.Include(l => l.Title));
                ctx.ExecuteQuery();

                string firstListName = "SmartMetadata"; // Define your list names
                string secondListName = "Portfolio Types";
                string thirdListName = "Task Types";
                string fourthListName = "TopNavigation";
                string firstExcelFilePath = "C:\\Users\\HARSH\\Documents\\PnPDocuments\\SmartMetaDataSheet.xlsx";
                string secondExcelFilePath = "C:\\Users\\HARSH\\Documents\\PnPDocuments\\Portfolio Types.xlsx";
                string thirdExcelFilePath = "C:\\Users\\HARSH\\Documents\\PnPDocuments\\Task Types.xlsx";
                string fourthExcelPath = "C:\\Users\\HARSH\\Downloads\\TopNavigation_updated.xlsx";



                Console.WriteLine($"Processing lists for site: {webURL}");
               // ProcessList(ctx, lists, firstListName, firstExcelFilePath);
                ProcessList(ctx, lists, secondListName, secondExcelFilePath);
                ProcessList(ctx, lists, thirdListName, thirdExcelFilePath);
                UploadExcelDataToSharePoint(webURL, fourthListName, fourthExcelPath);
                UploadExcelDataToSharePoint(webURL, firstListName, firstExcelFilePath);
                //UploadExcelToSharePoint("C:\\Users\\HARSH\\Downloads\\Site Configurations.csv", "https://hhhhteams.sharepoint.com/sites/HHHH/twentyTest", "Site Configurations");
            }
        }

        //Upload Excel to SharePoint
        public static void UploadExcelDataToSharePoint(string siteUrl, string listName, string excelFilePath)
        {
            ExcelPackage.LicenseContext = LicenseContext.NonCommercial;
            var authManager = new OfficeDevPnP.Core.AuthenticationManager();

            try
            {
                using (ClientContext ctx = authManager.GetWebLoginClientContext(siteUrl))
                {
                    Web web = ctx.Web;
                    ListCollection lists = web.Lists;
                    ctx.Load(lists, allLists => allLists.Include(l => l.Title));
                    ctx.ExecuteQuery();

                    List list = lists.FirstOrDefault(l => l.Title.Equals(listName, StringComparison.OrdinalIgnoreCase));

                    if (list != null)
                    {
                        var titleToIdMap = ProcessNewExcelData(ctx, list, excelFilePath);
                        UpdateParentLookup(ctx, list, excelFilePath, titleToIdMap);
                    }
                    else
                    {
                        Console.WriteLine($"List '{listName}' not found.");
                    }
                }
            }
            catch (ServerException ex)
            {
                Console.ForegroundColor = ConsoleColor.Red;
                Console.WriteLine($"SharePoint Server Error: {ex.Message}");
            }
            catch (Exception ex)
            {
                Console.ForegroundColor = ConsoleColor.Red;
                Console.WriteLine($"Error: {ex.Message}");
            }
        }

        private static Dictionary<string, int> ProcessNewExcelData(ClientContext ctx, List list, string excelFilePath)
        {
            var titleToIdMap = new Dictionary<string, int>();

            try
            {
                FileInfo excelFile = new FileInfo(excelFilePath);
                using (ExcelPackage package = new ExcelPackage(excelFile))
                {
                    ExcelWorksheet worksheet = package.Workbook.Worksheets.FirstOrDefault();

                    if (worksheet != null)
                    {
                        int rowCount = worksheet.Dimension.Rows;
                        int columnCount = worksheet.Dimension.Columns;

                        var excelHeaders = new List<string>();
                        for (int col = 1; col <= columnCount; col++)
                        {
                            string header = worksheet.Cells[1, col].Value?.ToString();
                            if (!string.IsNullOrEmpty(header))
                            {
                                excelHeaders.Add(header.Trim());
                            }
                        }

                        if (!excelHeaders.Contains("Title"))
                        {
                            Console.WriteLine("Excel sheet must contain 'Title' column.");
                            return titleToIdMap;
                        }

                        for (int row = 2; row <= rowCount; row++)
                        {
                            try
                            {
                                ListItemCreationInformation itemCreateInfo = new ListItemCreationInformation();
                                ListItem newItem = list.AddItem(itemCreateInfo);
                                foreach (string header in excelHeaders)
                                {
                                    if (header.Equals("Parent", StringComparison.OrdinalIgnoreCase))
                                        continue;

                                    string excelValue = worksheet.Cells[row, excelHeaders.IndexOf(header) + 1].Value?.ToString();
                                    if (!string.IsNullOrEmpty(excelValue))
                                    {
                                        try
                                        {
                                            newItem[header] = excelValue;
                                        }
                                        catch (Exception innerEx)
                                        {
                                            Console.ForegroundColor = ConsoleColor.Yellow;
                                            Console.WriteLine($"Cannot update field '{header}' with value '{excelValue}' at row {row}: {innerEx.Message}");
                                        }
                                    }
                                }

                                newItem.Update();
                                ctx.ExecuteQuery();

                                string itemTitle = newItem["Title"].ToString();
                                int itemId = newItem.Id;
                                titleToIdMap[itemTitle] = itemId;

                                Console.WriteLine($"Item '{itemTitle}' added to '{list.Title}' list from Excel data.");
                            }
                            catch (Exception ex)
                            {
                                Console.ForegroundColor = ConsoleColor.Yellow;
                                Console.WriteLine($"Error adding item from Excel data: {ex.Message}");
                                Console.WriteLine($"Row causing the error: {row}");
                            }
                        }
                    }
                    else
                    {
                        Console.WriteLine("No worksheet found in the Excel file.");
                    }
                }
            }
            catch (Exception ex)
            {
                Console.ForegroundColor = ConsoleColor.Red;
                Console.WriteLine($"Error processing Excel data: {ex.Message}");
            }

            return titleToIdMap;
        }

        private static void UpdateParentLookup(ClientContext ctx, List list, string excelFilePath, Dictionary<string, int> titleToIdMap)
        {
            try
            {
                FileInfo excelFile = new FileInfo(excelFilePath);
                using (ExcelPackage package = new ExcelPackage(excelFile))
                {
                    ExcelWorksheet worksheet = package.Workbook.Worksheets.FirstOrDefault();

                    if (worksheet != null)
                    {
                        int rowCount = worksheet.Dimension.Rows;
                        int columnCount = worksheet.Dimension.Columns;

                        var excelHeaders = new List<string>();
                        for (int col = 1; col <= columnCount; col++)
                        {
                            string header = worksheet.Cells[1, col].Value?.ToString();
                            if (!string.IsNullOrEmpty(header))
                            {
                                excelHeaders.Add(header.Trim());
                            }
                        }

                        if (!excelHeaders.Contains("Title") || !excelHeaders.Contains("Parent"))
                        {
                            Console.WriteLine("Excel sheet must contain 'Title' and 'Parent' columns.");
                            return;
                        }

                        for (int row = 2; row <= rowCount; row++)
                        {
                            try
                            {
                                string itemTitle = worksheet.Cells[row, excelHeaders.IndexOf("Title") + 1].Value?.ToString();
                                string parentTitle = worksheet.Cells[row, excelHeaders.IndexOf("Parent") + 1].Value?.ToString();

                                if (!string.IsNullOrEmpty(itemTitle))
                                {
                                    if (titleToIdMap.ContainsKey(itemTitle))
                                    {
                                        ListItem listItem = list.GetItemById(titleToIdMap[itemTitle]);
                                        ctx.Load(listItem, i => i["Title"], i => i["Parent"]);
                                        ctx.ExecuteQuery();

                                        if (!string.IsNullOrEmpty(parentTitle) && titleToIdMap.ContainsKey(parentTitle))
                                        {
                                            int parentId = titleToIdMap[parentTitle];

                                            listItem["Parent"] = new FieldLookupValue { LookupId = parentId };
                                            listItem.Update();
                                            ctx.ExecuteQuery();

                                            Console.WriteLine($"Parent '{parentTitle}' set for item '{itemTitle}' in '{list.Title}' list.");
                                        }
                                        else if (!string.IsNullOrEmpty(parentTitle))
                                        {
                                            Console.WriteLine($"Parent item '{parentTitle}' not found in '{list.Title}' list.");
                                        }
                                    }
                                    else
                                    {
                                        Console.WriteLine($"Item '{itemTitle}' not found in '{list.Title}' list.");
                                    }
                                }
                            }
                            catch (Exception ex)
                            {
                                Console.ForegroundColor = ConsoleColor.Yellow;
                                Console.WriteLine($"Error setting parent for item from Excel data: {ex.Message}");
                                Console.WriteLine($"Row causing the error: {row}");
                            }
                        }
                    }
                    else
                    {
                        Console.WriteLine("No worksheet found in the Excel file.");
                    }
                }
            }
            catch (Exception ex)
            {
                Console.ForegroundColor = ConsoleColor.Red;
                Console.WriteLine($"Error updating Parent lookup: {ex.Message}");
            }
        }
    
        private static void ProcessList(ClientContext ctx, ListCollection lists, string listName, string excelFilePath)
        {
            SPList list = lists.FirstOrDefault(l => l.Title.Equals(listName, StringComparison.OrdinalIgnoreCase));
            if (list != null)
            {
                Console.WriteLine($"List '{listName}' found.");
                ProcessExcelData(ctx, list, excelFilePath);
            }
            else
            {
                Console.WriteLine($"List '{listName}' not found.");
            }
        }

        private static void ProcessExcelData(ClientContext ctx, SPList list, string excelFilePath)
        {
            try
            {
                // Set the license context
                ExcelPackage.LicenseContext = LicenseContext.NonCommercial;
                FileInfo excelFile = new FileInfo(excelFilePath);
                using (ExcelPackage package = new ExcelPackage(excelFile))
                {
                    ExcelWorksheet worksheet = package.Workbook.Worksheets.FirstOrDefault();

                    if (worksheet != null)
                    {
                        int rowCount = worksheet.Dimension.Rows;
                        int columnCount = worksheet.Dimension.Columns;

                        // Read headers from the first row of Excel sheet
                        List<string> excelHeaders = new List<string>();
                        for (int col = 1; col <= columnCount; col++)
                        {
                            string header = worksheet.Cells[1, col].Value?.ToString();
                            if (!string.IsNullOrEmpty(header))
                            {
                                excelHeaders.Add(header.Trim()); // Trim to remove extra spaces
                            }
                        }

                        // Iterate through rows starting from the second row
                        for (int row = 2; row <= rowCount; row++)
                        {
                            try
                            {
                                // Create a new list item
                                ListItemCreationInformation itemCreateInfo = new ListItemCreationInformation();
                                ListItem newItem = list.AddItem(itemCreateInfo);

                                // Fill item properties dynamically based on Excel headers
                                foreach (string header in excelHeaders)
                                {
                                    string excelValue = worksheet.Cells[row, excelHeaders.IndexOf(header) + 1].Value?.ToString();
                                    if (!string.IsNullOrEmpty(excelValue))
                                    {
                                        newItem[header] = excelValue;
                                    }
                                }

                                newItem.Update();
                                ctx.ExecuteQuery();

                                Console.WriteLine($"Item added to '{list.Title}' list from Excel data.");
                            }
                            catch (Exception ex)
                            {
                                Console.ForegroundColor = ConsoleColor.Yellow;
                                Console.WriteLine($"Error adding item from Excel data: {ex.Message}");
                                Console.WriteLine($"Row causing the error: {row}");
                            }
                        }
                    }
                    else
                    {
                        Console.WriteLine("No worksheet found in the Excel file.");
                    }
                }
            }
            catch (Exception ex)
            {
                Console.ForegroundColor = ConsoleColor.Red;
                Console.WriteLine($"Error processing Excel data: {ex.Message}");
            }
        }


        private static void MakeDynamicListTemplate(string webUrl, string fileName, string listname)
        {
            // Configure the XML file system provider
            XMLTemplateProvider provider = new XMLFileSystemTemplateProvider("..\\..\\pnpTemplate", "");
            ProvisioningTemplate template = provider.GetTemplate(fileName);
            string newXML = template.ToXML();
            newXML = newXML.Replace("txtListName", listname).Replace("GUID1", Guid.NewGuid().ToString());
            newXML = newXML.Replace("GUID_TimeSheetCol", Guid.NewGuid().ToString());
            newXML = newXML.Replace("GUID_DocumentCol", Guid.NewGuid().ToString());

            MemoryStream stream = new MemoryStream();
            StreamWriter writer = new StreamWriter(stream);
            writer.Write(newXML);
            writer.Flush();
            stream.Position = 0;

            ProvisioningTemplate template1 = provider.GetTemplate(stream);

            ApplyTemplateByXMLString(webUrl, template1, listname);

        }

        private static void ApplyTemplateByXMLString(string targetWebUrl, ProvisioningTemplate template, string lstName)
        {

            var authManager = new OfficeDevPnP.Core.AuthenticationManager();
            // This method calls a pop up window with the login page and it also prompts  
            // for the multi factor authentication code.  
            ClientContext context = authManager.GetWebLoginClientContext(targetWebUrl);

            Web web = context.Web;
            context.Load(web, w => w.Title);
            context.ExecuteQueryRetry();

            ProvisioningTemplateApplyingInformation ptai
                = new ProvisioningTemplateApplyingInformation();
            ptai.ProgressDelegate = delegate (String message, Int32 progress, Int32 total)
            {
                Console.WriteLine("{0:00}/{1:00} – {2}", progress, total, message);
            };

            web.ApplyProvisioningTemplate(template, ptai);

            //Add Entry in SMD List
            CreateItemInSMDList(context, targetWebUrl, lstName);
            //Update Timesheet Configuration entry in SMD                
            UpdateTimeListConfig(context, lstName);


        }

        private static void CreateItemInSMDList(ClientContext context, string targetWebUrl, string lstName)
        {

            Microsoft.SharePoint.Client.List TaskTimesheet = context.Web.Lists.GetByTitle("TaskTimesheet");
            context.Load(TaskTimesheet);
            context.ExecuteQuery();

            Microsoft.SharePoint.Client.List tlstName = context.Web.Lists.GetByTitle(lstName);
            context.Load(tlstName);
            context.ExecuteQuery();

            Microsoft.SharePoint.Client.List SMDList = context.Web.Lists.GetByTitle("SmartMetadata");
            ListItemCreationInformation itemCreateInfo = new ListItemCreationInformation();
            Microsoft.SharePoint.Client.ListItem listItem = SMDList.AddItem(itemCreateInfo);

            var json = new object();
            if (timesheetYes.ToLower() == "yes")
            {
                var obj = new
                {
                    Title = tlstName.Title,
                    listId = tlstName.Id.ToString().Replace('{', ' ').Replace('}', ' ').Trim(),
                    siteName = tlstName.Title,
                    siteUrl = targetWebUrl,
                    TaxType = "Sites",
                    DomainUrl = targetWebUrl,
                    MetadataName = tlstName.ListItemEntityTypeFullName,
                    TimesheetListName = "TaskTimesheet",
                    TimesheetListId = TaskTimesheet.Id.ToString().Replace('{', ' ').Replace('}', ' ').Trim(),
                    TimesheetListmetadata = "SP.Data.TaskTimesheetListItem",
                    ImageUrl = targetWebUrl + "/SiteCollectionImages/ICONS/Foundation/SH_icon.png"
                };
                json = System.Text.Json.JsonSerializer.Serialize(obj);
            }
            else
            {
                var obj = new
                {
                    Title = tlstName.Title,
                    listId = tlstName.Id.ToString().Replace('{', ' ').Replace('}', ' ').Trim(),
                    siteName = tlstName.Title,
                    siteUrl = targetWebUrl,
                    TaxType = "Sites",
                    DomainUrl = targetWebUrl,
                    MetadataName = tlstName.ListItemEntityTypeFullName,
                    ImageUrl = targetWebUrl + "/SiteCollectionImages/ICONS/Foundation/SH_icon.png"
                };
                json = System.Text.Json.JsonSerializer.Serialize(obj);
            }

            // fill the fields with the required information
            listItem["Title"] = tlstName.Title;
            listItem["Description"] = "test";
            listItem["SortOrder"] = "4";
            listItem["SmartFilters"] = "Dashboard, Portfolio, Advanced Search";
            listItem["TaxType"] = "Sites";
            listItem["ParentID"] = 0;
            listItem["IsVisible"] = 1;
            listItem["_Status"] = "Not Started";
            listItem["Selectable"] = 1;
            listItem["listId"] = tlstName.Id.ToString().Replace('{', ' ').Replace('}', ' ').Trim();
            listItem["siteName"] = tlstName.Title;
            listItem["siteUrl"] = targetWebUrl;
            listItem["Configurations"] = json;

            listItem.Update();
            context.ExecuteQuery();
            Console.WriteLine("Entry created in Smart Metadata list.");

        }

        private static void UpdateTimeListConfig(ClientContext context, string lstName)
        {
            //string siteUrl = "https://smalsusinfolabs.sharepoint.com/sites/HHHHQA/SmartManagement";
            //var authManager = new OfficeDevPnP.Core.AuthenticationManager();
            // This method calls a pop up window with the login page and it also prompts  
            // for the multi factor authentication code.  
            //ClientContext context = authManager.GetWebLoginClientContext(siteUrl);
            Microsoft.SharePoint.Client.List spSMD = context.Web.Lists.GetByTitle("SmartMetadata");
            CamlQuery camlQuery = new CamlQuery();
            camlQuery.ViewXml = "<View><Query><Where><Eq><FieldRef Name='Title'/>" +
                "<Value Type='Text'>Timesheet List Configrations</Value></Eq></Where></Query></View>";

            ListItemCollection smtdata = spSMD.GetItems(camlQuery);
            context.Load(smtdata);
            context.ExecuteQuery();

            try
            {
                if (smtdata.Count > 0)
                {
                    foreach (Microsoft.SharePoint.Client.ListItem oItem in smtdata)
                    {
                        Console.WriteLine("Get data - " + oItem["Title"]);
                        List<TimeSheetConfig> tdata = System.Text.Json.JsonSerializer.Deserialize<List<TimeSheetConfig>>(oItem["Configurations"].ToString());

                        string[] a = new string[] { lstName };
                        var b = tdata[0].taskSites.Union(a);
                        var tList = tdata[0].taskSites.ToList();
                        // Add Item To List
                        tList.Add(lstName);

                        // Convert back to List
                        tdata[0].taskSites = tList.ToArray();
                        string query = tdata[0].query;
                        query = query.Split('&')[0] + ",Task" + lstName + "/Id,Task" + lstName + "/Title&" + query.Split('&')[1] + ",Task" + lstName;
                        tdata[0].query = query;
                        string dt = Newtonsoft.Json.JsonConvert.SerializeObject(tdata);

                        oItem["Configurations"] = dt;
                        oItem.Update();
                        context.ExecuteQuery();
                        Console.WriteLine("Value upated for Timesheet List Configrations");

                    }
                }
            }
            catch (Exception)
            {

                throw;
            }
        }

        public static void TestSample()
        {


            //
            //details for hhhhteams tenant
            //username = "developer1@hochhuth-consulting.de";
            //password = "Fub62326";

            //details for QA tenant
            username = "chetan@smalsusinfolabs.onmicrosoft.com";
            password = "Bug88082";
            //ApplyTemplateOnWeb("https://smalsusinfolabs.sharepoint.com/sites/HHHHQA/SmartManagement", "PnPProvisioning_SM.xml");

            //Fix lookup Columns
            //Console.WriteLine("Validating the lookup in Lists...");
            //ApplyTemplateOnWeb(webUrl, "PnPProvisioning_S5_FixLookup.xml"); for SH
            //ApplyTemplateOnWeb("https://smalsusinfolabs.sharepoint.com/sites/HHHHQA/ST", "PnPProvisioning_Fix_Lookup.xml");

            //Console.WriteLine("Updateing Parent Id in SmartMetadata List");
            //SetParentforSmartMetaData("https://smalsusinfolabs.sharepoint.com/sites/HHHHQA/SH", "SmartMetadata", "SmartMetaData.json");

            //Console.WriteLine("Updateing Parent Id in Top Navigation List");
            //SetParentforSmartMetaData("https://smalsusinfolabs.sharepoint.com/sites/HHHHQA/SH", "TopNavigation", "TopNavigation.json");
            //GetData();
            //provision the root site
            //ProvisionRootSite("https://smalsusinfolabs.sharepoint.com/sites/HHHHQA");

            //GetDataFromRequest("https://hhhhteams.sharepoint.com/sites/HHHH/SH");
            //MakeDynamicListTemplate("PnPProvisioning_DynamicList.xml");

            //siteProvisioing(); // for first time site creation.

            //In case of site provisioning agi9n on created subsite.
            //In case of site provisioning agi9n on created subsite.
            /*

            username = "developer1@hochhuth-consulting.de";
            password = "Fub62326";

             ApplyTemplateOnWeb("https://hhhhteams.sharepoint.com/sites/HHHH/Test1", "PnPProvisioning_S5_NewLists.xml");


             //ApplyTemplateOnWeb("https://hhhhteams.sharepoint.com/sites/HHHH/SH", "PnPProvisioning_S5_FixLookup.xml");



             //SetParentforSmartMetaData("https://hhhhteams.sharepoint.com/sites/HHHH/SH", "TopNavigation", "TopNavigation.json");

            */
            //Extract the web template
            //GetProvisioningTemplateRoot("https://hhhhteams.sharepoint.com/sites/HHHH");
            //GetProvisioningTemplate("https://smalsusinfolabs.sharepoint.com/sites/HHHHQA/SP");

            //Set For PropertyBag

            //ApplyTemplateOnWeb("https://hhhhteams.sharepoint.com/sites/TestProvisioning/S10", "PnPProvisioning_S5_Props.xml");
        }

        private static void GetDataFromRequest(string targetWebUrl)
        {
            using (ClientContext context = new ClientContext(targetWebUrl))
            {
                //SecureString securePassword = GetSecureString("Fub62326");
                //context.Credentials = new SharePointOnlineCredentials("developer1@hochhuth-consulting.de", securePassword);
                //context.RequestTimeout = System.Threading.Timeout.Infinite;
                SecureString securePassword = GetSecureString(password);
                context.Credentials = new SharePointOnlineCredentials(username, securePassword);
                context.RequestTimeout = System.Threading.Timeout.Infinite;

                Microsoft.SharePoint.Client.List spRequestList = context.Web.Lists.GetByTitle("ProvisioningRequest");


                CamlQuery camlQuery = new CamlQuery();
                camlQuery.ViewXml = "<View><Query><Where><Eq><FieldRef Name='Status'/>" +
                    "<Value Type='Text'>Pending</Value></Eq></Where></Query></View>";

                ListItemCollection lstItems = spRequestList.GetItems(camlQuery);
                context.Load(lstItems);
                context.ExecuteQuery();

                foreach (Microsoft.SharePoint.Client.ListItem item in lstItems)
                {
                    Console.WriteLine("Provisioning Tasks List Name - " + item.FieldValues["TaskLists"].ToString());
                    string[] tLists = item.FieldValues["TaskLists"].ToString().Split(new char[] { ';' });
                    Microsoft.SharePoint.Client.ListItem targetListItem = spRequestList.GetItemById((int)item["ID"]);

                    if (tLists.Length > 0)
                    {
                        targetListItem["Status"] = "InProgress";
                        targetListItem.Update();
                        context.ExecuteQuery();

                        for (int i = 0; i < tLists.Length - 1; i++)
                        {
                            Console.WriteLine("Provisioning Tasks List Name - " + tLists[i]);
                            //MakeDynamicListTemplate("PnPProvisioning_DynamicList.xml", tLists[i]);
                            Console.WriteLine("Task List Created. Adding in SMD List..");
                            CreateItemInSMDList(context, targetWebUrl, tLists[i]);
                        }


                        targetListItem["Status"] = "Completed";
                        targetListItem.Update();
                        context.ExecuteQuery();
                        Console.WriteLine("------------------------------------");
                    }

                }
            }
        }

        private static void ReadJsonFile()
        {

            string text = System.IO.File.ReadAllText("..\\..\\MasterData\\SmartMetaData.json");

            List<SMDData> smtdata = System.Text.Json.JsonSerializer.Deserialize<List<SMDData>>(text);

            if (smtdata.Count > 0)
            {
                for (int i = 0; i < smtdata.Count; i++)
                {
                    Console.WriteLine("Parent : " + smtdata[i].Title);
                    if (smtdata[i].Child.Count > 0)
                    {
                        for (int j = 0; j < smtdata[i].Child.Count; j++)
                        {
                            Console.WriteLine("Child : " + smtdata[i].Child[j].Title);
                        }
                    }
                    Console.WriteLine("------------------------------");
                }
            }

            ConsoleKeyInfo ch;
            Console.WriteLine("Press the Escape (Esc) key to quit: \n");
            ch = Console.ReadKey();
            if (ch.Key == ConsoleKey.Escape)
            {
                Environment.Exit(0);
            }
        }

        private static string GetInput(string label, bool isPassword, ConsoleColor defaultForeground)
        {
            Console.ForegroundColor = ConsoleColor.Green;
            Console.WriteLine("{0} : ", label);
            Console.ForegroundColor = defaultForeground;

            string value = "";

            for (ConsoleKeyInfo keyInfo = Console.ReadKey(true); keyInfo.Key != ConsoleKey.Enter; keyInfo = Console.ReadKey(true))
            {
                if (keyInfo.Key == ConsoleKey.Backspace)
                {
                    if (value.Length > 0)
                    {
                        value = value.Remove(value.Length - 1);
                        Console.SetCursorPosition(Console.CursorLeft - 1, Console.CursorTop);
                        Console.Write(" ");
                        Console.SetCursorPosition(Console.CursorLeft - 1, Console.CursorTop);
                    }
                }
                else if (keyInfo.Key != ConsoleKey.Enter)
                {
                    if (isPassword)
                    {
                        Console.Write("*");
                    }
                    else
                    {
                        Console.Write(keyInfo.KeyChar);
                    }
                    value += keyInfo.KeyChar;
                }
            }
            Console.WriteLine("");

            return value;
        }

        //to provision new site collections for content typs and columns
        static void ProvisionRootSite(string siteURL)
        {
            Console.WriteLine("Start Provisioing : {0:hh.mm.ss}", DateTime.Now);
            try
            {
                ApplyTemplateOnWeb(siteURL, "PnPProvisioning_main_root.xml");

            }
            catch (Exception ex)
            {
                Console.WriteLine("Something went wrong. " + ex.Message);
            }

            Console.WriteLine("End Provisioing : {0:hh.mm.ss}", DateTime.Now);
        }

        public static SecureString GetSecureString(string userPassword)
        {
            SecureString securePassword = new SecureString();

            foreach (char c in userPassword.ToCharArray())
            {
                securePassword.AppendChar(c);
            }

            return securePassword;
        }

        private static void GetProvisioningTemplate(string webUrl)
        {
            var authManager = new OfficeDevPnP.Core.AuthenticationManager();
            // This method calls a pop up window with the login page and it also prompts  
            // for the multi factor authentication code.  

            ClientContext ctx = authManager.GetWebLoginClientContext(webUrl);
            //static string username;// = "developer1@hochhuth-consulting.de";
            //static string password;// = "Fub62326";

            //using (var ctx = new ClientContext(webUrl))
            //{
            //SecureString securePassword = GetSecureString("Fub62326");
            //ctx.Credentials = new SharePointOnlineCredentials("developer1@hochhuth-consulting.de", securePassword);

            // Just to output the site details
            Web web = ctx.Web;
            ctx.Load(web, w => w.Title);
            ctx.ExecuteQueryRetry();

            Console.ForegroundColor = ConsoleColor.White;
            Console.WriteLine("Your site title is:" + ctx.Web.Title);
            Console.ForegroundColor = Console.ForegroundColor;

            ProvisioningTemplateCreationInformation ptci
                    = new ProvisioningTemplateCreationInformation(ctx.Web);

            // Create FileSystemConnector to store a temporary copy of the template 
            ptci.FileConnector = new FileSystemConnector(@"c:\temp\pnpprovisioningdemo", "");
            ptci.PersistBrandingFiles = false;
            ptci.PersistPublishingFiles = false;
            //ptci.PersistComposedLookFiles = true;
            //ptci.IncludeSiteGroups = true;
            //ptci.HandlersToProcess = Handlers.SiteSecurity;
            ptci.IncludeAllClientSidePages = true; //if you need to get Pages
                                                   //ptci.HandlersToProcess = Handlers.Lists;

            //ptci.HandlersToProcess = Handlers.None;
            //ptci.HandlersToProcess ^= Handlers.Lists;
            // ptci.HandlersToProcess ^= Handlers.Fields;
            // ptci.HandlersToProcess ^= Handlers.ContentTypes;
            //ptci.HandlersToProcess = Handlers.All;
            //ptci.IncludeNativePublishingFiles = true;

            ptci.HandlersToProcess = Handlers.Pages | Handlers.PageContents; //if onlye modern page needs to extracted
                                                                             //ptci.HandlersToProcess = Handlers.Lists;

            //ptci.HandlersToProcess = Handlers.Pages;
            //ptci.HandlersToProcess = Handlers.PageContents; //if onlye modern page needs to extracted

            ptci.ProgressDelegate = delegate (String message, Int32 progress, Int32 total)
            {
                // Only to output progress for console UI
                Console.WriteLine("{0:00}/{1:00} – {2}", progress, total, message);
            };

            // Execute actual extraction of the template
            ProvisioningTemplate template = ctx.Web.GetProvisioningTemplate(ptci);

            // We can serialize this template to save and reuse it
            // Optional step 
            XMLTemplateProvider provider =
                    new XMLFileSystemTemplateProvider(@"c:\temp\pnpprovisioningdemo", "");
            provider.SaveAs(template, "PnPProvisioning_SP_Pages2.xml");

            //return template;
            // }
        }

        private static void GetProvisioningTemplateRoot(string webUrl)
        {
            //var authManager = new OfficeDevPnP.Core.AuthenticationManager();
            // This method calls a pop up window with the login page and it also prompts  
            // for the multi factor authentication code.  

            //ClientContext ctx = authManager.GetWebLoginClientContext(webUrl);
            //static string username;// = "developer1@hochhuth-consulting.de";
            //static string password;// = "Fub62326";

            using (var ctx = new ClientContext(webUrl))
            {
                SecureString securePassword = GetSecureString("Fub62326");
                ctx.Credentials = new SharePointOnlineCredentials("developer1@hochhuth-consulting.de", securePassword);

                // Just to output the site details
                Web web = ctx.Web;
                ctx.Load(web, w => w.Title);
                ctx.ExecuteQueryRetry();

                Console.ForegroundColor = ConsoleColor.White;
                Console.WriteLine("Your site title is:" + ctx.Web.Title);
                Console.ForegroundColor = Console.ForegroundColor;

                ProvisioningTemplateCreationInformation ptci
                        = new ProvisioningTemplateCreationInformation(ctx.Web);

                // Create FileSystemConnector to store a temporary copy of the template 
                ptci.FileConnector = new FileSystemConnector(@"c:\temp\pnpprovisioningdemo", "");
                ptci.PersistBrandingFiles = false;
                ptci.PersistPublishingFiles = false;
                //ptci.PersistComposedLookFiles = true;
                //ptci.IncludeSiteGroups = true;
                //ptci.HandlersToProcess = Handlers.SiteSecurity;
                ptci.IncludeAllClientSidePages = true; //if you need to get Pages
                ptci.HandlersToProcess = Handlers.All;

                //ptci.HandlersToProcess = Handlers.None;
                //ptci.HandlersToProcess ^= Handlers.Lists;
                // ptci.HandlersToProcess ^= Handlers.Fields;
                // ptci.HandlersToProcess ^= Handlers.ContentTypes;
                // ptci.HandlersToProcess = Handlers.All;
                //ptci.IncludeNativePublishingFiles = true;

                //ptci.HandlersToProcess = Handlers.Pages | Handlers.PageContents; //if onlye modern page needs to extracted
                //ptci.HandlersToProcess = Handlers.Lists;

                //ptci.HandlersToProcess = Handlers.Pages;
                //ptci.HandlersToProcess = Handlers.PageContents; //if onlye modern page needs to extracted

                ptci.ProgressDelegate = delegate (String message, Int32 progress, Int32 total)
                {
                    // Only to output progress for console UI
                    Console.WriteLine("{0:00}/{1:00} – {2}", progress, total, message);
                };

                // Execute actual extraction of the template
                ProvisioningTemplate template = ctx.Web.GetProvisioningTemplate(ptci);

                // We can serialize this template to save and reuse it
                // Optional step 
                XMLTemplateProvider provider =
                        new XMLFileSystemTemplateProvider(@"c:\temp\pnpprovisioningdemo", "");
                provider.SaveAs(template, "PnPProvisioning_main_root.xml");

                //return template;
            }
        }

        private static void ApplyPnPTemplate(string targetWebUrl, string fileName)
        {
            //var authManager = new OfficeDevPnP.Core.AuthenticationManager();
            // This method calls a pop up window with the login page and it also prompts  
            // for the multi factor authentication code.  
            //ClientContext context = authManager.GetWebLoginClientContext(targetWebUrl);
            //using (var context = new ClientContext(targetWebUrl))
            //{
            //context.Credentials = new SharePointOnlineCredentials(userName, pwd);
            using (ClientContext context = new ClientContext(targetWebUrl))
            {
                SecureString securePassword = GetSecureString("Fub62326");
                context.Credentials = new SharePointOnlineCredentials("developer1@hochhuth-consulting.de", securePassword);


                //SecureString securePassword = GetSecureString(password);
                //context.Credentials = new SharePointOnlineCredentials(username, securePassword);
                Web web = context.Web;
                context.Load(web, w => w.Title);
                context.ExecuteQueryRetry();

                // Configure the XML file system provider
                XMLTemplateProvider provider = new XMLFileSystemTemplateProvider(@"C:\temp\pnpprovisioningdemo", "");
                ProvisioningTemplate template = provider.GetTemplate(fileName);

                ProvisioningTemplateApplyingInformation ptai
                    = new ProvisioningTemplateApplyingInformation();
                ptai.ProgressDelegate = delegate (String message, Int32 progress, Int32 total)
                {
                    Console.WriteLine("{0:00}/{1:00} – {2}", progress, total, message);
                };

                // Associate file connector for assets
                FileSystemConnector connector = new FileSystemConnector(@"c:\temp\pnpprovisioningdemo", "");
                template.Connector = connector;
                web.ApplyProvisioningTemplate(template, ptai);

                Console.WriteLine("End: {0:hh.mm.ss}", DateTime.Now);
            }
        }

        private static void ApplyTemplate(string targetWebUrl, string fileName)
        {
            //var authManager = new OfficeDevPnP.Core.AuthenticationManager();
            // This method calls a pop up window with the login page and it also prompts  
            // for the multi factor authentication code.  
            //ClientContext context = authManager.GetWebLoginClientContext(targetWebUrl);
            //using (var context = new ClientContext(targetWebUrl))
            //{
            //context.Credentials = new SharePointOnlineCredentials(userName, pwd);
            using (ClientContext context = new ClientContext(targetWebUrl))
            {
                SecureString securePassword = GetSecureString("Fub62326");
                context.Credentials = new SharePointOnlineCredentials("developer1@hochhuth-consulting.de", securePassword);

                //SecureString securePassword = GetSecureString(password);
                //context.Credentials = new SharePointOnlineCredentials(username, securePassword);

                Web web = context.Web;
                context.Load(web, w => w.Title);
                context.ExecuteQueryRetry();

                // Configure the XML file system provider
                XMLTemplateProvider provider = new XMLFileSystemTemplateProvider(@"C:\temp\pnpprovisioningdemo", "");
                ProvisioningTemplate template = provider.GetTemplate(fileName);

                ProvisioningTemplateApplyingInformation ptai
                    = new ProvisioningTemplateApplyingInformation();
                ptai.ProgressDelegate = delegate (String message, Int32 progress, Int32 total)
                {
                    Console.WriteLine("{0:00}/{1:00} – {2}", progress, total, message);
                };

                // Associate file connector for assets
                FileSystemConnector connector = new FileSystemConnector(@"c:\temp\pnpprovisioningdemo", "");
                template.Connector = connector;
                //template.Security.SiteGroups.Add(
                web.ApplyProvisioningTemplate(template, ptai);

                //Console.WriteLine("Updateing Parent Id in SmartMetadata List");
                //SetParentforSmartMetaData(targetWebUrl);

                Console.WriteLine("End: {0:hh.mm.ss}", DateTime.Now);
            }
        }

        private static void ApplyTemplateOnRoot(string targetWebUrl, string fileName)
        {
            //var authManager = new OfficeDevPnP.Core.AuthenticationManager();
            // This method calls a pop up window with the login page and it also prompts  
            // for the multi factor authentication code.  
            //ClientContext context = authManager.GetWebLoginClientContext(targetWebUrl);
            //using (var context = new ClientContext(targetWebUrl))
            //{
            //context.Credentials = new SharePointOnlineCredentials(userName, pwd);
            Console.WriteLine("Start Provisioing : {0:hh.mm.ss}", DateTime.Now);
            using (ClientContext context = new ClientContext(targetWebUrl))
            {
                SecureString securePassword = GetSecureString("Fub62326");
                context.Credentials = new SharePointOnlineCredentials("developer1@hochhuth-consulting.de", securePassword);

                //SecureString securePassword = GetSecureString(password);
                //context.Credentials = new SharePointOnlineCredentials(username, securePassword);

                Web web = context.Web;
                context.Load(web, w => w.Title);
                context.ExecuteQueryRetry();

                // Configure the XML file system provider
                XMLTemplateProvider provider = new XMLFileSystemTemplateProvider("..\\..\\pnpTemplate", "");
                ProvisioningTemplate template = provider.GetTemplate(fileName);

                ProvisioningTemplateApplyingInformation ptai
                    = new ProvisioningTemplateApplyingInformation();
                ptai.ProgressDelegate = delegate (String message, Int32 progress, Int32 total)
                {
                    Console.WriteLine("{0:00}/{1:00} – {2}", progress, total, message);
                };

                // Associate file connector for assets
                FileSystemConnector connector = new FileSystemConnector("..\\..\\pnpTemplate", "");
                template.Connector = connector;
                //template.Security.SiteGroups.Add(
                web.ApplyProvisioningTemplate(template, ptai);

                Console.WriteLine("End Provisioing : {0:hh.mm.ss}", DateTime.Now);
            }
        }

        private static void SetParentforSmartMetaData(string targetWebUrl, string ListName, string fileName)
        {
            //SmartMetaData.json
            string text = System.IO.File.ReadAllText("..\\..\\MasterData\\" + fileName);
            List<SMDData> smtdata = System.Text.Json.JsonSerializer.Deserialize<List<SMDData>>(text);

            var authManager = new OfficeDevPnP.Core.AuthenticationManager();
            // This method calls a pop up window with the login page and it also prompts  
            // for the multi factor authentication code.  
            ClientContext context = authManager.GetWebLoginClientContext(targetWebUrl);

            //using (ClientContext context = new ClientContext(targetWebUrl))
            //{
            //SecureString securePassword = GetSecureString("Fub62326");
            //context.Credentials = new SharePointOnlineCredentials("developer1@hochhuth-consulting.de", securePassword);

            //SecureString securePassword = GetSecureString(password);
            //context.Credentials = new SharePointOnlineCredentials(username, securePassword);

            Microsoft.SharePoint.Client.List spSMD = context.Web.Lists.GetByTitle(ListName);

            CamlQuery oQuery = CamlQuery.CreateAllItemsQuery();

            ListItemCollection oCollection = spSMD.GetItems(oQuery);
            context.Load(oCollection);
            context.ExecuteQuery();
            Console.WriteLine("Total Count: " + oCollection.Count);
            try
            {
                if (smtdata.Count > 0)
                {
                    for (int i = 0; i < smtdata.Count; i++)
                    {
                        Console.WriteLine("Parent : " + smtdata[i].Title);
                        string parentName = smtdata[i].Title;
                        int parnetID = 0;
                        //Get Item ID of the Parent Element
                        foreach (Microsoft.SharePoint.Client.ListItem oItem in oCollection)
                        {
                            if (oItem["ParentID"] != null)
                            {
                                if (oItem["Title"].ToString() == "In preparation")
                                {
                                    parnetID = oItem.Id;
                                    break;
                                }
                                if (parentName == oItem["Title"].ToString() && oItem["ParentID"].ToString() == "0")
                                {
                                    parnetID = oItem.Id;
                                    break;
                                }
                            }

                        }
                        //Update Parent Id of the child element
                        if (parnetID != 0 && smtdata[i].Child.Count > 0)
                        {
                            for (int j = 0; j < smtdata[i].Child.Count; j++)
                            {
                                string childName = smtdata[i].Child[j].Title;
                                Console.WriteLine("Child : " + smtdata[i].Child[j].Title);
                                foreach (Microsoft.SharePoint.Client.ListItem oItem in oCollection)
                                {

                                    if (childName == oItem["Title"].ToString() && oItem["ParentID"] == null)
                                    {

                                        int childID = oItem.Id;
                                        //UpdateItem
                                        Microsoft.SharePoint.Client.ListItem oListItem = spSMD.GetItemById(childID);
                                        oListItem["Parent"] = parnetID;
                                        oListItem["ParentID"] = parnetID;
                                        oListItem.Update();
                                        context.ExecuteQuery();
                                        Console.WriteLine("Parent ID updated for " + childName);
                                        break;
                                    }

                                }
                            }
                        }
                        Console.WriteLine("------------------------------");
                    }
                }
            }
            catch (Exception)
            {

                throw;
            }


            //}
        }
    }


    public class SMDData
    {
        public string Title { get; set; } = string.Empty;
        public List<ChildName> Child { get; set; }
    }

    public class ChildName
    {
        public string Title { get; set; } = string.Empty;
    }
    public class TimeSheetConfig
    {
        public string listName { get; set; } = string.Empty;
        public string listId { get; set; } = string.Empty;
        public string siteUrl { get; set; } = string.Empty;
        public string query { get; set; } = string.Empty;
        public string[] taskSites { get; set; } = { };

    }
}
