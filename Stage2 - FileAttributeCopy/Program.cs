using System;
using System.Collections.Generic;
using System.Data.Services.Client;
using System.Linq;
using System.Runtime.Remoting.Contexts;
using System.Security;
using System.Text;
using System.Threading.Tasks;
using Microsoft.Online.SharePoint.Client;
using Microsoft.SharePoint.Client;
using Stage2___FileAttributeSync.Properties;
using SP = Microsoft.SharePoint.Client;

namespace Stage2___FileAttributeCopy
{
    class Program
    {
        static void Main(string[] args)
        {
            //GetAllItem();
            InitializeFile();
            GetAllItemInBatch();
            
           
            Console.ReadLine();
        }

        private static void InitializeFile()
        {
            var exportFileNameWithPath = Settings.Default.MetadataFileExportLocation;
            string s = ("ItemID,FilePath,CreatedOn,CreatedBy,ModifiedOn,ModifiedBy" + Environment.NewLine);
            System.IO.File.AppendAllText(exportFileNameWithPath, s);
        }

        private static void GetAllItemInBatch()
        {
            bool ColorToggle = true;

            var siteUrl = Settings.Default.SourceSPOSiteURL;
            var doclib = Settings.Default.DocLibDisplayName;
            var uname = Settings.Default.Office365Username;
            var pss = ConvertToSecureString(Settings.Default.Office365Password);
            var exportFileNameWithPath = Settings.Default.MetadataFileExportLocation;
            var SPOQueryBatchSize = Settings.Default.SPOQueryBatchSize;

            ClientContext clientContext = new ClientContext(siteUrl);
            clientContext.Credentials = new SharePointOnlineCredentials(uname, pss);

            SP.List oList = clientContext.Web.Lists.GetByTitle(doclib);

            ListItemCollectionPosition itemPosition = null;

            while (true)
            {
                if (ColorToggle)
                { ColorToggle = false; }
                else
                { ColorToggle = true; }

                toggleColor(ColorToggle);
                    

                CamlQuery camlQuery = new CamlQuery();

                camlQuery.ListItemCollectionPosition = itemPosition;

                camlQuery.ViewXml = "<View Scope='Recursive'><ViewFields><FieldRef Name='ID'/>" +
                    "<FieldRef Name='FileLeafRef'/><FieldRef Name='FileRef'/>" +
                    "<FieldRef Name='Created'/><FieldRef Name='Author'/>" +
                    "<FieldRef Name='Modified'/><FieldRef Name='Editor'/>" +
                    "</ViewFields><RowLimit>"+SPOQueryBatchSize+"</RowLimit></View>";

                ListItemCollection collListItem = oList.GetItems(camlQuery);

                clientContext.Load(collListItem);
                

                clientContext.ExecuteQuery();

                itemPosition = collListItem.ListItemCollectionPosition;

                foreach (ListItem oListItem in collListItem)
                {
                    string filename = oListItem["FileRef"].ToString();

                    string created = oListItem["Created"].ToString();

                    FieldUserValue user = (FieldUserValue)oListItem["Author"];
                    string author = user.LookupValue;

                    string Modified = oListItem["Modified"].ToString();
                    user = null;

                    user = (FieldUserValue)oListItem["Editor"];
                    string Editor = user.LookupValue;


                    string s = ($" {oListItem.Id},{filename},{created},{author},{Modified},{Editor}"); 
                    s = s + Environment.NewLine;
                    System.IO.File.AppendAllText(exportFileNameWithPath, s);
                    Console.WriteLine(s);

                    
                }

                if (itemPosition == null)
                {
                    break;
                }

                //Console.WriteLine("\n" + itemPosition.PagingInfo + "\n");
            }
            Console.ResetColor();
        }

        private static void toggleColor(bool State)
        {
            

            if (State)
            {
                Console.ForegroundColor = ConsoleColor.White;
            }
            else
            {
                Console.ForegroundColor = ConsoleColor.Green;
            }
        }

        private static void GetAllItem()
        {
            var siteurl = "https://m365x938597.sharepoint.com/sites/sourceSite";
            var doclib = "source1";
            var uname = "admin@M365x938597.onmicrosoft.com";
            var unsecurePss = "B5TBg8Q4KX";
            var pss = ConvertToSecureString(unsecurePss);

            try
            {
                using (ClientContext ctx = new ClientContext(siteurl))
                {
                    ctx.Credentials = new SharePointOnlineCredentials(uname, pss);

                    List l = ctx.Web.GetListByTitle(doclib);
                    ctx.Load(l);


                    ListItemCollection LI = l.GetItems(CamlQuery.CreateAllItemsQuery());
                    ctx.Load(LI, eachItem => eachItem.Include(
                                                    item => item.Id,
                                                    item => item["FileRef"],
                                                    item => item["FileLeafRef"],
                                                    item => item.Folder.ServerRelativeUrl,
                                                    item => item["Created"],
                                                    item => item["Author"],
                                                    item => item["Modified"],
                                                    item => item["Editor"]));

                    ctx.ExecuteQueryRetry();

                    foreach (var item in LI.ToArray())
                    {
                        string filename = item["FileRef"].ToString();

                        string created = item["Created"].ToString();

                        FieldUserValue user = (FieldUserValue)item["Author"];
                        string author = user.LookupValue;

                        string Modified = item["Modified"].ToString();
                        user = null;

                        user = (FieldUserValue)item["Editor"];
                        string Editor = user.LookupValue;


                        string s = ($" {item.Id} {filename} {created} {author} {Modified} {Editor}");
                        s = s + Environment.NewLine;
                        System.IO.File.AppendAllText("c:\\export\\Allfiles.txt", s);
                        Console.WriteLine(s);
                    }



                }
            }
            catch (Exception E)
            {

                Console.WriteLine(E.InnerException);
            }
            Console.ReadLine();
        }
        private static SecureString ConvertToSecureString(string strPassword)
        {
            var secureStr = new SecureString();
            if (strPassword.Length > 0)
            {
                foreach (var c in strPassword.ToCharArray()) secureStr.AppendChar(c);
            }
            return secureStr;

        }

    }
}
