using Microsoft.SharePoint.Client;
using Stage2___FileAttributeSync.Properties;
using System;
using System.Security;
using SP = Microsoft.SharePoint.Client;

namespace Stage2___FileAttributeCopy
{
    class Program
    {
        static void Main(string[] args)
        {
            Log("Starting program");

            GetAllItemInBatch(); // gets all items
            UpdateFileProperty(); // updates all itesm via resulting PS script ( that user is prompted to run)

            var str = $"Finished! Open PS as admin and run the file {Settings.Default.ProgramWorkingDir}";

            Console.WriteLine(str);
            Log(str);

            Console.ReadLine();
        }

        private static void InitializeFile()
        {
            //initilise the CSV file
            var exportFileNameWithPath = Settings.Default.MetadataFileExportLocation; //read the settings in the config file

            string s = ("ItemID,FilePath,CreatedOn,CreatedBy,ModifiedOn,ModifiedBy" + Environment.NewLine);
            System.IO.File.AppendAllText(exportFileNameWithPath, s);
        }

        private static void InitializePSFile()
        {
            //initialise the PS file with required text

            string ps1_1 = "$filename = [System.IO.Path]::GetRandomFileName()";
            WriteToPSTempFiles(ps1_1);

            string ps1_2 = $@"$filename = ""PROGRESS_LOG_""+ $filename.Remove($filename.Length-4)+"".log""";
            WriteToPSTempFiles(ps1_2);

            string ps1_3 = $@"Start-Transcript -Path $filename -NoClobber";
            WriteToPSTempFiles(ps1_3);

            string ps1 = $@"$secStringPassword = ConvertTo-SecureString ""{Settings.Default.DestinationOffice365Password}"" -AsPlainText -Force";
            WriteToPSTempFiles(ps1);

            string ps2 = $@"; $credOject = New-Object System.Management.Automation.PSCredential (""{Settings.Default.DestinationOffice365Username}"", $secStringPassword)";
            WriteToPSTempFiles(ps2);

            string ps3 = $@"; Connect-PnPOnline -Url ""{Settings.Default.DestinationSPOSiteURL}"" -Credentials $credOject";
            WriteToPSTempFiles(ps3);

            string ps4 = $"Install-Module SharePointPnPPowerShellOnline";
            WriteToPSTempFiles(ps4);



        }
        private static void UpdateFileProperty()
        {

            //this func create a PS file, which when run with UPDATE gracefully all the file metadata to SPO

            string str = "Converting..Source Item to Destination";
            Console.WriteLine(str);
            Log(str);

            InitializePSFile(); // Add the basic code for PnP PS script file, login to right site, creds, start transcript etc

            var siteUrl = Settings.Default.DestinationSPOSiteURL;
            var doclib = Settings.Default.DestinationDocLibDisplayName;
            var uname = Settings.Default.DestinationOffice365Username;
            var pss = ConvertToSecureString(Settings.Default.DestinationOffice365Password);
            var exportFileNameWithPath = Settings.Default.MetadataFileExportLocation;

            ClientContext clientContext = new ClientContext(siteUrl);
            clientContext.Credentials = new SharePointOnlineCredentials(uname, pss);

            SP.List oList = clientContext.Web.Lists.GetByTitle(doclib);


            foreach (var item in System.IO.File.ReadAllLines(Settings.Default.MetadataFileExportLocation))
            {
                if (!item.StartsWith("ItemID"))
                {
                    DateTime spdate = DateTime.Now;

                    var temp = item.Split(',');
                    var ID = temp[0];
                    var FilePath = ConvertToDestinationRelativeURL(temp[1]); // call to convert SOURCE URL to DESTINATION URL

                    DateTime.TryParse(temp[2], out spdate);
                    var CreatedON = spdate.ToString("o"); //convert to format that SPO likes


                    DateTime.TryParse(temp[4], out spdate);
                    var ModifiedON = spdate.ToString("o");

                    var CreatedBy = temp[3];
                    var ModifiedBy = temp[5];


                    try
                    {

                        var File1 = clientContext.Web.GetFileByServerRelativeUrl(FilePath);

                        clientContext.Load(File1, f => f.ListItemAllFields.Id);

                        clientContext.ExecuteQueryRetry(10, 5000, "SPOMigration JOB"); //will retry in case of server busy for 10 times, will retry after 5 second gap, identitfy as this string or our services


                        Console.Write($"Sucessfully Converted. Old Item:");
                        Console.ForegroundColor = ConsoleColor.Yellow;
                        Console.Write($"{ID}");
                        Console.ResetColor();



                        Console.Write($" with New Item:");
                        Console.ForegroundColor = ConsoleColor.Green;
                        Console.WriteLine($"{File1.ListItemAllFields.Id}");
                        Console.ResetColor();

                        Log($"Sucessfully Converted. Old Item:{ID} with New Item:{File1.ListItemAllFields.Id}");

                        string PSCmd1 = $@"$ListName = ""{doclib}""";
                        string PSCmd2 = $@"; $itemid = ""{File1.ListItemAllFields.Id}""";
                        string PSCmd3 = $@"; $CreatedBy = ""{CreatedBy}""";
                        string PSCmd4 = $@"; $ModifiedBy = ""{ModifiedBy}""";
                        string PSCmd5 = $@"; $Created = ""{CreatedON}""";
                        string PSCmd6 = $@"; $Modified = ""{ModifiedON}""";

                        string PSCmd7 = $@"@{{""Created""=$Created; ""Modified""=$Modified; ""Author"" =$CreatedBy; ""Editor"" =$ModifiedBy; }}";
                        string PSCmd = string.Format(";Set-PnPListItem -List $ListName -Identity $itemid -Values {0}", PSCmd7);

                        string Fullcmd = PSCmd1 + PSCmd2 + PSCmd3 + PSCmd4 + PSCmd5 + PSCmd6 + PSCmd;

                        WriteToPSTempFiles(Fullcmd); // write to resulting SPO File

                    }
                    catch (Exception e)
                    {
                        Log($"FAILED TO Convert. Old Item:{ID}.");
                        Log($"{e.Message}");
                    }

                }

            }

            WriteToPSTempFiles("Stop-Transcript"); // PS transcription feature is OFF here, ON while initialising it
        }

        private static void WriteToPSTempFiles(string cmd)
        {
            //writes the PS file
            System.IO.File.AppendAllText(Settings.Default.ProgramWorkingDir, cmd + Environment.NewLine);
        }

        private static void Log(string line)
        {
            string logstr = DateTime.Now.ToString() + $" {line}";
            System.IO.File.AppendAllText(Settings.Default.ProgramLogging, logstr + Environment.NewLine);

        }

        private static string ConvertToDestinationRelativeURL(string filePath)
        {
            string s1 = Settings.Default.SourceSPOSiteURL;
            string s2 = Settings.Default.DestinationSPOSiteURL;

            Uri u1 = new Uri(s2);

            if (s1.EndsWith("/"))
            {
                s1 = s1.Remove(s1.Length - 1);
            }

            if (s2.EndsWith("/"))
            {
                s2 = s1.Remove(s1.Length - 1);
            }

            string oldsite = s1.Split('/')[4];
            string newsite = s2.Split('/')[4];

            return filePath.Replace(oldsite, newsite).Replace(Settings.Default.DocLibDisplayName, Settings.Default.DestinationDocLibDisplayName);


        }

        private static void GetAllItemInBatch()
        {
            //this func gets ALL the item in batch (configurable) from Source SPO list/library

            InitializeFile();
            bool ColorToggle = true;

            var siteUrl = Settings.Default.SourceSPOSiteURL;
            var doclib = Settings.Default.DocLibDisplayName;
            var uname = Settings.Default.Office365Username;
            var pss = ConvertToSecureString(Settings.Default.Office365Password);
            var exportFileNameWithPath = Settings.Default.MetadataFileExportLocation;
            var SPOQueryBatchSize = Settings.Default.SPOQueryBatchSize;

            Console.WriteLine($"Connecting to Source SPO Site {siteUrl} as user {uname} {Environment.NewLine}");

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
                    "</ViewFields><RowLimit>" + SPOQueryBatchSize + "</RowLimit></View>";

                ListItemCollection collListItem = oList.GetItems(camlQuery);

                clientContext.Load(collListItem);


                clientContext.ExecuteQuery();

                itemPosition = collListItem.ListItemCollectionPosition;

                foreach (ListItem oListItem in collListItem)
                {
                    string filename = oListItem["FileRef"].ToString();

                    string created = oListItem["Created"].ToString();

                    FieldUserValue user = (FieldUserValue)oListItem["Author"];
                    //string author = user.LookupValue;
                    string author = user.Email;

                    string Modified = oListItem["Modified"].ToString();
                    user = null;

                    user = (FieldUserValue)oListItem["Editor"];
                    //string Editor = user.LookupValue;
                    string Editor = user.Email;


                    string s = ($"Found - {oListItem.Id},{filename},{created},{author},{Modified},{Editor}");
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
