using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Stage3_FileAttributeWriteToDestination
{
    class Program
    {
        static void Main(string[] args)
        {
        }
        private static void UpdateFilePropertyTEMP()
        {

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
                    var temp = item.Split(',');
                    var ID = temp[0];
                    var FilePath = ConvertToDestinationRelativeURL(temp[1]);
                    var CreatedON = temp[2];
                    var CreatedBy = temp[3];
                    var ModifiedON = temp[4];
                    var ModifiedBy = temp[5];

                    //Console.WriteLine(FilePath);
                    try
                    {
                        Console.WriteLine($"Source File: {FilePath}");
                        var File1 = clientContext.Web.GetFileByServerRelativeUrl(FilePath);
                        clientContext.Load(File1, i => i.Title);
                        clientContext.ExecuteQuery();

                        Console.WriteLine($"Destination File: {File1.Title}");

                    }
                    catch (Exception e)
                    {

                        Console.WriteLine(e.Message);
                    }

                }

            }


        }
    }
}
