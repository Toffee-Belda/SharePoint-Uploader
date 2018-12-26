using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Microsoft.SharePoint;
using System.Collections.Concurrent;
using System.Security;
using Microsoft.SharePoint.Client;
using System.Web;
using System.IO;
using System.Configuration;

namespace SharePointUploader
{
    class Program
    {
        static string fileToUpload;
        static string sharePointUrl;
        static string subFolderName;
        static string username;
        static string password;

        static void Main(string[] args)
        {
            Console.Clear();

            //retrieve the credentials from the configuration
            username = ConfigurationManager.AppSettings["Username"];
            password = ConfigurationManager.AppSettings["Password"];
            
            if (args.Count() == 0)
            {
                Console.WriteLine("======================================================");
                Console.WriteLine("SharePoint Uploader");
                Console.WriteLine("By: TOFFEE");
                Console.WriteLine("======================================================");

                Console.WriteLine("");

                Console.WriteLine("\nPlease type the SharePoint URL or press ENTER to use the default in the configuration: ");
                sharePointUrl = string.Format("{0}", Console.ReadLine());

                if(string.IsNullOrEmpty(sharePointUrl))
                {
                    sharePointUrl = ConfigurationManager.AppSettings["SharePointUrl"];
                    Console.WriteLine(sharePointUrl);
                }

                Console.WriteLine("\nSelect an option below:\n U for Upload file\n L for List a folder");
                var option = Console.ReadKey(true);

                // check if the selected operation is correct
                char[] options = { 'U', 'L' };
                if(!options.Contains(char.ToUpper(option.KeyChar)))
                {
                    Console.WriteLine("\nInvalid option. Press any key to exit.");
                    Console.Read();
                    return;
                }

                // validate fileToUpload and subFolderName
                if (option.KeyChar == char.ToLower('U'))
                {
                    for (var i = 0; i <= 2; i++)
                    {
                        if (string.IsNullOrEmpty(fileToUpload))
                        {
                            Console.WriteLine("\nType the UNC path of the file to upload: ");
                            fileToUpload = Console.ReadLine();
                            if (i == 2 && string.IsNullOrEmpty(fileToUpload))
                            {
                                Console.WriteLine("File not found.");
                            }
                        }
                    }

                   
                    for (var i = 0; i <= 2; i++)
                    {
                        if (string.IsNullOrEmpty(subFolderName))
                        {
                            Console.WriteLine("\nType the SharePoint sub-folder: ");
                            subFolderName = Console.ReadLine();
                            if (i == 2 && string.IsNullOrEmpty(subFolderName))
                            {
                                Console.WriteLine("\nInvalid sub-folder.");
                            }
                        }
                    }

                    //Process the file upload to a sharepoint folder.
                    if (!string.IsNullOrEmpty(fileToUpload) && !string.IsNullOrEmpty(sharePointUrl) && !string.IsNullOrEmpty(subFolderName))
                    {    
                        UploadFile(fileToUpload, sharePointUrl, subFolderName);
                    }
                    else
                    {
                        Console.WriteLine("\nInvalid parameters. Please press any key to exit.");
                        Console.Read();
                    }
                }

                if (option.KeyChar == char.ToLower('L'))
                {
                    ListFolders(sharePointUrl);
                }
            }
        }
        
        public static void UploadFile(string fileName, string url, string subFolderName)
        {
           try
            {
                Console.WriteLine("\nUpload in progress.....");
                using ( var context = new ClientContext(url))
                {
                    var securePassword = new SecureString();
                    foreach (var c in password) securePassword.AppendChar(c);
                    context.Credentials = new SharePointOnlineCredentials(username, securePassword);

                    var web = context.Web;
                    var newFile = new FileCreationInformation { Content = System.IO.File.ReadAllBytes(fileName), Url = Path.GetFileName(fileName) };
                    var docs = web.Lists.GetByTitle("Documents");

                    context.Load(docs.RootFolder.Folders);
                    context.ExecuteQuery();

                    docs.RootFolder.Folders.GetByUrl(subFolderName).Files.Add(newFile);
                    context.ExecuteQuery();
                    Console.WriteLine("\nUpload successful.");
                    Console.WriteLine("\nPress any key to exit.");
                    Console.Read();
                }
            }
            catch(Exception ex)
            {
                Console.WriteLine("\nUnexpected error has occured.\nDetails: {0}",ex.Message);
                Console.WriteLine("\nPlease press any key to exit.");
                Console.Read();
            }
        }

        private static void ListFolders(string url)
        {
            try
            {
                Console.WriteLine("\nFetching folders in progress.....");
                using (var context = new ClientContext(url))
                {
                    var securePassword = new SecureString();
                    foreach (var c in password) securePassword.AppendChar(c);
                    context.Credentials = new SharePointOnlineCredentials(username, securePassword);

                    var web = context.Web;
                    var list = web.Lists.GetByTitle("Documents");

                    context.Load(list.RootFolder.Folders);
                    context.ExecuteQuery();

                    //retreive folders and role assignments
                    foreach (Folder folder in list.RootFolder.Folders.OrderBy(f => f.Name))
                    {
                       Console.WriteLine(folder.Name);
                       try
                        {
                            context.Load(folder.Folders);
                            context.ExecuteQuery();

                            context.Load(folder.Files);
                            context.ExecuteQuery();

                            foreach (var f in folder.Folders.OrderBy(f => f.Name))
                            {
                                Console.WriteLine("  - [{0}]", f.Name);
                            }
                            foreach (var f in folder.Files.OrderBy(f => f.Name))
                            {
                                Console.WriteLine("  - {0}", f.Name);
                            }
                        }
                        catch
                        {
                            Console.WriteLine("  -");
                        }
                    }
                    Console.WriteLine("\nEnd of folder list");
                    Console.WriteLine("\nPress any key to exit.");
                    Console.Read();
                }
            }
            catch (Exception ex)
            {
                Console.WriteLine("\nUnexpected error has occured.\nDetails: {0}", ex.Message);
                Console.WriteLine("\nPlease press any key to exit.");
                Console.Read();
            }
        }
    }
}
