using System;
using System.Collections.Generic;
using System.IO;
using Outlook = Microsoft.Office.Interop.Outlook;
using System.Configuration;

namespace EAR
{
    class Program
    {
        static string basePath = ConfigurationManager.AppSettings["fileTransferPath"];
        
        static void Main(string[] args)
        {
            
            EnumerateAccounts();
        }
        
        static void EnumerateAccounts()//checks the number of accounts configured. In our case, there is only one account configured.
        {
            Outlook.Account primaryAccount = null;
            Outlook.Application Application = new Outlook.Application();
            Outlook.Accounts accounts = Application.Session.Accounts;
            foreach (Outlook.Account account in accounts)
            {
                primaryAccount = account;
                break;
            }
            /* foreach (Outlook.Account account in accounts)     //this loop must be used if there are more than one accounts configured in the system. Replace "Gaurav" with the a word that is contained in the email name.
             {
                 if (account.DisplayName.Contains("Gaurav"))
                  {
                 primaryAccount = account;
                 break;
                 }
             }*/
            Outlook.Folder selectedFolder = Application.Session.DefaultStore.GetRootFolder() as Outlook.Folder;
            selectedFolder = getFolder(@"\\" + primaryAccount.DisplayName);  //Fetches the inbox folder
            enumerateFolders(selectedFolder);  //Iterates amongst the folders and selects the Inbox folder to retreive the attachment from the inbox
        }
        
        static Outlook.Folder getFolder(string folderPath)  //Gets the Inbox folder and confines the program search only within it.
        {
            Outlook.Folder folder;
            string backslash = @"\";
            try
            {
                if (folderPath.StartsWith(@"\\"))   //The folder path is the name of the folder configured in outlook, which is the userName. Uncomment the blow two lines and run to understand what is the folderPath here.
                {
                    folderPath = folderPath.Remove(0, 2);
                }
                String[] folders = folderPath.Split(backslash.ToCharArray()); //folders[] array contains the userNames configured in outlook as folders.
                Outlook.Application Application = new Outlook.Application(); 
                folder = Application.Session.Folders[folders[0]] as Outlook.Folder; 
                Outlook.Folders subfolders = folder.Folders;
                for (int i = 1; i <= subfolders.Count; i++)
                {
                    if (subfolders[i].Name.Contains("Inbox"))
                    {
                        folder = (Outlook.Folder)subfolders[i];
                    }
                }
                return folder;   //Returns the inbox folder
            }
            catch (Exception e)
            {
                generalExceptions(e);
                return null;
            }
        }
        
        static void enumerateFolders(Outlook.Folder folder)  //Checks if there are sub folders inside the Inbox folder.
        {
            Outlook.Folders subfolders = folder.Folders;
            if (subfolders.Count > 0)
            {
                for (int i = 0; i < subfolders.Count; i++)
                {
                    Outlook.Folder subFolder = (Outlook.Folder) subfolders[i]; 
                    iterateMessages(subFolder); //This searches for the attachment in every subfolder inside the inbox folder. If there are any subfolders.
                }
            }
            else
            {
                iterateMessages(folder);     //This implements the core functionality of the program. It iterates amongst the emails to retrieve the clearstream attachment.
            }
        }
        
        static void iterateMessages(Outlook.Folder folder)  //Core function of the program. Checks the mails for the subject name given in the config file and fetches the attachment from it.
        {
            var fi = folder.Items;
            Outlook.MailItem mi = null;
            var today = DateTime.Today.ToString("ddMMyyyy");
            var yesterDay = DateTime.Today.AddDays(-1).ToString("ddMMyyy");
            var fileName = ConfigurationManager.AppSettings["fileName"];  //Gets the static part of the file name from the config file. The date part is dynamic and is added to this later using the today variable from the above declaration.
            var fileCreationCurrentDay = false;
            var fileCreationPreviousDay = false;
            string ex = "File not received"; 
            string res = "File received and copied to path";
            var attachment = "";
            var currentDayAttachment = fileName.ToUpper().Trim()+"_"+today;
            var previousDayAttachment = fileName.ToUpper().Trim() +"_"+ yesterDay;
            var subject = "";
            if (fi != null)
            {
                try
                {
                    foreach (dynamic item in fi)     //iterates amongst all the attachments in the inbox untill a match is found.
                    {
                        try
                        {
                            mi = item;   //Some attachments are of the type Outlook.MailItem
                            subject = mi.Subject.ToString();
                            try
                            {
                                mi = (Outlook.MailItem)item; //Some need to be implicitly converted
                                subject = mi.Subject.ToString();
                            }
                            catch (Exception e)
                            {
                                generalExceptions(e);
                            }
                        }
                        catch (Exception e)
                        {
                            generalExceptions(e);
                        }
                        finally
                        {
                            var attachments = mi.Attachments;
                            try
                            {
                                if (subject.Contains(fileName + today))
                                {
                                    if (attachments.Count != 0)
                                    {
                                        res = "File recieved and copied to path";
                                        for (int j = 1; j <= attachments.Count; j++)
                                        {
                                            attachment = attachments[j].FileName;
                                            attachment = attachment.Substring(0, attachment.IndexOf(".")); //This will remove the file type (.txt) from the file name so that it can be added later after appending the date to the file name. Ref: [1] mentioned in the below comment
                                            if (!Directory.Exists(basePath))
                                            {
                                                Directory.CreateDirectory(basePath);
                                            }
                                            if (!File.Exists(basePath + attachment + today))
                                            {
                                                attachments[j].SaveAsFile(basePath + attachment + "_" + today + ".txt");//[1]
                                                fileCreationCurrentDay = true; //This flag is used later for logging. If the flag is false, and if the file is not found in the remote path, errorlogging is performed (refer to the if conditions used for logging below)
                                                successLogging(res, true, today);
                                                break;
                                            }
                                        }
                                    }
                                }

                                if (subject.Contains(fileName + yesterDay)) //This loop checks if the file is missing for the previous day.
                                {
                                    if (attachments.Count != 0)
                                    {
                                        res = "File received and copied to path for ";
                                        for (int j = 1; j <= attachments.Count; j++)
                                        {
                                            attachment = attachments[j].FileName;
                                            attachment = attachment.Substring(0, attachment.IndexOf(".")); //This will remove the file type (.txt) from the file name so that it can be added later after appending the date to the file name. Ref: [1] mentioned in the below comment
                                            if (!Directory.Exists(basePath))
                                            {
                                                Directory.CreateDirectory(basePath);
                                            }
                                            if (!File.Exists(basePath + attachment + yesterDay))
                                            {
                                                res = res + yesterDay;
                                                attachments[j].SaveAsFile(basePath + attachment + "_" + yesterDay + ".txt"); //[1]
                                                fileCreationPreviousDay = true; //This flag is used later for logging. If the flag is false, and if the file is not found in the remote path, errorlogging is performed (refer to the if conditions used for logging below)
                                                successLogging(res, true, yesterDay);
                                                break;
                                            }
                                        }
                                    }
                                }
                            }
                            catch (Exception e)
                            {
                                generalExceptions(e);
                            }
                        }
                    }
                    if (!File.Exists(basePath + currentDayAttachment) && (fileCreationCurrentDay != true))  //This condition is executed if there file is not found for the current day in the remote path and not even received in the mail.
                    {
                        errorLogging(today, ex);
                    }
                    if (!File.Exists(basePath + previousDayAttachment) && (fileCreationPreviousDay != true)) //The same as the above condition, but checks for the previous day
                    {
                        errorLogging(yesterDay, ex);
                    }
                    if (File.Exists(basePath + currentDayAttachment) && fileCreationCurrentDay != true) // This condition is executed if the file transfer for the current day is successfull
                    {
                        successLogging("", true,today);
                    }
                    if (File.Exists(basePath + previousDayAttachment) && fileCreationPreviousDay != true) //This condition is executed if the file transfer for the previous day is successfull 
                    {
                        successLogging("", true, yesterDay);
                    }
                }
  
                catch (Exception e)     //This is executed if there is an error while running and is logged appropriately
                {
                    generalExceptions(e);
                }
            }
        }
        
        static void errorLogging(string fileDate, string ex)  //This initializes Appends a log file whenever the attachment is not yet received
        {
            string loggingPath = ConfigurationManager.AppSettings["fileCopyErrorLoggingPath"];
            var today = DateTime.Today.ToString("ddMMyyyy");
            var yesterday = DateTime.Today.AddDays(-1).ToString("ddMMyyyy");
            string logFolderPath = "";
            if (fileDate.Equals(today))    //This condition is execute if the name of the log file to be created is for today
            {
                logFolderPath = @loggingPath + "LOG_" + today + ".txt";
            }
            else if (fileDate.Equals(yesterday))  //This condition is executed if the name of the log file to be created is for the previous day.
            {
                logFolderPath = @loggingPath + "LOG_" + yesterday + ".txt";
            }
            if (!File.Exists(logFolderPath))
            {
                File.Create(logFolderPath).Dispose();
            }
            using (StreamWriter sw = File.AppendText(logFolderPath))
            {
                sw.WriteLine(DateTime.Now + " : " + ex);
            }
        }
        
        static void successLogging(string res, Boolean logCreation,string day) //This logs the successfull file transfers in the path provided in the config file
        {
            string loggingPath = ConfigurationManager.AppSettings["successLoggingPath"];
            day = day.Replace(".", "");
            string logFolderPath = @loggingPath + "LOG_" + day + ".txt";
            string errorFolderPath = ConfigurationManager.AppSettings["fileCopyErrorLoggingPath"] + "LOG_" + day + ".txt";
            if (logCreation == true)
            {
                if (!File.Exists(logFolderPath))
                {
                    File.Create(logFolderPath).Dispose();
                    using (StreamWriter sw = File.AppendText(logFolderPath))
                    {
                        sw.WriteLine(DateTime.Now + " : " + res);
                    }
                }
            }
            if (File.Exists(errorFolderPath))
            {
                File.Delete(errorFolderPath);
            }
        }
        
        static void generalExceptions(Exception e) //Saves the run time errors in the specified folder path
        {
            string loggingPath = ConfigurationManager.AppSettings["generalExceptionLoggingPath"];
            var today = DateTime.Today.ToString("ddMMyyyy");
            string logFolderPath = @loggingPath + "LOG_" + today + ".txt";
            if (!File.Exists(logFolderPath))
            {
                File.Create(logFolderPath).Dispose();
            }
            using (StreamWriter sw = File.AppendText(logFolderPath))
            {
                sw.WriteLine(DateTime.Now + " : " + e);
            }
        }
    }
}
