//using System;
//using System.Collections.Generic;
//using System.Linq;
//using System.Text;

using System.Diagnostics;
using System.IO;
using System.Net.Mime;
using Outlook = Microsoft.Office.Interop.Outlook;


namespace OutlookControl
{
    public class emailsend
    {
        public void ComposeEmail(string strTo, string strCC, string strSubject, string strHtmlBody, string strFolderEntryID, string strFolderStoreID)
        {
            Microsoft.Office.Interop.Outlook.Application app = new Microsoft.Office.Interop.Outlook.Application();

            Outlook.Folder DestFolder = app.Session.GetFolderFromID(strFolderEntryID, strFolderStoreID) as Outlook.Folder;

            Outlook.MailItem item = app.CreateItem(Outlook.OlItemType.olMailItem);

            item.HTMLBody = strHtmlBody;
            item.To = strTo;
            item.CC = strCC;
            item.Subject = strSubject;

            item.Display();
            item.Save();
            item.Move(DestFolder);

        }

        public static string SaveTopEmail(string strMailbox, string strFolder1, string strFolder2, string strFileTarget)
        {
            Microsoft.Office.Interop.Outlook.Application app = new Microsoft.Office.Interop.Outlook.Application();
            int intFolderCount = 0;
            int intItemsCount = 0;
            string strFolderFound = "N";
            Outlook.MAPIFolder fldrsource = app.GetNamespace("MAPI").Folders[0];
            string strEntryID = "";
            string strStoreId = "";

            if (strMailbox.ToString().Trim() != "")
            {
                foreach (Outlook.MAPIFolder fldrmlbox in app.GetNamespace("MAPI").Folders)
                {
                    if (fldrmlbox.Name.ToString().ToLower().Trim() == strMailbox.ToString().ToLower().Trim())
                    {
                        
                        if (strFolder1.ToString().Trim() != "")
                        {
                            intFolderCount = fldrmlbox.Folders.Count;
                            if (intFolderCount > 0)
                            { 
                                foreach (Outlook.MAPIFolder fldrlvl1 in fldrmlbox.Folders)
                                {
                                    if (fldrlvl1.Name.ToString().ToLower().Trim() == strFolder1.ToString().ToLower().Trim())
                                    {

                                        if (strFolder2.ToString().Trim() != "")
                                        {
                                            intFolderCount = fldrlvl1.Folders.Count;
                                            if (intFolderCount > 0)
                                            { 
                                                foreach (Outlook.MAPIFolder fldrlvl2 in fldrmlbox.Folders)
                                                {
                                                    if (fldrlvl2.Name.ToString().ToLower().Trim() == strFolder2.ToString().ToLower().Trim())
                                                    {
                                                        fldrsource = fldrlvl2;
                                                        strFolderFound = "Y";
                                                        break;
                                                    }
                                                }
                                            }
                                        }
                                        else
                                        {
                                            fldrsource = fldrlvl1;
                                            strFolderFound = "Y";
                                        }
                                        break;
                                    }
                                }
                            }
                        }
                        else
                        {
                            fldrsource = fldrmlbox;
                            strFolderFound = "Y";
                        }
                        
                        break;
                    }
                }
            }
            
            if (strFolderFound == "Y")
            {
                intItemsCount = fldrsource.Items.Count;
                
                if (intItemsCount > 0)   
                {
                    for (int i = 0; i <= intItemsCount; i++)
                    {
                        if (fldrsource.Items[i] is Outlook.MailItem)
                        { 
                            fldrsource.Items[i].SaveAs(strFileTarget, Outlook.OlSaveAsType.olMSG);
                            strEntryID = fldrsource.Items[i].EntryID;
                            strStoreId = fldrsource.Items[i].Parent.StoreID;
                            break;
                        }
                    }
                }
            }
            
            return "" + strEntryID + "|xxx|" + strStoreId;
        }
    }
}