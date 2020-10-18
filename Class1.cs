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

        public void SendEmail(string strTo, string strCC, string strSubject, string strHtmlBody)
        {
            Microsoft.Office.Interop.Outlook.Application app = new Microsoft.Office.Interop.Outlook.Application();

            Outlook.MailItem item = app.CreateItem(Outlook.OlItemType.olMailItem);

            item.HTMLBody = strHtmlBody;
            item.To = strTo;
            item.CC = strCC;
            item.Subject = strSubject;

            item.Recipients.ResolveAll();
            item.Send();

        }

        public static string SaveTopEmail(string strMailbox, string strFolder1, string strFolder2, string strFolder3, string strFileTarget)
        {
            Microsoft.Office.Interop.Outlook.Application app = new Microsoft.Office.Interop.Outlook.Application();
            int intFolderCount = 0;
            int intItemsCount = 0;
            string strFolderFound = "N";
            Outlook.MAPIFolder fldrsource = app.GetNamespace("MAPI").GetDefaultFolder(Outlook.OlDefaultFolders.olFolder‌​Inbox).Parent;
            string strEntryID = "";
            string strStoreId = "";
            string strLevel = "";

            if (strMailbox.ToString().Trim() != "")
            {
                strLevel = "1";
                foreach (Outlook.MAPIFolder fldrmlbox in app.GetNamespace("MAPI").Folders)
                {
                    if (fldrmlbox.Name.ToString().ToLower().Trim() == strMailbox.ToString().ToLower().Trim())
                    {
                        strLevel = "2";
                        if (strFolder1.ToString().Trim() != "")
                        {
                            intFolderCount = fldrmlbox.Folders.Count;
                            if (intFolderCount > 0)
                            {
                                strLevel = "3";
                                foreach (Outlook.MAPIFolder fldrlvl1 in fldrmlbox.Folders)
                                {
                                    if (fldrlvl1.Name.ToString().ToLower().Trim() == strFolder1.ToString().ToLower().Trim())
                                    {
                                        strLevel = "4";
                                        if (strFolder2.ToString().Trim() != "")
                                        {
                                            strLevel = "5";
                                            intFolderCount = fldrlvl1.Folders.Count;
                                            if (intFolderCount > 0)
                                            {
                                                strLevel = "6";
                                                foreach (Outlook.MAPIFolder fldrlvl2 in fldrlvl1.Folders)
                                                {
                                                    strLevel = "7" + fldrlvl2.Name.ToString().ToLower().Trim();
                                                    if (fldrlvl2.Name.ToString().ToLower().Trim() == strFolder2.ToString().ToLower().Trim())
                                                    {
                                                        strLevel = "8";
                                                        if (strFolder3.ToString().Trim() != "")
                                                        {
                                                            intFolderCount = fldrlvl2.Folders.Count;
                                                            if (intFolderCount > 0)
                                                            {
                                                                foreach (Outlook.MAPIFolder fldrlvl3 in fldrlvl1.Folders)
                                                                {
                                                                    if (fldrlvl3.Name.ToString().ToLower().Trim() == strFolder3.ToString().ToLower().Trim())
                                                                    {
                                                                        fldrsource = fldrlvl3;
                                                                        strFolderFound = "Y";
                                                                        break;
                                                                    }
                                                                }
                                                            }
                                                        }
                                                        else
                                                        {
                                                            fldrsource = fldrlvl2;
                                                        strFolderFound = "Y";
                                                        }
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

            strLevel = "9";
            if (strFolderFound == "Y")
            {
                intItemsCount = fldrsource.Items.Count;

                if (intItemsCount > 0)   
                {
                    
                    foreach (Outlook.MailItem themailitem in fldrsource.Items)
                    {
                        themailitem.SaveAs(strFileTarget, Outlook.OlSaveAsType.olMSG);
                        strEntryID = themailitem.EntryID;
                        strStoreId = themailitem.Parent.StoreID;
                        break;
                    }
                }
            }

            return "" + strFolderFound + "|xxx|" + strEntryID + "|xxx|" + strStoreId + "" ;
        }

        public static string ValidateFolder(string strMailbox, string strFolder1, string strFolder2, string strFolder3)
        {
            Microsoft.Office.Interop.Outlook.Application app = new Microsoft.Office.Interop.Outlook.Application();
            int intFolderCount = 0;
            string strFolderFound = "N";
            Outlook.MAPIFolder fldrsource = app.GetNamespace("MAPI").GetDefaultFolder(Outlook.OlDefaultFolders.olFolder‌​Inbox).Parent;
            string strLevel = "";

            if (strMailbox.ToString().Trim() != "")
            {
                strLevel = "1";
                foreach (Outlook.MAPIFolder fldrmlbox in app.GetNamespace("MAPI").Folders)
                {
                    if (fldrmlbox.Name.ToString().ToLower().Trim() == strMailbox.ToString().ToLower().Trim())
                    {
                        strLevel = "2";
                        if (strFolder1.ToString().Trim() != "")
                        {
                            intFolderCount = fldrmlbox.Folders.Count;
                            if (intFolderCount > 0)
                            {
                                strLevel = "3";
                                foreach (Outlook.MAPIFolder fldrlvl1 in fldrmlbox.Folders)
                                {
                                    if (fldrlvl1.Name.ToString().ToLower().Trim() == strFolder1.ToString().ToLower().Trim())
                                    {
                                        strLevel = "4";
                                        if (strFolder2.ToString().Trim() != "")
                                        {
                                            strLevel = "5";
                                            intFolderCount = fldrlvl1.Folders.Count;
                                            if (intFolderCount > 0)
                                            {
                                                strLevel = "6";
                                                foreach (Outlook.MAPIFolder fldrlvl2 in fldrlvl1.Folders)
                                                {
                                                    strLevel = "7" + fldrlvl2.Name.ToString().ToLower().Trim();
                                                    if (fldrlvl2.Name.ToString().ToLower().Trim() == strFolder2.ToString().ToLower().Trim())
                                                    {
                                                        strLevel = "8";
                                                        if (strFolder3.ToString().Trim() != "")
                                                        {
                                                            intFolderCount = fldrlvl2.Folders.Count;
                                                            if (intFolderCount > 0)
                                                            {
                                                                foreach (Outlook.MAPIFolder fldrlvl3 in fldrlvl1.Folders)
                                                                {
                                                                    if (fldrlvl3.Name.ToString().ToLower().Trim() == strFolder3.ToString().ToLower().Trim())
                                                                    {
                                                                        fldrsource = fldrlvl3;
                                                                        strFolderFound = "Y";
                                                                        break;
                                                                    }
                                                                }
                                                            }
                                                        }
                                                        else
                                                        {
                                                            fldrsource = fldrlvl2;
                                                            strFolderFound = "Y";
                                                        }
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


            return "" + strFolderFound + "";
        }

        public static string CountAllEmails(string strMailbox, string strFolder1, string strFolder2, string strFolder3)
        {
            Microsoft.Office.Interop.Outlook.Application app = new Microsoft.Office.Interop.Outlook.Application();
            int intFolderCount = 0;
            int intItemsCount = 0;
            string strFolderFound = "N";
            Outlook.MAPIFolder fldrsource = app.GetNamespace("MAPI").GetDefaultFolder(Outlook.OlDefaultFolders.olFolder‌​Inbox).Parent;
            string strLevel = "";

            if (strMailbox.ToString().Trim() != "")
            {
                strLevel = "1";
                foreach (Outlook.MAPIFolder fldrmlbox in app.GetNamespace("MAPI").Folders)
                {
                    if (fldrmlbox.Name.ToString().ToLower().Trim() == strMailbox.ToString().ToLower().Trim())
                    {
                        strLevel = "2";
                        if (strFolder1.ToString().Trim() != "")
                        {
                            intFolderCount = fldrmlbox.Folders.Count;
                            if (intFolderCount > 0)
                            {
                                strLevel = "3";
                                foreach (Outlook.MAPIFolder fldrlvl1 in fldrmlbox.Folders)
                                {
                                    if (fldrlvl1.Name.ToString().ToLower().Trim() == strFolder1.ToString().ToLower().Trim())
                                    {
                                        strLevel = "4";
                                        if (strFolder2.ToString().Trim() != "")
                                        {
                                            strLevel = "5";
                                            intFolderCount = fldrlvl1.Folders.Count;
                                            if (intFolderCount > 0)
                                            {
                                                strLevel = "6";
                                                foreach (Outlook.MAPIFolder fldrlvl2 in fldrlvl1.Folders)
                                                {
                                                    strLevel = "7" + fldrlvl2.Name.ToString().ToLower().Trim();
                                                    if (fldrlvl2.Name.ToString().ToLower().Trim() == strFolder2.ToString().ToLower().Trim())
                                                    {
                                                        strLevel = "8";
                                                        if (strFolder3.ToString().Trim() != "")
                                                        {
                                                            intFolderCount = fldrlvl2.Folders.Count;
                                                            if (intFolderCount > 0)
                                                            {
                                                                foreach (Outlook.MAPIFolder fldrlvl3 in fldrlvl1.Folders)
                                                                {
                                                                    if (fldrlvl3.Name.ToString().ToLower().Trim() == strFolder3.ToString().ToLower().Trim())
                                                                    {
                                                                        fldrsource = fldrlvl3;
                                                                        strFolderFound = "Y";
                                                                        break;
                                                                    }
                                                                }
                                                            }
                                                        }
                                                        else
                                                        {
                                                            fldrsource = fldrlvl2;
                                                            strFolderFound = "Y";
                                                        }
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

            strLevel = "9";
            if (strFolderFound == "Y")
            {
                intItemsCount = fldrsource.Items.Count;
            }

            return "" + strFolderFound + "|xxx|" + intItemsCount.ToString()  + "";
        }


    }
}
