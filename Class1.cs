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
    }
}
