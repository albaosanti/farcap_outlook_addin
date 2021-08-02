using System;
using System.Collections.Generic;
using System.Linq;

namespace DragDrapWatcher_AddIn
{
  class SuperMailFolder
  {
    #region private variables
    Microsoft.Office.Interop.Outlook.Folder _wrappedFolder;
    string _profileName;
    public List<SuperMailFolder> wrappedSubFolders = new List<SuperMailFolder>();
    string folderName = string.Empty;
    #endregion

    #region constructor
    internal SuperMailFolder(Microsoft.Office.Interop.Outlook.Folder folder, string profileName)
    {
      try
      {
        //assign it to local private master
        _wrappedFolder = folder;
        folderName = folder.Name;
        _profileName = profileName;
        //assign event handlers for the folder
        _wrappedFolder.Items.ItemAdd += Items_ItemAdd;
        _wrappedFolder.BeforeItemMove += Before_ItemMoveListener;
        //_wrappedFolder.Items.ItemChange += new Outlook.ItemsEvents_ItemChangeEventHandler(Items_ItemChange);

        //Go through all the subfolders and wrap them as well
        foreach (Microsoft.Office.Interop.Outlook.Folder tmpFolder in _wrappedFolder.Folders)
        {
          if (folder.Name.StartsWith("Vault", StringComparison.InvariantCultureIgnoreCase) ||
              folder.Name.StartsWith("Public Folder", StringComparison.InvariantCultureIgnoreCase))
            continue;

          if (folder.DefaultItemType == Microsoft.Office.Interop.Outlook.OlItemType.olMailItem)
          {
            var tmpWrapFolder = new SuperMailFolder(tmpFolder, _profileName);
            wrappedSubFolders.Add(tmpWrapFolder);
            wrappedSubFolders.AddRange(tmpWrapFolder.wrappedSubFolders);
          }
        }
      }
      catch (System.Exception ex)
      { Globals.ThisAddIn.Error_Sender.SendNotification(ex.Message + ex.StackTrace); }
    }
    #endregion

    private void Before_ItemMoveListener(object Item, Microsoft.Office.Interop.Outlook.MAPIFolder TargetFolder, ref bool Cancel)
    {
      string src_ruleprefix = string.Empty;
      string target_ruleprefix = string.Empty;
      string rule_prefix = string.Empty;
      string folder_prefix = string.Empty;
      string sender_address = string.Empty;

      bool ok_added = false;
      bool ok_removed = false;

      Microsoft.Office.Interop.Outlook.MailItem oMsg = null;
      Microsoft.Office.Interop.Outlook.Folder src_folder = null;

      try
      {
        if (Item is Microsoft.Office.Interop.Outlook.MailItem && TargetFolder != null)
        {
          if (!TargetFolder.Name.StartsWith("deleted items",StringComparison.OrdinalIgnoreCase))
          {
            oMsg = (Microsoft.Office.Interop.Outlook.MailItem)Item;
            src_folder = (Microsoft.Office.Interop.Outlook.Folder)oMsg.Parent;
            rule_prefix = Properties.Settings.Default.RuleName_Prefix.Trim();
            folder_prefix = Properties.Settings.Default.WatchFolder_Prefix.Trim();
            sender_address = Globals.ThisAddIn.fnGetSenderAddress(oMsg.Sender);

            if (string.IsNullOrWhiteSpace(sender_address))
              return;

            //SOURCE FOLDER -> REMOVE FROM RULE
            if (src_folder.Name.StartsWith(folder_prefix, StringComparison.OrdinalIgnoreCase))
            {
              src_ruleprefix = rule_prefix + src_folder.Name;
              var to_remove = Globals.ThisAddIn.OutlookRules.FarCapRuleSenders.Where(
                row => row.sender_email.Equals(sender_address, StringComparison.OrdinalIgnoreCase) && 
                       row.rulename.StartsWith(src_ruleprefix, StringComparison.OrdinalIgnoreCase)
              ).ToList();

              foreach(var row in to_remove)
              {
                if(Globals.ThisAddIn.OutlookRules.RemoveEmailFromRule(row.rulename, row.sender_email))
                  ok_removed = true;
              }
            }

            //DESTINATION FOLDER -> ADD TO RULE
            if (TargetFolder.Name.StartsWith(folder_prefix,StringComparison.OrdinalIgnoreCase))
            {
              target_ruleprefix = rule_prefix + TargetFolder.Name;
              ok_added = Globals.ThisAddIn.OutlookRules.AddEmailToRule(target_ruleprefix, sender_address, oMsg.SenderName, TargetFolder);
            }

            //Save rules
            if (Globals.ThisAddIn.OutlookRules != null && (ok_added || ok_removed))
              Globals.ThisAddIn.OutlookRules.Save(true);
          }
        }
      }
      catch (System.Exception ex)
      {
        Globals.ThisAddIn.Error_Sender.SendNotification(ex.Message + ex.StackTrace);
      }
    }

    void Items_ItemRemove()
    {

    }

    void Items_ItemChange(object item)
    {
     
    }

    #region Handler of addition item into a folder
    void Items_ItemAdd(object Item)
    {
      try
      {
        if (Item is Microsoft.Office.Interop.Outlook.Folder)
        {
          SuperMailFolder tmpWrapFolder = new SuperMailFolder((Microsoft.Office.Interop.Outlook.Folder)Item, _profileName);
          wrappedSubFolders.Add(tmpWrapFolder);
          wrappedSubFolders.AddRange(tmpWrapFolder.wrappedSubFolders);
        }
      }
      catch (System.Exception ex)
      {
        Globals.ThisAddIn.Error_Sender.SendNotification(ex.Message + ex.StackTrace);
      }
    }

    #endregion
  }
}