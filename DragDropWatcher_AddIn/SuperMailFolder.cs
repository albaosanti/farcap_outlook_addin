﻿using System;
using System.Collections.Generic;
using System.Linq;
using System.Reflection;
using Microsoft.Office.Interop.Outlook;

namespace DragDrapWatcher_AddIn
{
  public class SuperMailFolder
  {
    Folder _wrappedFolder;
    string _profileName;
    public List<SuperMailFolder> wrappedSubFolders = new List<SuperMailFolder>();

    internal SuperMailFolder(Folder folder, string profileName)
    {
      try
      {
        //assign it to local private master
        _wrappedFolder = folder;
        _profileName = profileName;
        //assign event handlers for the folder
        _wrappedFolder.Items.ItemAdd += Items_ItemAdd;
        _wrappedFolder.BeforeItemMove += Before_ItemMoveListener;
        //_wrappedFolder.Items.ItemChange += new Outlook.ItemsEvents_ItemChangeEventHandler(Items_ItemChange);

        //Go through all the subfolders and wrap them as well
        foreach (Folder tmpFolder in _wrappedFolder.Folders)
        {
          if (folder.Name.StartsWith("Vault", StringComparison.InvariantCultureIgnoreCase) ||
              folder.Name.StartsWith("Public Folder", StringComparison.InvariantCultureIgnoreCase))
            continue;

          if (folder.DefaultItemType != OlItemType.olMailItem) continue;

          var tmpWrapFolder = new SuperMailFolder(tmpFolder, _profileName);
          wrappedSubFolders.Add(tmpWrapFolder);
          wrappedSubFolders.AddRange(tmpWrapFolder.wrappedSubFolders);
        }
      }
      catch (System.Exception ex)
      {
        Globals.ThisAddIn.Error_Sender.SendNotification(ex.Message + ex.StackTrace);
      }
    }

    private void Before_ItemMoveListener(object Item, MAPIFolder TargetFolder, ref bool Cancel)
    {
      string src_ruleprefix = string.Empty;
      string target_ruleprefix = string.Empty;
      string rule_prefix = string.Empty;
      string folder_prefix = string.Empty;
      string sender_address = string.Empty;

      bool ok_added = false;
      bool ok_removed = false;

      MailItem oMsg = null;
      Folder src_folder = null;

      try
      {
        Globals.ThisAddIn.Error_Sender.WriteLog($"{MethodBase.GetCurrentMethod().Name} :: Drag and drop triggered!");
        if (Item is MailItem && TargetFolder != null)
        {
          Globals.ThisAddIn.Error_Sender.WriteLog($"{MethodBase.GetCurrentMethod().Name} :: Drag and drop happening for folder : {TargetFolder.Name}");
          if (!TargetFolder.Name.StartsWith("deleted items", StringComparison.OrdinalIgnoreCase))
          {
            oMsg = (MailItem)Item;
            src_folder = (Folder)oMsg.Parent;
            rule_prefix = Properties.Settings.Default.RuleName_Prefix.Trim();
            folder_prefix = Properties.Settings.Default.WatchFolder_Prefix.Trim();
            sender_address = Globals.ThisAddIn.fnGetSenderAddress(oMsg.Sender);

            if (string.IsNullOrWhiteSpace(sender_address))
            {
              Globals.ThisAddIn.Error_Sender.WriteLog($"{MethodBase.GetCurrentMethod().Name} :: Sender Address is Null or empty!");
              return;
            }

            //SOURCE FOLDER -> REMOVE FROM RULE
            if (src_folder.Name.StartsWith(folder_prefix, StringComparison.OrdinalIgnoreCase))
            {
              src_ruleprefix = rule_prefix + src_folder.Name;
              var to_remove = Globals.ThisAddIn.OutlookRules.FarCapRuleSenders.Where(
                row => row.sender_email.Equals(sender_address, StringComparison.OrdinalIgnoreCase) &&
                       row.rulename.StartsWith(src_ruleprefix, StringComparison.OrdinalIgnoreCase)
              ).ToList();

              foreach (var row in to_remove)
              {
                Globals.ThisAddIn.Error_Sender.WriteLog($"{MethodBase.GetCurrentMethod().Name} :: Removing {row.sender_email} from {row.rulename} !");
                if (Globals.ThisAddIn.OutlookRules.RemoveEmailFromRule(row.rulename, row.sender_email))
                  ok_removed = true;
              }
            }
            else
            {
              Globals.ThisAddIn.Error_Sender.WriteLog($"{MethodBase.GetCurrentMethod().Name} :: Skip scanning source folder as it didn't start with #");
            }

            //DESTINATION FOLDER -> ADD TO RULE
            if (TargetFolder.Name.StartsWith(folder_prefix, StringComparison.OrdinalIgnoreCase))
            {
              target_ruleprefix = rule_prefix + TargetFolder.Name;
              Globals.ThisAddIn.Error_Sender.WriteLog($"{MethodBase.GetCurrentMethod().Name} :: Adding {sender_address} to {target_ruleprefix} on {TargetFolder.Name}!");
              ok_added = Globals.ThisAddIn.OutlookRules.AddEmailToRule(target_ruleprefix, sender_address, oMsg.SenderName, TargetFolder);
            }
            else
            {
              Globals.ThisAddIn.Error_Sender.WriteLog($"{MethodBase.GetCurrentMethod().Name} :: Skip scanning target folder as it didn't start with #");
            }

            Globals.ThisAddIn.Error_Sender.WriteLog($"{MethodBase.GetCurrentMethod().Name} :: Drag and drop for {sender_address}, ok_added : {ok_added}, ok_removed : {ok_removed}");

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

    void Items_ItemAdd(object Item)
    {
      try
      {
        if (Item is Folder item)
        {
          SuperMailFolder tmpWrapFolder = new SuperMailFolder(item, _profileName);
          wrappedSubFolders.Add(tmpWrapFolder);
          wrappedSubFolders.AddRange(tmpWrapFolder.wrappedSubFolders);
        }
      }
      catch (System.Exception ex)
      {
        Globals.ThisAddIn.Error_Sender.SendNotification(ex.Message + ex.StackTrace);
      }
    }
  }
}