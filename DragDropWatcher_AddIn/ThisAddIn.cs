using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Linq;
using System.Text;
using System.Xml.Linq;
using Microsoft.VisualBasic;
using System.Windows.Forms;
using System.Runtime.InteropServices;
using System.IO;
using System.Text.RegularExpressions;

using Outlook = Microsoft.Office.Interop.Outlook;
using Office = Microsoft.Office.Core;

namespace DragDrapWatcher_AddIn
{
  public partial class ThisAddIn
  {
    #region Global Variables
    public string CAT_RULE_PREFIX = string.IsNullOrWhiteSpace(Properties.Settings.Default.CategoryRulePrefix) ? "#fcap_cat_" : Properties.Settings.Default.CategoryRulePrefix;
    public GlobalRules OutlookRules = null;
    public clsSendNotif Error_Sender = null;
    #endregion

    private void ThisAddIn_Startup(object sender, System.EventArgs e)
    {
      Outlook._NameSpace outNS = null;
      SuperMailFolder folderToWrap = null;

      try
      {

        Error_Sender = new clsSendNotif();
        Outlook.Application application = this.Application;
        //Get the MAPI namespace
        outNS = application.GetNamespace("MAPI");
        OutlookRules = new GlobalRules(application, this);
        //Get UserName
        string profileName = outNS.CurrentUser.Name;

        //DRAG & DROP WILL BE CREATED HERE
        Outlook.Folders folders = outNS.Folders;
        foreach (Outlook.Folder folder in folders)
        {
          var stopwatch = Stopwatch.StartNew();
          Error_Sender.WriteLog(string.Empty,
            $"Start Scanning folder :: FullFolderPath: {folder.FullFolderPath}, Name: {folder.Name}, IsSharePointFolder: {folder.IsSharePointFolder}, InAppFolderSyncObject: {folder.InAppFolderSyncObject}");
          if (folder.Name.StartsWith("Vault", StringComparison.InvariantCultureIgnoreCase) ||
              folder.Name.StartsWith("Public Folder", StringComparison.InvariantCultureIgnoreCase))
          {
            stopwatch.Stop();
            Error_Sender.WriteLog(string.Empty,
              $"Skip Scanning folder :: FullFolderPath: {folder.FullFolderPath}, Name: {folder.Name}, Time taken : {stopwatch.Elapsed.ToString()}");
            continue;
          }

          if (folder.DefaultItemType == Outlook.OlItemType.olMailItem)
            folderToWrap = new SuperMailFolder(folder, profileName);
          stopwatch.Stop();
          Error_Sender.WriteLog(string.Empty,
            $"End Scanning folder :: FullFolderPath: {folder.FullFolderPath}, Name: {folder.Name}, Time taken : {stopwatch.Elapsed.ToString()}");
        }
      }
      catch (Exception ex)
      { Error_Sender.SendNotification(ex.Message + ex.StackTrace); }

    }

    private void ThisAddIn_Shutdown(object sender, System.EventArgs e)
    {
    }
    //Start the Ribbon & Context Menu
    protected override Microsoft.Office.Core.IRibbonExtensibility CreateRibbonExtensibilityObject()
    {
      return new Ribbon();
    }

    #region Rules Manipulation






    public string fnGetSenderAddress(object address)
    {
      string email = null;

      if (address == null) return null;

      if (address is Outlook.Recipient)
      {
        Outlook.Recipient recipient = (Outlook.Recipient)address;
        email = recipient.Address;
        if (!Error_Sender.IsValidEmailAdd(email))
        {
          if (recipient.AddressEntry.GetExchangeUser() != null)
            email = recipient.AddressEntry.GetExchangeUser().PrimarySmtpAddress.Trim().ToLower();
          else if (Error_Sender.IsValidEmailAdd(recipient.Name))
            email = recipient.Name.Trim().ToLower();
        }
      }
      else if (address is Outlook.AddressEntry)
      {
        Outlook.AddressEntry address_entry = (Outlook.AddressEntry)address;
        email = address_entry.Address;
        if (!Error_Sender.IsValidEmailAdd(email))
        {
          if (address_entry.GetExchangeUser() != null)
            email = address_entry.GetExchangeUser().PrimarySmtpAddress.Trim().ToLower();
          else if (Error_Sender.IsValidEmailAdd(address_entry.Name))
            email = address_entry.Name.Trim().ToLower();
        }
      }

      return email;
    }
    #endregion

    #region Rule Category Assigning
    public bool fnAddEmailToRule_Category(string rule_name, string email_address, string _category)
    {
      Outlook.Rule rule = OutlookRules.fnFindRuleByName(rule_name);
      bool ok_added = false;
      bool email_exist = false;

      string recipient_address;
      object[] categories = { _category };

      //CREATE NEW RULE
      if (rule == null)
      {
        rule = this.OutlookRules.Create(rule_name, Outlook.OlRuleType.olRuleReceive);
        rule.Actions.AssignToCategory.Categories = categories;
        rule.Actions.AssignToCategory.Enabled = true;
        ok_added = true;
      }

      //CHECK IF THE EMAIL ADDRESS IS ALREADY ADDED
      if (rule.Conditions.From.Recipients.Count > 0)
      {
        foreach (Outlook.Recipient _recipient in rule.Conditions.From.Recipients)
        {
          recipient_address = fnGetSenderAddress(_recipient);
          if (!string.IsNullOrEmpty(recipient_address))
          {
            if (recipient_address.ToLower() == email_address.ToLower())
            {
              email_exist = true;
              break;
            }
          }
        }
      }

      //ADD THE NON EXISTING EMAILADDRESS
      if (!email_exist)
      {
        rule.Conditions.From.Recipients.Add(email_address);
        rule.Conditions.From.Recipients.ResolveAll();
        rule.Conditions.From.Enabled = true;
        ok_added = true;
      }
      return ok_added;
    }

    public bool fnAppendToMailCategory(ref string categories, string _category)
    {
      bool ok = false;

      if (string.IsNullOrWhiteSpace(categories))
      {
        categories = _category;
        ok = true;
      }
      else
      {
        categories = categories.Trim();
        if (!categories.Contains(_category))
        {
          //ADD CATEGORY
          if (!string.IsNullOrEmpty(categories) && !categories.EndsWith(","))
            categories += ", ";

          categories += _category;
          ok = true;
        }
      }
      return ok;
    }

    public Outlook.Category fnGetCategoryByName(string name)
    {
      Outlook.Category category = null;
      foreach (Outlook.Category _cat in Globals.ThisAddIn.Application.Session.Categories)
      {
        if (_cat.Name.ToLower() == name.ToLower())
        {
          category = _cat;
          break;
        }
      }
      if (category == null)
      {
        category = Globals.ThisAddIn.Application.Session.Categories.Add(name);
        category.Color = Outlook.OlCategoryColor.olCategoryColorBlue;
      }
      return category;
    }

    public bool fnRemoveMailCategory(ref string categories, string _category)
    {
      bool ok = false;

      if (!string.IsNullOrWhiteSpace(categories))
      {
        if (categories.Contains(_category))
        {
          categories = categories.Replace(string.Format("{0}, ", _category), "");
          categories = categories.Replace(string.Format("{0}", _category), "").Trim();
          if (categories.EndsWith(","))
            categories = categories.Substring(0, categories.Length - 1);

          ok = true;
        }
      }

      return ok;
    }
    #endregion

    #region VSTO generated code

    /// <summary>
    /// Required method for Designer support - do not modify
    /// the contents of this method with the code editor.
    /// </summary>
    private void InternalStartup()
    {
      this.Startup += new System.EventHandler(ThisAddIn_Startup);
      this.Shutdown += new System.EventHandler(ThisAddIn_Shutdown);
    }

    #endregion
  }

  #region Error Notification Class
  public class clsSendNotif
  {
    private string local_log_path = $"C:\\FarCap_Outlook_AddIn\\Error_{DateTime.Today:yyyyMMdd}.log";
    private const string Subject = "FarCap Outlook Add-In";

    public bool SendNotification(string str_message)
    {
      bool ok_sent = false;

      Outlook.MailItem mail = null;
      Outlook.Recipients mailRecipients = null;
      Outlook.Recipient mailRecipient = null;

      List<string> recipients = Split_Recipients(Properties.Settings.Default.Recipient);
      string ex_msg = "";
      try
      {
        if (recipients.Count > 0)
        {
          mail = (Outlook.MailItem)Globals.ThisAddIn.Application.CreateItem(Outlook.OlItemType.olMailItem);
          mail.Subject = Subject + " Error";
          mail.Body = "An exception has occured in the code of add-in.";
          mail.Body += "\n\n" + str_message;
          mailRecipients = mail.Recipients;

          foreach (string eadd in recipients)
          {
            mailRecipient = mailRecipients.Add(eadd);
            mailRecipient.Resolve();
          }
          if (mailRecipient.Resolved)
          {
            ((Outlook._MailItem)mail).Send();
            ok_sent = true;
          }
          else
          {
            ex_msg = "Unable to send the error notification";
          }
        }
        else
        {
          ex_msg = "No recipient.";
        }
      }
      catch (Exception ex)
      {
        ex_msg = ex.Message + ex.StackTrace;
      }
      finally
      {
        if (mailRecipient != null)
          Marshal.ReleaseComObject(mailRecipient);
        if (mailRecipients != null)
          Marshal.ReleaseComObject(mailRecipients);
        if (mail != null)
          Marshal.ReleaseComObject(mail);
      }

      if (!ok_sent)
        WriteLog(ex_msg, str_message);

      return ok_sent;

    }

    public bool SendTestNotification(string str_message, string str_recipients)
    {
      bool ok_sent = false;

      Outlook.MailItem mail = null;
      Outlook.Recipients mailRecipients = null;
      Outlook.Recipient mailRecipient = null;

      List<string> recipients = Split_Recipients(str_recipients);
      string ex_msg = "";
      try
      {
        if (recipients.Count > 0)
        {
          mail = (Outlook.MailItem)Globals.ThisAddIn.Application.CreateItem(Outlook.OlItemType.olMailItem);
          mail.Subject = Subject + " Error";
          mail.Body = "An exception has occured in the code of add-in.";
          mail.Body += "\n\n" + str_message;
          mailRecipients = mail.Recipients;
          foreach (string eadd in recipients)
          {
            mailRecipient = mailRecipients.Add(eadd);
            mailRecipient.Resolve();
          }
          if (mailRecipient.Resolved)
          {
            ((Outlook._MailItem)mail).Send();
            ok_sent = true;
          }
          else
          {
            ex_msg = "Unable to send the error notification";
          }
        }
        else
        {
          ex_msg = "No recipient.";
        }
      }
      catch (Exception ex)
      {
        ex_msg = ex.Message + ex.StackTrace;
      }
      finally
      {
        if (mailRecipient != null)
          Marshal.ReleaseComObject(mailRecipient);
        if (mailRecipients != null)
          Marshal.ReleaseComObject(mailRecipients);
        if (mail != null)
          Marshal.ReleaseComObject(mail);
      }

      if (!ok_sent)
        WriteLog(ex_msg, str_message);

      return ok_sent;
    }

    public void WriteLog(string ex_msg, string str_message)
    {

      StreamWriter writer = null;
      try
      {
        if (!Directory.Exists(Path.GetDirectoryName(local_log_path)))
          Directory.CreateDirectory(Path.GetDirectoryName(local_log_path));

        writer = new StreamWriter(local_log_path, true);
        if (!string.IsNullOrWhiteSpace(ex_msg))
        {
          writer.WriteLine("Unsend Notification Error: " + ex_msg);
        }
        writer.WriteLine($"Timestamp: {DateTime.Now:dd-MM-yyyy hh:mm:ss}, Message: {str_message}");
        writer.Close();
        writer.Dispose();
        writer = null;
      }
      catch (Exception ex)
      {
        MessageBox.Show("Failed to write log!\nException: " + ex.Message +
            "\n\nMessage:" + str_message, "FarCap Outlook Add-in");
      }
      finally
      {
        if (writer != null)
        {
          writer.Close();
          writer.Dispose();
        }
      }
    }

    public bool IsValidEmailAdd(string email_add)
    {
      Regex regex = new Regex(@"^([a-zA-Z0-9_\-\.]+)@((\[[0-9]{1,3}\.[0-9]{1,3}\.[0-9]{1,3}\.)|(([a-zA-Z0-9\-]+\.)+))([a-zA-Z]{2,4}|[0-9]{1,3})(\]?)$",
       RegexOptions.CultureInvariant | RegexOptions.Singleline);

      if (string.IsNullOrWhiteSpace(email_add))
        return false;

      return regex.IsMatch(email_add.Trim());
    }

    private List<string> Split_Recipients(string str_recipients)
    {
      string[] sp = str_recipients.Split(new char[] { ';' });
      List<string> rec = new List<string>();

      foreach (string s1 in sp)
      {
        if (IsValidEmailAdd(s1))
        {
          rec.Add(s1.Trim());
        }
      }
      return rec;
    }
  }
  #endregion

  #region ClassMailFolder object
  class SuperMailFolder
  {
    #region private variables
    Outlook.Folder _wrappedFolder;
    string _profileName;
    public List<SuperMailFolder> wrappedSubFolders = new List<SuperMailFolder>();
    string folderName = string.Empty;
    #endregion

    #region constructor
    internal SuperMailFolder(Outlook.Folder folder, string profileName)
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
        _wrappedFolder.Items.ItemChange += new Outlook.ItemsEvents_ItemChangeEventHandler(Items_ItemChange);

        //Go through all the subfolders and wrap them as well
        foreach (Outlook.Folder tmpFolder in _wrappedFolder.Folders)
        {
          if (folder.Name.StartsWith("Vault", StringComparison.InvariantCultureIgnoreCase) ||
              folder.Name.StartsWith("Public Folder", StringComparison.InvariantCultureIgnoreCase))
            continue;

          if (folder.DefaultItemType == Outlook.OlItemType.olMailItem)
          {
            var tmpWrapFolder = new SuperMailFolder(tmpFolder, _profileName);
            wrappedSubFolders.Add(tmpWrapFolder);
            wrappedSubFolders.AddRange(tmpWrapFolder.wrappedSubFolders);
          }
        }
      }
      catch (Exception ex)
      { Globals.ThisAddIn.Error_Sender.SendNotification(ex.Message + ex.StackTrace); }
    }
    #endregion

    private void Before_ItemMoveListener(object Item, Outlook.MAPIFolder TargetFolder, ref bool Cancel)
    {
      string src_rule_name = "";
      string tar_rule_name = "";
      string rule_prefix = "";

      bool ok_added = false;
      bool ok_removed = false;

      Outlook.MailItem oMsg = null;
      Outlook.Folder src_folder = null;

      try
      {
        if (Item is Outlook.MailItem && TargetFolder != null)
        {
          if (!TargetFolder.Name.ToLower().Equals("deleted items"))
          {
            oMsg = (Outlook.MailItem)Item;
            src_folder = (Outlook.Folder)oMsg.Parent;
            rule_prefix = Properties.Settings.Default.WatchFolder_Prefix.Trim();

            if (string.IsNullOrWhiteSpace(oMsg.SenderEmailAddress))
              return;

            //REMOVE RULE FROM SOURCE FOLDER
            if (src_folder.Name.ToLower() != "inbox" &&
                    src_folder.Name.ToLower().StartsWith(rule_prefix.ToLower()))
            {
              src_rule_name = rule_prefix + src_folder.Name;
              ok_removed = Globals.ThisAddIn.OutlookRules.fnRemoveEmailFromRule(src_rule_name, oMsg.SenderEmailAddress);
            }

            //DESTINATION FOLDER
            if (TargetFolder.Name.ToLower() != "inbox" &&
                    TargetFolder.Name.ToLower().StartsWith(rule_prefix.ToLower()))
            {

              tar_rule_name = rule_prefix + TargetFolder.Name;
              ok_added = Globals.ThisAddIn.OutlookRules.fnAddEmailToRule(tar_rule_name, oMsg.SenderEmailAddress, TargetFolder);
            }

            //Save rules
            if (Globals.ThisAddIn.OutlookRules != null && (ok_added || ok_removed))
              Globals.ThisAddIn.OutlookRules.Save(true);
          }
        }
      }
      catch (Exception ex)
      {
        Globals.ThisAddIn.Error_Sender.SendNotification(ex.Message + ex.StackTrace);
      }
    }

    void Items_ItemRemove()
    {

    }

    void Items_ItemChange(object item)
    {
      MessageBox.Show("Change");
      //Outlook.TaskItem task = item as Outlook.TaskItem;
      //if (task != null)
      //{
      //    MessageBox.Show(task.Subject);
      //}
    }

    #region Handler of addition item into a folder
    void Items_ItemAdd(object Item)
    {
      try
      {
        if (Item is Outlook.Folder)
        {
          SuperMailFolder tmpWrapFolder = new SuperMailFolder((Outlook.Folder)Item, _profileName);
          wrappedSubFolders.Add(tmpWrapFolder);
          wrappedSubFolders.AddRange(tmpWrapFolder.wrappedSubFolders);
        }
      }
      catch (Exception ex)
      {
        Globals.ThisAddIn.Error_Sender.SendNotification(ex.Message + ex.StackTrace);
      }
    }

    #endregion
  }
  #endregion

  public class GlobalRules
  {
    private readonly Outlook.Application _application;
    private readonly ThisAddIn _thisAddIn;
    public Outlook.Rules Rules = null;

    public GlobalRules(Outlook.Application application, ThisAddIn thisAddIn)
    {
      _application = application;
      _thisAddIn = thisAddIn;
      Rules = _application.Session.DefaultStore.GetRules();
    }

    public Outlook.Rule fnFindRuleByName(string rule_name)
    {
      Outlook.Rule rule = null;
      if (this.Rules != null && !string.IsNullOrEmpty(rule_name))
      {
        foreach (Outlook.Rule r in this.Rules)
        {
          if (r.Name.ToLower() == rule_name.ToLower())
          {
            rule = r;
            break;
          }
        }
      }
      return rule;
    }

    public bool fnAddEmailToRule(string rule_name, string email_address, Outlook.MAPIFolder target_folder)
    {
      Outlook.Rule rule = fnFindRuleByName(rule_name);
      bool ok_added = false;
      bool email_exist = false;
      string recipient_address;

      //CREATE NEW RULE
      if (rule == null)
      {
        rule = this.Rules.Create(rule_name, Outlook.OlRuleType.olRuleReceive);
        rule.Actions.MoveToFolder.Folder = (target_folder);
        rule.Actions.MoveToFolder.Enabled = true;
        ok_added = true;
      }
      //CHECK IF THE EMAIL ADDRESS IS ALREADY ADDED
      if (rule.Conditions.From.Recipients.Count > 0)
      {
        foreach (Outlook.Recipient _recipient in rule.Conditions.From.Recipients)
        {
          recipient_address = _thisAddIn.fnGetSenderAddress(_recipient);
          if (!string.IsNullOrEmpty(recipient_address))
          {
            if (recipient_address.ToLower() == email_address.ToLower())
            {
              email_exist = true;
              break;
            }
          }
        }
      }

      //ADD THE NON EXISTING EMAILADDRESS
      if (!email_exist)
      {
        rule.Conditions.From.Recipients.Add(email_address);
        rule.Conditions.From.Recipients.ResolveAll();
        rule.Conditions.From.Enabled = true;
        ok_added = true;
      }
      return ok_added;
    }

    public bool fnRemoveEmailFromRule(string rule_name, string email_address)
    {
      string recipient_address;
      bool ok_remove = false;
      Outlook.Rule src_rule = this.fnFindRuleByName(rule_name);

      if (src_rule != null)
      {
        foreach (Outlook.Recipient _recipient in src_rule.Conditions.From.Recipients)
        {
          recipient_address = _thisAddIn.fnGetSenderAddress(_recipient);
          if (!string.IsNullOrEmpty(recipient_address))
          {
            if (recipient_address.ToLower() == email_address.ToLower())
            {
              _recipient.Delete();
              _recipient.Resolve();
              ok_remove = true;
              break;
            }
          }
        }
      }
      if (src_rule != null)
      {
        if (src_rule.Conditions.From.Recipients.Count == 0)
        {
          this.Rules.Remove(rule_name);
          ok_remove = true;
        }
      }

      return ok_remove;
    }

    public Outlook.Rule Create(string tarRulename, Outlook.OlRuleType olRuleReceive)
    {
      return Rules.Create(tarRulename, Outlook.OlRuleType.olRuleReceive);
    }

    public void Remove(string srcRulename)
    {
      Rules.Remove(srcRulename);
    }

    public void Save(bool b)
    {
      Rules.Save(b);
      Reload();
    }

    public void Reload()
    {
      Rules = _application.Session.DefaultStore.GetRules();
    }
  }

  #region SenderData
  public class SenderData
  {
    public string Folder_name;
    public string Name;
    public string EmailAddress;
    public string SenderType;

    public SenderData(string _foldername, string _name, string _emailaddress, string _emailtype)
    {
      this.Folder_name = _foldername;
      this.Name = _name;
      this.EmailAddress = _emailaddress;
      this.SenderType = _emailtype;
    }
  }
  #endregion

}
