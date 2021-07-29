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
using System.CodeDom;
using Microsoft.Office.Interop.Outlook;

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
      catch (System.Exception ex)
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
      string PR_SMTP_ADDRESS = @"http://schemas.microsoft.com/mapi/proptag/0x39FE001E";
      if (address == null) return null;

      if (address is Outlook.Recipient)
      {
        Outlook.Recipient recipient = (Outlook.Recipient)address;

        email = recipient.Address;
        if (Error_Sender.IsValidEmailAdd(email))
          return email;
        email = recipient.PropertyAccessor.GetProperty(PR_SMTP_ADDRESS);
        if (Error_Sender.IsValidEmailAdd(email))
          return email;
        email = recipient.AddressEntry.GetExchangeUser()?.PrimarySmtpAddress.Trim().ToLower();
        if (Error_Sender.IsValidEmailAdd(email))
          return email;
        email = recipient.Name.Trim().ToLower();
        if (Error_Sender.IsValidEmailAdd(email))
          return email;
      }
      else if (address is Outlook.AddressEntry)
      {
        Outlook.AddressEntry address_entry = (Outlook.AddressEntry)address;
        email = address_entry.Address;
        if (Error_Sender.IsValidEmailAdd(email))
          return email;
        email = address_entry.GetExchangeUser()?.PrimarySmtpAddress.Trim().ToLower();
        if (Error_Sender.IsValidEmailAdd(email))
          return email;
        email = address_entry.Name.Trim().ToLower();
        if (Error_Sender.IsValidEmailAdd(email))
          return email;
      }

      return email;
    }
    #endregion

    #region Rule Category Assigning
    public bool fnAddEmailToRule_Category(string rule_name, string email_address, string _category)
    {
      Outlook.Rule rule = OutlookRules.FindRuleByName(rule_name);
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

  #region FarCapSender Class
  public class FarCapSender
  {
    public string rulename;
    public string sender_name;
    public string sender_email;
    public string folder_name;
    public string folder_path;
    public int rule_number;

    public FarCapSender(string _rulename, string _email, string _name, string _folder, string _folder_path)
    {
      this.rulename = _rulename;
      this.sender_email = _email;
      this.sender_name = _name;
      this.folder_name = _folder;
      this.rule_number = 0;
      //GET FOLDER INDEX
      var idx = rulename.LastIndexOf('_');
      if (idx > -1 && idx < rulename.Length - 1)
        int.TryParse(rulename.Substring(idx + 1), out rule_number);
    }
  }
  #endregion

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
      string ex_msg = string.Empty;
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
      catch (System.Exception ex)
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
      string ex_msg = string.Empty;
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
      catch (System.Exception ex)
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
      catch (System.Exception ex)
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
        //_wrappedFolder.Items.ItemChange += new Outlook.ItemsEvents_ItemChangeEventHandler(Items_ItemChange);

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
      catch (System.Exception ex)
      { Globals.ThisAddIn.Error_Sender.SendNotification(ex.Message + ex.StackTrace); }
    }
    #endregion

    private void Before_ItemMoveListener(object Item, Outlook.MAPIFolder TargetFolder, ref bool Cancel)
    {
      string src_ruleprefix = string.Empty;
      string target_ruleprefix = string.Empty;
      string rule_prefix = string.Empty;
      string folder_prefix = string.Empty;
      string sender_address = string.Empty;

      bool ok_added = false;
      bool ok_removed = false;

      Outlook.MailItem oMsg = null;
      Outlook.Folder src_folder = null;

      try
      {
        if (Item is Outlook.MailItem && TargetFolder != null)
        {
          if (!TargetFolder.Name.StartsWith("deleted items",StringComparison.OrdinalIgnoreCase))
          {
            oMsg = (Outlook.MailItem)Item;
            src_folder = (Outlook.Folder)oMsg.Parent;
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
        if (Item is Outlook.Folder)
        {
          SuperMailFolder tmpWrapFolder = new SuperMailFolder((Outlook.Folder)Item, _profileName);
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
  #endregion

  public class GlobalRules
  {
    private readonly Outlook.Application _application;
    private readonly ThisAddIn _thisAddIn;

    public Outlook.Rules Rules = null;
    public List<FarCapSender> FarCapRuleSenders = null;  

    public GlobalRules(Outlook.Application application, ThisAddIn thisAddIn)
    {
      _application = application;
      _thisAddIn = thisAddIn;
      this.Reload();
    }

    public List<string> GetGroupRuleNames(string rulename_prefix)
    {
      List<string> rule_names = new List<string>();
      if (FarCapRuleSenders != null)
      {
        var list = FarCapRuleSenders.Where(row =>
                      row.rulename.StartsWith(rulename_prefix, StringComparison.OrdinalIgnoreCase))
                      .GroupBy(g => new { g.rulename })
                      .Select(s => new { Name = s.Key.rulename, Count = s.Count()}).ToList();

        if(list != null)
        {
          foreach (var item in list)
            rule_names.Add(item.Name);
        }
      }
      return rule_names;
    }

    public void ClearRuleGroups(string rulename_prefix)
    {
      if (FarCapRuleSenders != null)
      {
        var list = FarCapRuleSenders.Where(row =>
                     row.rulename.StartsWith(rulename_prefix, StringComparison.OrdinalIgnoreCase))
                     .GroupBy(g => new { g.rulename })
                     .Select(s => new { Name = s.Key.rulename, Count = s.Count() }).ToList();

        if (list != null)
        {
          foreach (var item in list)
            this.Remove(item.Name);
        }
      }
    }

    private string GetTargetRulenameGroup(string rulename_prefix)
    {
      string target_rulename = null;
      var groups = FarCapRuleSenders.Where(row => row.rulename.StartsWith(rulename_prefix, StringComparison.OrdinalIgnoreCase))
              .GroupBy(
                  g => new { g.rulename, g.rule_number }
              ).Select(
                  new_group => new
                      {
                        Name = new_group.Key.rulename,
                        Number = new_group.Key.rule_number,
                        Count = new_group.Count()
                      }
              ).OrderBy(row => row.Number);

      if (groups != null)
      {
        foreach (var g in groups)
        {
          if (g.Count < Properties.Settings.Default.MaxRuleRecipients)
          {
            target_rulename = g.Name;
            break;
          }
        }
      }
     
      //RULES ARE FULL or NO RULE CREATED YET 
      if(target_rulename == null)
        target_rulename = rulename_prefix + "_" + Convert.ToString((groups.Count() > 0 ? (groups.Last().Number + 1) : 1));

      return target_rulename;
    }


    public Outlook.Rule FindRuleByName(string rule_name)
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

    public bool AddEmailToRule(string rulename_prefix, string email_address, string sender_name, Outlook.MAPIFolder target_folder)
    {
      Outlook.Rule rule = null;
      string target_rulename = null;
      bool ok_added = false;
      bool email_exist = false;

      if (FarCapRuleSenders == null || Rules == null) Reload();

      var existing_emails = FarCapRuleSenders.FindAll(
         row => row.sender_email.Equals(email_address, StringComparison.OrdinalIgnoreCase)
      ).ToList();

      //remove from other rules && FarCapSenderList
      if (existing_emails != null)
      {
        foreach (var row in existing_emails)
        {
          if (row.rulename.StartsWith(rulename_prefix, StringComparison.OrdinalIgnoreCase))
            email_exist = true;
          else
            RemoveEmailFromRule(row.rulename, row.sender_email);
        }
      }
     
      if (!email_exist)
      {
        target_rulename = GetTargetRulenameGroup(rulename_prefix);
        rule = FindRuleByName(target_rulename);       
        if (rule == null)
        {
          rule = this.Rules.Create(target_rulename, Outlook.OlRuleType.olRuleReceive);
          rule.Actions.MoveToFolder.Folder = (target_folder);
          rule.Actions.MoveToFolder.Enabled = true;
        }
        rule.Conditions.From.Recipients.Add(email_address);
        rule.Conditions.From.Recipients.ResolveAll();
        rule.Conditions.From.Enabled = true;
        //ADD TO FarCapSenders
        FarCapRuleSenders.Add(new FarCapSender(target_rulename, 
          email_address, 
          sender_name,
          target_folder.Name,
          target_folder.FolderPath));

        ok_added = true;
      }
      return ok_added;
    }

    public bool RemoveEmailFromRule(string rule_name, string email_address)
    {
      string recipient_address;
      bool ok_remove = false;
      Outlook.Rule src_rule = null;

      if (FarCapRuleSenders == null || Rules == null) Reload();

      if (FarCapRuleSenders.Exists(row => row.rulename.Equals(rule_name, StringComparison.OrdinalIgnoreCase)
             && row.sender_email.Equals(email_address, StringComparison.OrdinalIgnoreCase)))
      {
        src_rule = this.FindRuleByName(rule_name);
        if (src_rule != null)
        {
          foreach (Outlook.Recipient _recipient in src_rule.Conditions.From.Recipients)
          {
            recipient_address = _thisAddIn.fnGetSenderAddress(_recipient);
            if (!string.IsNullOrEmpty(recipient_address))
            {
              if (recipient_address.Equals(email_address,StringComparison.OrdinalIgnoreCase))
              {
                _recipient.Delete();
                _recipient.Resolve();
                ok_remove = true;
                break;
              }
            }
          }
          if (src_rule.Conditions.From.Recipients.Count == 0)
          {
            this.Rules.Remove(rule_name);
            ok_remove = true;
          }
        }

        //REMOVE FROM FARCAPSENDER
        FarCapRuleSenders.RemoveAll(
            row => row.rulename.Equals(rule_name, StringComparison.OrdinalIgnoreCase) &&
            row.sender_email.Equals(email_address, StringComparison.OrdinalIgnoreCase));
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
      FarCapRuleSenders.RemoveAll(
        row => row.rulename.Equals(srcRulename,
        StringComparison.OrdinalIgnoreCase));
    }

    public void Save(bool b)
    {
      Rules.Save(b);
      Reload();
    }

    public void Reload()
    {
      Rules = _application.Session.DefaultStore.GetRules();
      FarCapRuleSenders = new List<FarCapSender>();
      if (Rules != null)
      {
        foreach (Outlook.Rule rule in Rules)
        {
          if (rule.Name.Trim().StartsWith(Properties.Settings.Default.RuleName_Prefix, StringComparison.OrdinalIgnoreCase))
          {
            foreach (Outlook.Recipient _recipient in rule.Conditions.From.Recipients)
              this.FarCapRuleSenders.Add(new FarCapSender(
                rule.Name,
                _thisAddIn.fnGetSenderAddress(_recipient),
                _recipient.Name,
                rule.Actions.MoveToFolder.Folder.Name,
                rule.Actions.MoveToFolder.Folder.FolderPath));
          }
        }
      }
    }
  }
}
