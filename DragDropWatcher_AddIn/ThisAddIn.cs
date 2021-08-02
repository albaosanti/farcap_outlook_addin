﻿using System;
using System.Diagnostics;
using Outlook = Microsoft.Office.Interop.Outlook;
using Office = Microsoft.Office.Core;

namespace DragDrapWatcher_AddIn
{
  public partial class ThisAddIn
  {
    #region Global Variables
    public string CAT_RULE_PREFIX = string.IsNullOrWhiteSpace(Properties.Settings.Default.CategoryRulePrefix) ? "#fcap_cat_" : Properties.Settings.Default.CategoryRulePrefix;
    public GlobalRules OutlookRules = null;
    public ClsSendNotif Error_Sender = null;
    #endregion

    private void ThisAddIn_Startup(object sender, System.EventArgs e)
    {
      Outlook._NameSpace outNS = null;
      SuperMailFolder folderToWrap = null;
      
      try
      {

        Error_Sender = new ClsSendNotif();
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

      if (address is Outlook.Recipient recipient)
      {
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
      else if (address is Outlook.AddressEntry addressEntry)
      {
        email = addressEntry.Address;
        if (Error_Sender.IsValidEmailAdd(email))
          return email;
        email = addressEntry.PropertyAccessor.GetProperty(PR_SMTP_ADDRESS);
        if (Error_Sender.IsValidEmailAdd(email))
          return email;
        email = addressEntry.GetExchangeUser()?.PrimarySmtpAddress.Trim().ToLower();
        if (Error_Sender.IsValidEmailAdd(email))
          return email;
        email = addressEntry.Name.Trim().ToLower();
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
}
