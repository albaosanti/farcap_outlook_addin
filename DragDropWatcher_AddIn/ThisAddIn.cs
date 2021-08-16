using System;
using System.Diagnostics;
using System.Reflection;
using Outlook = Microsoft.Office.Interop.Outlook;
using Office = Microsoft.Office.Core;

namespace DragDrapWatcher_AddIn
{
  public partial class ThisAddIn
  {
    public string CAT_RULE_PREFIX = string.IsNullOrWhiteSpace(Properties.Settings.Default.CategoryRulePrefix) ? "#fcap_cat_" : Properties.Settings.Default.CategoryRulePrefix;
    public GlobalRules OutlookRules = null;
    public ClsSendNotif Error_Sender = null;

    private void ThisAddIn_Startup(object sender, System.EventArgs e)
    {
      string loggerPrefix = $"{this.GetType().Name}->{MethodBase.GetCurrentMethod().Name} ::";
      SuperMailFolder folderToWrap = null;
      try
      {
        Error_Sender = new ClsSendNotif();
        Error_Sender.WriteLog($"{loggerPrefix} =============== Beginning Startup ===============");
        Outlook.Application application = this.Application;
        Outlook._NameSpace outNS = application.GetNamespace("MAPI");
        string profileName = outNS.CurrentUser.Name;
        Error_Sender.WriteLog($"{loggerPrefix}  Profile Name : {profileName}");
        //DRAG & DROP WILL BE CREATED HERE
        Outlook.Folders folders = outNS.Folders;
        foreach (Outlook.Folder folder in folders)
        {
          try
          {
            var stopwatch = Stopwatch.StartNew();
            Error_Sender.WriteLog(string.Empty,
              $"{loggerPrefix}  Start Scanning folder :: Name: {folder.Name}");
            if (folder.Name.Contains("Vault") ||
                folder.Name.StartsWith("Public Folder", StringComparison.InvariantCultureIgnoreCase))
            {
              stopwatch.Stop();
              Error_Sender.WriteLog(string.Empty,
                $"{loggerPrefix}  Skip Scanning folder :: Name: {folder.Name}, Time taken : {stopwatch.Elapsed.ToString()}");
              continue;
            }

            if (folder.DefaultItemType == Outlook.OlItemType.olMailItem)
              folderToWrap = new SuperMailFolder(folder, profileName);
            stopwatch.Stop();
            Error_Sender.WriteLog(string.Empty,
              $"{loggerPrefix}  End Scanning folder :: Name: {folder.Name}, Time taken : {stopwatch.Elapsed.ToString()}");
          }
          catch (Exception ex)
          {
            Error_Sender.WriteLog(string.Empty,
              $"{loggerPrefix}  Exception recorded on scanning folder {ex.ToString()}");
          }
        }
        Error_Sender.WriteLog($"{loggerPrefix}  Completed Scanning folders list");

        OutlookRules = new GlobalRules(application, this);

        Error_Sender.WriteLog($"{loggerPrefix}  =============== Startup Completed ===============");
      }
      catch (System.Exception ex)
      {
        Error_Sender.SendNotification(ex.Message + ex.StackTrace);
      }
    }

    private void ThisAddIn_Shutdown(object sender, System.EventArgs e)
    {
    }
    //Start the Ribbon & Context Menu
    protected override Microsoft.Office.Core.IRibbonExtensibility CreateRibbonExtensibilityObject()
    {
      return new Ribbon();
    }

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

    /// <summary>
    /// Required method for Designer support - do not modify
    /// the contents of this method with the code editor.
    /// </summary>
    private void InternalStartup()
    {
      this.Startup += ThisAddIn_Startup;
      this.Shutdown += ThisAddIn_Shutdown;
    }
  }
}
