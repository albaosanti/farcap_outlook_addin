using System;
using System.Collections.Generic;
using System.Linq;
using System.Reflection;

namespace DragDrapWatcher_AddIn
{
  public class GlobalRules
  {
    private readonly Microsoft.Office.Interop.Outlook.Application _application;
    private readonly ThisAddIn _thisAddIn;

    public Microsoft.Office.Interop.Outlook.Rules Rules = null;
    public List<FarCapSender> FarCapRuleSenders = null;

    public GlobalRules(Microsoft.Office.Interop.Outlook.Application application, ThisAddIn thisAddIn)
    {
      _application = application;
      _thisAddIn = thisAddIn;
      this.Reload();
    }

    public void ClearRuleGroups(string rulename_prefix)
    {
      string loggerPrefix = $"{this.GetType().Name}->{MethodBase.GetCurrentMethod().Name} ::";
      _thisAddIn.Error_Sender.WriteLog($"{loggerPrefix}  called with rulename_prefix-> {rulename_prefix}");
      if (FarCapRuleSenders == null)
      {
        _thisAddIn.Error_Sender.WriteLog($"{loggerPrefix}  NUll FarCapRuleSenders -> {FarCapRuleSenders == null} or NULL Rules {Rules == null} !");
        return;
      }

      var list = FarCapRuleSenders.Where(row =>
          row.rulename.StartsWith(rulename_prefix, StringComparison.OrdinalIgnoreCase))
        .GroupBy(g => new { g.rulename })
        .Select(s => new { Name = s.Key.rulename, Count = s.Count() }).ToList();

      foreach (var item in list)
        this.Remove(item.Name);
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
      if (target_rulename == null)
        target_rulename = rulename_prefix + "_" + Convert.ToString((groups.Count() > 0 ? (groups.Last().Number + 1) : 1));

      return target_rulename;
    }


    public Microsoft.Office.Interop.Outlook.Rule FindRuleByName(string rule_name)
    {
      Microsoft.Office.Interop.Outlook.Rule rule = null;
      if (this.Rules != null && !string.IsNullOrEmpty(rule_name))
      {
        foreach (Microsoft.Office.Interop.Outlook.Rule r in this.Rules)
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

    public bool AddEmailToRule(string rulename_prefix, string email_address, string sender_name, Microsoft.Office.Interop.Outlook.MAPIFolder target_folder)
    {
      string loggerPrefix = $"{this.GetType().Name}->{MethodBase.GetCurrentMethod().Name} ::";
      Microsoft.Office.Interop.Outlook.Rule rule = null;
      string target_rulename = null;
      bool ok_added = false;
      bool email_exist = false;

      if (FarCapRuleSenders == null || Rules == null)
      {
        _thisAddIn.Error_Sender.WriteLog($"{loggerPrefix}  NUll FarCapRuleSenders -> {FarCapRuleSenders == null} or NULL Rules {Rules == null} !");
        Reload();
      }

      var existing_emails = FarCapRuleSenders.FindAll(
        row => row.sender_email.Equals(email_address, StringComparison.OrdinalIgnoreCase)
      ).ToList();

      //remove from other rules && FarCapSenderList
      foreach (var row in existing_emails)
      {
        if (row.rulename.StartsWith(rulename_prefix, StringComparison.OrdinalIgnoreCase))
          email_exist = true;
        else
        {
          _thisAddIn.Error_Sender.WriteLog($"{loggerPrefix}  Removing {row.sender_email} from {row.rulename} !");
          RemoveEmailFromRule(row.rulename, row.sender_email);
        }
      }

      if (!email_exist)
      {
        target_rulename = GetTargetRulenameGroup(rulename_prefix);
        rule = FindRuleByName(target_rulename);
        if (rule == null)
        {
          _thisAddIn.Error_Sender.WriteLog($"{loggerPrefix}  Creating Rule {target_rulename}!");
          rule = this.Rules.Create(target_rulename, Microsoft.Office.Interop.Outlook.OlRuleType.olRuleReceive);
          rule.Actions.MoveToFolder.Folder = (target_folder);
          rule.Actions.MoveToFolder.Enabled = true;
        }
        _thisAddIn.Error_Sender.WriteLog($"{loggerPrefix}  Adding {email_address} to {target_rulename}!");
        rule.Conditions.From.Recipients.Add(email_address);
        rule.Conditions.From.Recipients.ResolveAll();
        rule.Conditions.From.Enabled = true;
        //ADD TO FarCapSenders
        FarCapRuleSenders.Add(new FarCapSender(target_rulename,
          email_address,
          sender_name,
          target_folder.Name));

        ok_added = true;
      }
      return ok_added;
    }

    public bool RemoveEmailFromRule(string rule_name, string email_address)
    {
      string recipient_address;
      bool ok_remove = false;
      Microsoft.Office.Interop.Outlook.Rule src_rule = null;
      string loggerPrefix = $"{this.GetType().Name}->{MethodBase.GetCurrentMethod().Name} ::";

      if (FarCapRuleSenders == null || Rules == null)
      {
        _thisAddIn.Error_Sender.WriteLog($"{loggerPrefix}  NUll FarCapRuleSenders -> {FarCapRuleSenders == null} or NULL Rules {Rules == null} !");
        Reload();
      }

      if (FarCapRuleSenders.Exists(row => row.rulename.Equals(rule_name, StringComparison.OrdinalIgnoreCase)
                                          && row.sender_email.Equals(email_address, StringComparison.OrdinalIgnoreCase)))
      {
        src_rule = this.FindRuleByName(rule_name);
        if (src_rule != null)
        {
          foreach (Microsoft.Office.Interop.Outlook.Recipient _recipient in src_rule.Conditions.From.Recipients)
          {
            recipient_address = _thisAddIn.fnGetSenderAddress(_recipient);
            if (string.IsNullOrEmpty(recipient_address)) continue;
            if (!recipient_address.Equals(email_address, StringComparison.OrdinalIgnoreCase)) continue;

            _thisAddIn.Error_Sender.WriteLog($"{loggerPrefix}  Removing -> {recipient_address} from {src_rule.Name} !");
            _recipient.Delete();
            _recipient.Resolve();
            ok_remove = true;
            break;
          }

          if (src_rule.Conditions.From.Recipients.Count == 0)
          {
            _thisAddIn.Error_Sender.WriteLog($"{loggerPrefix}  Removing rule {src_rule.Name} as it has no recipient");
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

    public Microsoft.Office.Interop.Outlook.Rule Create(string tarRulename, Microsoft.Office.Interop.Outlook.OlRuleType olRuleReceive)
    {
      string loggerPrefix = $"{this.GetType().Name}->{MethodBase.GetCurrentMethod().Name} ::";
      _thisAddIn.Error_Sender.WriteLog($"{loggerPrefix}  called with tarRulename-> {tarRulename}");
      return Rules.Create(tarRulename, Microsoft.Office.Interop.Outlook.OlRuleType.olRuleReceive);
    }

    public void Remove(string srcRulename)
    {
      string loggerPrefix = $"{this.GetType().Name}->{MethodBase.GetCurrentMethod().Name} ::";
      _thisAddIn.Error_Sender.WriteLog($"{loggerPrefix}  called with srcRulename-> {srcRulename}");
      Rules.Remove(srcRulename);
      FarCapRuleSenders.RemoveAll(
        row => row.rulename.Equals(srcRulename,
          StringComparison.OrdinalIgnoreCase));
    }

    public void Save(bool b)
    {
      string loggerPrefix = $"{this.GetType().Name}->{MethodBase.GetCurrentMethod().Name} ::";
      _thisAddIn.Error_Sender.WriteLog($"{loggerPrefix}  called with b-> {b}");
      Rules.Save(b);
      Reload();
    }

    public void Reload()
    {
      string loggerPrefix = $"{this.GetType().Name}->{MethodBase.GetCurrentMethod().Name} ::";

      _thisAddIn.Error_Sender.WriteLog($"{loggerPrefix} Getting Existing Rules!");
      Rules = _application.Session.DefaultStore.GetRules();
      if (Rules == null)
      {
        _thisAddIn.Error_Sender.WriteLog($"{loggerPrefix} Rules is null!");
        return;
      }

      FarCapRuleSenders = new List<FarCapSender>();
      foreach (Microsoft.Office.Interop.Outlook.Rule rule in Rules)
      {
        if (!rule.Name.Trim()
          .StartsWith(Properties.Settings.Default.RuleName_Prefix, StringComparison.OrdinalIgnoreCase)) continue;

        _thisAddIn.Error_Sender.WriteLog($"{loggerPrefix} Scanning rule {rule.Name} & it has {rule.Conditions.From.Recipients.Count} Recipients");

        foreach (Microsoft.Office.Interop.Outlook.Recipient _recipient in rule.Conditions.From.Recipients)
        {
          string fnGetSenderAddress = string.Empty, recipientName = string.Empty, folderName = string.Empty;
          try
          {
            fnGetSenderAddress = _thisAddIn.fnGetSenderAddress(_recipient);
            recipientName = _recipient.Name;
            var ruleActions = rule.Actions;
            var ruleActionsMoveToFolder = ruleActions.MoveToFolder;
            var mapiFolder = ruleActionsMoveToFolder.Folder;
            folderName = mapiFolder.Name;
            var farCapSender = new FarCapSender(rule.Name, fnGetSenderAddress, recipientName, folderName);
            this.FarCapRuleSenders.Add(farCapSender);
          }
          catch (Exception e)
          {
            _thisAddIn.Error_Sender.WriteLog($"{loggerPrefix} {fnGetSenderAddress} {recipientName} {folderName} Exception Message {e.Message}  {e.StackTrace}");
          }
        }
      }
      _thisAddIn.Error_Sender.WriteLog($"{loggerPrefix} Prepared {FarCapRuleSenders.Count} FarCapRuleSenders");
    }
  }
}