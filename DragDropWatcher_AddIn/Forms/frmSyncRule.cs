using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using Outlook = Microsoft.Office.Interop.Outlook;

namespace DragDrapWatcher_AddIn
{
  public partial class frmSyncRule : Form
  {
    #region Variables
    public Outlook.Folder parent_folder;

    private List<GroupSender> scanned_senders;
    private string rulename_prefix;
    private bool cancelled = false;
    private int scan_period;

    private delegate void delFillGrid();
    #endregion

    #region Class
    private class GroupSender
    {
      public FarCapSender sender;
      public int count;
      public GroupSender(FarCapSender _sender, int _count)
      {
        this.sender = _sender;
        this.count = _count;
      }
    }
    #endregion

    #region Methods
    private void fnFillGrid()
    {
      dgvList.Rows.Clear();
      foreach (var row in scanned_senders)
        dgvList.Rows.Add(new object[] { row.sender.sender_name, row.sender.sender_email, row.count });
    }

    private int CountSendersInRule()
    {
      return Globals.ThisAddIn.OutlookRules.FarCapRuleSenders.Where(
        item => item.rulename.StartsWith(rulename_prefix, StringComparison.OrdinalIgnoreCase)
      ).ToList().Count();
    }

    private string getFilterScanToDate()
    {
      DateTime scan_to = DateTime.Now;
      string filter = string.Empty;
      if (!chkAll.Checked)
      {
        switch (scan_period)
        {
          case 0:
            scan_to = scan_to.AddMonths(-(Convert.ToInt32(numScan.Value)));
            break;
          case 1:
            scan_to = scan_to.AddDays(-(Convert.ToDouble(numScan.Value) * 7));
            break;
          case 2:
            scan_to = scan_to.AddDays(-(Convert.ToDouble(numScan.Value)));
            break;
          default:
            break;
        }
        filter = "[Received]>'" + scan_to.AddDays(-1).ToShortDateString() + "'";
      }
      return filter;
    }
    #endregion

    public frmSyncRule()
    {
      InitializeComponent();
    }

    private void InitLoad()
    {
      rulename_prefix = Properties.Settings.Default.RuleName_Prefix + parent_folder.Name;
      lblFolderName.Text = parent_folder.Name;
      lblRuleSenders.Text = $"{CountSendersInRule()} unique sender/s.";

      numScan.Value = 3;
      cmbPeriod.SelectedIndex = 0;

      cancelled = false;
      chkAll.Checked = false;

      btnProcess.Text = "Start";
      btnSave.Text = "Save";
      btnProcess.Enabled = true;
      btnSave.Enabled = false;
      dgvList.Rows.Clear();
    }

    private void btnProcess_Click(object sender, EventArgs e)
    {
      if (btnProcess.Text.ToLower() == "start")
      {
        cancelled = false;
        btnProcess.Text = "Stop";
        lblFoundSenders.Text = "Scanning..";
        lblStatus.Text = "Started..";

        dgvList.Rows.Clear();
        dgvList.Enabled = false;
        chkAll.Enabled = false;
        numScan.Enabled = false;
        cmbPeriod.Enabled = false;
        scan_period = cmbPeriod.SelectedIndex;

        bgwProcess.RunWorkerAsync();
      }
      else
      {
        if (MessageBox.Show("Are you sure to STOP the running process?", "Confirm Exit - FarCap Add-In",
            MessageBoxButtons.YesNo, MessageBoxIcon.Question, MessageBoxDefaultButton.Button2) == System.Windows.Forms.DialogResult.Yes)
        {
          cancelled = true;
          btnProcess.Text = "Stopping..";
          btnProcess.Enabled = false;
        }
      }
    }

    private void bgwProcess_DoWork(object sender, DoWorkEventArgs e)
    {
      Outlook.Table items;
      Outlook.NameSpace ns;

      var stopWatch = System.Diagnostics.Stopwatch.StartNew();
      var uniqueEmails = new HashSet<string>();
      var progress = 0;
      var count = 0;

      try
      {
        ns = Globals.ThisAddIn.Application.GetNamespace("MAPI");

        scanned_senders = new List<GroupSender>();
        items = parent_folder.GetTable(getFilterScanToDate());
        count = items.GetRowCount();

        items.Columns.RemoveAll();
        items.Columns.Add("EntryID");
        items.Columns.Add("SenderName");
        items.Columns.Add("SenderEmailAddress");
        items.Columns.Add("SenderEmailType");
        items.Columns.Add("ReceivedTime");

        for (int index = 0; index < count; index++)
        {
          if (cancelled) break;
          progress = Convert.ToInt32(((Convert.ToDouble(index + 1) / Convert.ToDouble(count)) * 100));
          bgwProcess.ReportProgress(progress, "Processing " + (index + 1) + "/" + count);

          var mailItem = items.GetNextRow();
          var entryid = (string)mailItem["EntryID"];
          var name = (string)mailItem["SenderName"];
          var senderType = (string)mailItem["SenderEmailType"];
          var emailAddress = (string)mailItem["SenderEmailAddress"];
          
          if (senderType.Equals("EX", StringComparison.OrdinalIgnoreCase) && 
            !Globals.ThisAddIn.Error_Sender.IsValidEmailAdd(emailAddress))
          {
            var mail_ = (Outlook.MailItem) ns.GetItemFromID(entryid);
            if (mail_ != null)
               emailAddress = Globals.ThisAddIn.fnGetSenderAddress(mail_.Sender);           
          }

          if (!string.IsNullOrEmpty(emailAddress))
          {
            emailAddress = emailAddress.ToLower();
            var senderdata = new FarCapSender("", emailAddress, name, "", "");

            if (uniqueEmails.Add(emailAddress))
              scanned_senders.Add(new GroupSender(senderdata, 1));
            else
            {
              var idx = scanned_senders.FindIndex(itm => itm.sender.sender_email.Equals(emailAddress, StringComparison.OrdinalIgnoreCase));
              if (idx > -1) scanned_senders[idx].count++;
            }
          }
        }

        bgwProcess.ReportProgress(0, "Populating grid..");
        scanned_senders = scanned_senders.OrderByDescending(itm => itm.count).ToList();
        this.Invoke(new delFillGrid(fnFillGrid));
      }
      catch (Exception ex)
      {
        Globals.ThisAddIn.Error_Sender.SendNotification("@bgwProcess_SyncRule>> " + ex.Message + ex.StackTrace);
        bgwProcess.ReportProgress(0, "Error: " + ex.Message);
        scanned_senders = null;
      }
      stopWatch.Stop();

      if (cancelled)
        scanned_senders = null;
      else
        bgwProcess.ReportProgress(0, $"Getting unique emails in folder took: {stopWatch.ElapsedMilliseconds} ms.");
    }

    private void bgwProcess_ProgressChanged(object sender, ProgressChangedEventArgs e)
    {
      if (e.ProgressPercentage >= 1 && e.ProgressPercentage <= 100)
        progressBar1.Value = e.ProgressPercentage;

      lblStatus.Text = e.UserState.ToString();
    }

    private void bgwProcess_RunWorkerCompleted(object sender, RunWorkerCompletedEventArgs e)
    {
      dgvList.Enabled = true;
      chkAll.Enabled = true;
      if (!chkAll.Checked)
      {
        cmbPeriod.Enabled = true;
        numScan.Enabled = true;
      }

      btnProcess.Enabled = true;
      btnProcess.Text = "Start";

      if (!cancelled && scanned_senders != null)
      {
        btnSave.Enabled = true;
        lblFoundSenders.Text = string.Format("Found {0} unique sender/s. ", scanned_senders.Count());
      }
      else
        lblFoundSenders.Text = "Stopped.";
    }

    private void frmSyncRule_Load(object sender, EventArgs e)
    { InitLoad(); }

    private void chkAll_CheckedChanged(object sender, EventArgs e)
    {
      cmbPeriod.Enabled = !chkAll.Checked;
      numScan.Enabled = !chkAll.Checked;
    }

    private void btnSave_Click(object sender, EventArgs e)
    {
      Outlook.Rule rule = null;
      var name_idx = 0;
      var rulename = string.Empty;
      var stopWatch = System.Diagnostics.Stopwatch.StartNew();

      if (scanned_senders != null)
      {
        if (scanned_senders.Count > 0)
        {
          if (MessageBox.Show("Existing sender/s on this rule will be replace with " + scanned_senders.Count.ToString() + " unique sender/s.\n\nDo you want to continue???", "Confirm - FarCap Add-In",
              MessageBoxButtons.YesNo, MessageBoxIcon.Question, MessageBoxDefaultButton.Button2) == System.Windows.Forms.DialogResult.Yes)
          {
            try
            {
              lblStatus.Text = "Saving rule..."; 
              lblFoundSenders.Text = "Saving rule...";
              this.Refresh();

              //CLEAR RULE GROUP
              Globals.ThisAddIn.OutlookRules.ClearRuleGroups(rulename_prefix);

              for (var i = 0; i < scanned_senders.Count; i++)
              {
                //REMOVE EMAIL ADDRESS IN ANY OTHER RULE
                var match_email = Globals.ThisAddIn.OutlookRules.FarCapRuleSenders.Where(
                    item => item.sender_email.Equals(scanned_senders[i].sender.sender_email, StringComparison.OrdinalIgnoreCase) &&
                      !item.rulename.StartsWith(rulename_prefix, StringComparison.OrdinalIgnoreCase)).ToList();

                if (match_email!=null)
                {
                  foreach (var existing_sender in match_email)
                    Globals.ThisAddIn.OutlookRules.RemoveEmailFromRule(existing_sender.rulename, existing_sender.sender_email);
                }

                //RULE GROUP NAMING
                if ((i % Properties.Settings.Default.MaxRuleRecipients) == 0)
                {
                  name_idx++;
                  rulename = rulename_prefix + "_" + name_idx;
                  rule = Globals.ThisAddIn.OutlookRules.Create(rulename, Outlook.OlRuleType.olRuleReceive);
                  rule.Actions.MoveToFolder.Folder = (parent_folder);
                  rule.Actions.MoveToFolder.Enabled = true;
                }
                
                if (rule != null)
                {
                  rule.Conditions.From.Recipients.Add(scanned_senders[i].sender.sender_email);
                  rule.Conditions.From.Recipients.ResolveAll();
                  rule.Conditions.From.Enabled = true;

                  Globals.ThisAddIn.OutlookRules.FarCapRuleSenders.Add(new FarCapSender(rulename,
                    scanned_senders[i].sender.sender_email,
                    scanned_senders[i].sender.sender_name,
                    parent_folder.Name,
                    parent_folder.FolderPath));
                }
              }
              //SAVE                            
              Globals.ThisAddIn.OutlookRules.Save(true);
              stopWatch.Stop();

              lblFoundSenders.Text = "Rule was updated!";
              lblRuleSenders.Text = $"{scanned_senders.Count} unique sender {(scanned_senders.Count > 1 ? "s" : "")}.";
              lblStatus.Text = $"Saving rule took: {stopWatch.ElapsedMilliseconds} ms.";

              scanned_senders = null;
              btnSave.Enabled = false;

              MessageBox.Show("Done!","FarCap Add-In",MessageBoxButtons.OK,MessageBoxIcon.Information);
            }
            catch (Exception ex)
            { 
              Globals.ThisAddIn.Error_Sender.SendNotification("@btnSave >> " + ex.Message + ex.StackTrace);
              lblStatus.Text = "Error saving rule.";
            }
          }
        }
      }
    }
  }
}
