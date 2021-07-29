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
using Office = Microsoft.Office.Core;

namespace DragDrapWatcher_AddIn
{
  public partial class frmManager : Form
  {
    private delegate void del_AddRow(object[] column_values);
    
    #region
    private void UpdateWatchList(bool reload_rules = false)
    {
      try
      {
        if(Globals.ThisAddIn.OutlookRules == null)
          Globals.ThisAddIn.OutlookRules = new GlobalRules(Globals.ThisAddIn.Application,Globals.ThisAddIn);
        else if (Globals.ThisAddIn.OutlookRules.Rules == null || 
            Globals.ThisAddIn.OutlookRules.FarCapRuleSenders==null || 
              reload_rules)
          Globals.ThisAddIn.OutlookRules.Reload();        
      }
      catch (Exception ex)
      { MessageBox.Show(ex.Message + ex.StackTrace, "FarCap Outlook Add-In"); }
    }

    private void AddRow(object[] column_values)
    { dgvList.Rows.Add(column_values); }
    #endregion

    public frmManager()
    {
      InitializeComponent();
    }

    private void btnSearch_Click(object sender, EventArgs e)
    {
      string key_word = textBox1.Text.Trim().ToLower();
      bool match = false;

      dgvList.Rows.Clear();
      lblStatus.Text = "Searching... Please wait.";
      this.Refresh();

      if (!string.IsNullOrWhiteSpace(key_word))
      {
        try
        {
          foreach (var farcapsender in Globals.ThisAddIn.OutlookRules.FarCapRuleSenders)
          {
            match = (checkedListBox1.GetItemChecked(0) &&
                    farcapsender.sender_name.ToLower().Contains(key_word)) || 
                    (checkedListBox1.GetItemChecked(1) && farcapsender.sender_email.ToLower().Contains(key_word)) || 
                    (checkedListBox1.GetItemChecked(2) &&farcapsender.folder_name.ToLower().Contains(key_word));

            if (match)
              dgvList.Rows.Add(new object[]{ farcapsender.sender_name,
                                farcapsender.sender_email,
                                farcapsender.folder_name,
                                farcapsender.folder_path,
                                farcapsender.rulename});
          }
        }
        catch (Exception ex)
        {
          MessageBox.Show(ex.Message + ex.StackTrace, "FarCap Outlook Add-In");
        }
      }
      else
      {
        foreach (var farcapsender in Globals.ThisAddIn.OutlookRules.FarCapRuleSenders)
        {
          dgvList.Rows.Add(new object[]{ farcapsender.sender_name,
                                farcapsender.sender_email,
                                farcapsender.folder_name,
                                farcapsender.folder_path,
                                farcapsender.rulename});
        }
      }
      lblStatus.Text = "[" + dgvList.RowCount + "] account match found.";
    }

    private void btnRefresh_Click(object sender, EventArgs e)
    {
      btnDelete.Enabled = false;
      btnEdit.Enabled = false;
      btnRefresh.Enabled = false;
      dgvList.Rows.Clear();
      lblStatus.Text = "Loading Rules.. Please wait.";
      backgroundWorker1.RunWorkerAsync();
    }

    private void checkedListBox1_SelectedIndexChanged(object sender, EventArgs e)
    {

    }

    private void textBox1_KeyDown(object sender, KeyEventArgs e)
    {
      if (e.KeyCode == Keys.Return) btnSearch.PerformClick();
    }

    private void linkLabel1_LinkClicked(object sender, LinkLabelLinkClickedEventArgs e)
    {
      for (int i = 0; i < checkedListBox1.Items.Count; i++)
        checkedListBox1.SetItemChecked(i, true);

    }

    private void frmManager_Load(object sender, EventArgs e)
    {
      linkLabel1_LinkClicked(sender, null);
      txtRuleName.Text = Properties.Settings.Default.RuleName_Prefix;
      txtFolder.Text = Properties.Settings.Default.WatchFolder_Prefix;
      txtRecipient.Text = Properties.Settings.Default.Recipient;
      txtCategoryPrefix.Text = Properties.Settings.Default.CategoryRulePrefix;
      numMaxRecipient.Value = Properties.Settings.Default.MaxRuleRecipients;

      btnRefresh.PerformClick();
    }

    private void btnDelete_Click(object sender, EventArgs e)
    {
      DataGridViewSelectedRowCollection selected_rows = dgvList.SelectedRows;
      if (selected_rows.Count > 0)
      {
        if (MessageBox.Show("Are you sure to DELETE the selected account [" + dgvList.SelectedRows.Count + "] on watch list?", "Confirm Delete - FarCap Outlook Add-In", MessageBoxButtons.YesNo, MessageBoxIcon.Question, MessageBoxDefaultButton.Button2) == System.Windows.Forms.DialogResult.Yes)
        {
          try
          {
            int remove_count = 0;
            for (int i = 0; i < selected_rows.Count; i++)
            {
              DataGridViewRow itm = selected_rows[i];
              if (Globals.ThisAddIn.OutlookRules.RemoveEmailFromRule(
                  itm.Cells[4].Value.ToString(),
                  itm.Cells[1].Value.ToString()))
                remove_count++;
            }

            if (remove_count > 0)
            {
              Globals.ThisAddIn.OutlookRules.Save(true);
              MessageBox.Show("Deleted Email/s [" + remove_count + "] !", "FarCap Outlook Add-In");
              btnSearch.PerformClick();
            }
          }
          catch (Exception ex)
          {
            MessageBox.Show(ex.Message + ex.StackTrace, "Error @ Delete Email - FarCap Outlook Add-In");
          }

          lblStatus.Text = "[" + dgvList.RowCount + "] email account/s on watch list.";
        }
      }
    }

    private void button1_Click(object sender, EventArgs e)
    {
      string rule_prefix = txtRuleName.Text.Trim();
      string folder_prefix = txtFolder.Text.Trim();
      string err_recipients = txtRecipient.Text.Trim();
      string cat_prefix = txtCategoryPrefix.Text.Trim();
      int rulerecipient =  Convert.ToInt32(numMaxRecipient.Value);

      if (rule_prefix != string.Empty && folder_prefix != string.Empty && err_recipients != string.Empty && cat_prefix != string.Empty)
      {
        if (MessageBox.Show("Confirm to UPDATE the configuration?", "Confirm Update - FarCap Outlook Add-In", MessageBoxButtons.YesNo, MessageBoxIcon.Question, MessageBoxDefaultButton.Button2) == System.Windows.Forms.DialogResult.Yes)
        {
          Properties.Settings.Default.WatchFolder_Prefix = folder_prefix;
          Properties.Settings.Default.RuleName_Prefix = rule_prefix;
          Properties.Settings.Default.Recipient = err_recipients;
          Properties.Settings.Default.CategoryRulePrefix = cat_prefix;
          Properties.Settings.Default.MaxRuleRecipients = rulerecipient;

          Properties.Settings.Default.Save();

          Globals.ThisAddIn.CAT_RULE_PREFIX = cat_prefix;

          btnRefresh.PerformClick();
        }
      }
      else
        MessageBox.Show("All fields require!", "FarCap Outlook Add-In");

    }

    private void btnAdd_Click(object sender, EventArgs e)
    {

    }

    private void btnEdit_Click(object sender, EventArgs e)
    {
      if (dgvList.SelectedRows.Count > 0)
      {
        frmEditTarget f_edit = new frmEditTarget();
        f_edit.selected_emails = new List<DataGridViewRow>();

        foreach (DataGridViewRow itm in dgvList.SelectedRows)
          f_edit.selected_emails.Add(itm);

        if (f_edit.ShowDialog() == System.Windows.Forms.DialogResult.OK)
          btnRefresh.PerformClick();
      }
    }

    private void linkLabel2_LinkClicked(object sender, LinkLabelLinkClickedEventArgs e)
    {
      clsSendNotif err_notif = new clsSendNotif();
      if (err_notif.SendTestNotification("This is a test message.", txtRecipient.Text))
        MessageBox.Show("Sent!", "FarCap Add-In");
      else
        MessageBox.Show("Failed to send!", "FarCap Add-In");

    }

    private void backgroundWorker1_DoWork(object sender, DoWorkEventArgs e)
    {
      UpdateWatchList(true);
      foreach (var farcapsender in Globals.ThisAddIn.OutlookRules.FarCapRuleSenders)
      {
        this.Invoke(new del_AddRow(AddRow),
         new object[] { new object[]{ farcapsender.sender_name,
                                farcapsender.sender_email,
                                farcapsender.folder_name,
                                farcapsender.folder_path,
                                farcapsender.rulename}});
      }
    }

    private void backgroundWorker1_RunWorkerCompleted(object sender, RunWorkerCompletedEventArgs e)
    {
      btnDelete.Enabled = true;
      btnEdit.Enabled = true;
      btnRefresh.Enabled = true;
      lblStatus.Text = "["  + Globals.ThisAddIn.OutlookRules.FarCapRuleSenders.Count() + "] email account/s on watch list.";
    }

    private void numMaxRecipient_ValueChanged(object sender, EventArgs e)
    {

    }
  }
}
