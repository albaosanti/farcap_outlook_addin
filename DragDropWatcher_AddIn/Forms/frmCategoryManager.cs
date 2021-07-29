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
  public partial class frmCategoryManager : Form
  {
    private delegate void del_AddRow(object[] column_values);

    private List<myWatchEmail> watch_list = new List<myWatchEmail>();
    #region Classes
    private class myWatchEmail
    {
      public string target_categoryname;
      public string email_add;
      public string email_name;
      public string rule_name;

      public myWatchEmail() { }
      public myWatchEmail(string _categoryname, string _eadd, string _ename, string _rule)
      {
        this.target_categoryname = _categoryname;
        this.email_add = _eadd;
        this.rule_name = _rule;
        this.email_name = _ename;
      }
    }

    private void UpdateWatchList(bool reload_rules = false)
    {
      string categories = string.Empty;
      string rule_prefix = Globals.ThisAddIn.CAT_RULE_PREFIX.Trim().ToLower();

      watch_list = new List<myWatchEmail>();

      try
      {
        if (Globals.ThisAddIn.OutlookRules == null || reload_rules)
          Globals.ThisAddIn.OutlookRules.Reload();

        if (Globals.ThisAddIn.OutlookRules != null)
        {
          foreach (Outlook.Rule rule in Globals.ThisAddIn.OutlookRules.Rules)
          {
            if (rule.Name.ToLower().StartsWith(rule_prefix))
            {
              if (rule.RuleType == Outlook.OlRuleType.olRuleReceive
                       && rule.Conditions.From.Recipients != null)
              {
                categories = string.Empty;

                foreach (string _cat in rule.Actions.AssignToCategory.Categories)
                {
                  if (categories != string.Empty)
                    categories += ", ";

                  categories += _cat;
                }

                foreach (Outlook.Recipient _recipient in rule.Conditions.From.Recipients)
                {
                  myWatchEmail watch_email = new myWatchEmail();
                  watch_email.target_categoryname = categories;

                  watch_email.email_add = Globals.ThisAddIn.fnGetSenderAddress(_recipient);
                  if (!string.IsNullOrWhiteSpace(watch_email.email_add) &&
                      !string.IsNullOrWhiteSpace(watch_email.target_categoryname))
                  {
                    watch_email.email_name = _recipient.Name;
                    watch_email.rule_name = rule.Name;
                    watch_list.Add(watch_email);
                  }

                }
              }
            }
          }
        }
      }
      catch (Exception ex)
      {
        MessageBox.Show(ex.Message + ex.StackTrace, "FarCap Outlook Add-In");
      }

    }

    public bool DeleteWatchItem(string email_add, string rule_name)
    {
      bool rem = false;
      for (int i = 0; i < watch_list.Count; i++)
      {
        if (watch_list[i].rule_name.ToLower() == rule_name.ToLower() &&
            watch_list[i].email_add.ToLower() == email_add.ToLower())
        {
          watch_list.RemoveAt(i);
          rem = true;
          break;
        }
      }
      return rem;
    }

    #endregion

    private void AddRow(object[] column_values)
    { dgvList.Rows.Add(column_values); }

    public frmCategoryManager()
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
          foreach (myWatchEmail em in watch_list)
          {
            match = (checkedListBox1.GetItemChecked(0) &&
                    em.email_name.ToLower().Contains(key_word));


            if (!match)
            {
              match = (checkedListBox1.GetItemChecked(1) &&
                em.email_add.ToLower().Contains(key_word));
            }

            if (!match)
            {
              match = (checkedListBox1.GetItemChecked(2) &&
                em.target_categoryname.ToLower().Contains(key_word));
            }

            if (match)
            {
              dgvList.Rows.Add(new object[]{ em.email_name ,
                                em.email_add,
                                em.target_categoryname,
                                em.rule_name});
            }
          }
        }
        catch (Exception ex)
        {
          MessageBox.Show(ex.Message + ex.StackTrace, "FarCap Outlook Add-In");
        }
      }
      else
      {
        foreach (myWatchEmail em in watch_list)
        {
          dgvList.Rows.Add(new object[]{ em.email_name ,
                                em.email_add,
                                em.target_categoryname,
                                em.rule_name});
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
                  itm.Cells[3].Value.ToString(),
                  itm.Cells[1].Value.ToString()))
              {
                remove_count += 1;
                DeleteWatchItem(itm.Cells[1].Value.ToString(), itm.Cells[3].Value.ToString());
              }
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



    private void btnAdd_Click(object sender, EventArgs e)
    {

    }

    private void btnEdit_Click(object sender, EventArgs e)
    {
      if (dgvList.SelectedRows.Count > 0)
      {
        var cat_edit = new frmEditCategory();
        cat_edit.selected_emails = new List<DataGridViewRow>();

        foreach (DataGridViewRow itm in dgvList.SelectedRows)
          cat_edit.selected_emails.Add(itm);

        if (cat_edit.ShowDialog() == System.Windows.Forms.DialogResult.OK)
          btnRefresh.PerformClick();
      }
    }

    private void backgroundWorker1_DoWork(object sender, DoWorkEventArgs e)
    {
      UpdateWatchList(true);
      foreach (myWatchEmail em in watch_list)
      {
        this.Invoke(new del_AddRow(AddRow),
         new object[] { new object[]{ em.email_name ,
                        em.email_add,
                        em.target_categoryname,
                        em.rule_name}});
      }
    }

    private void backgroundWorker1_RunWorkerCompleted(object sender, RunWorkerCompletedEventArgs e)
    {
      btnDelete.Enabled = true;
      btnEdit.Enabled = true;
      btnRefresh.Enabled = true;
      lblStatus.Text = "[" + watch_list.Count + "] email account/s on watch list.";
    }
  }
}
