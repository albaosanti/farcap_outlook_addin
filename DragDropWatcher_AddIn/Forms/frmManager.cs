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

        private List<myWatchEmail> watch_list = new List<myWatchEmail>();      
        #region Classes
        private class myWatchEmail
        {
            public string destination_folder;
            public string email_add;
            public string email_name;
            public string rule_name;
            public string folder_name;
            
            public myWatchEmail() { }
            public myWatchEmail(string _dest, string _eadd, string _ename, string _rule, string _foldername)
            {
                this.destination_folder = _dest;
                this.email_add = _eadd;
                this.rule_name = _rule;
                this.email_name = _ename;
                this.folder_name = _foldername;
            }
        }

        private void UpdateWatchList(bool reload_rules = false)
        {
          string rule_prefix = Properties.Settings.Default.RuleName_Prefix.ToLower().Trim();
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
                                   && rule.Actions.MoveToFolder != null
                                   && rule.Conditions.From.Recipients != null)
                          {
                              foreach (Outlook.Recipient rp in rule.Conditions.From.Recipients)
                              {
                                  myWatchEmail watch_email = new myWatchEmail();
                                  watch_email.destination_folder = rule.Actions.MoveToFolder.Folder.FolderPath;
                                  watch_email.folder_name = rule.Actions.MoveToFolder.Folder.Name;
                                  watch_email.email_add = Globals.ThisAddIn.fnGetSenderAddress(rp);
                                  if (watch_email.email_add != null)
                                  {
                                      watch_email.email_name = rp.Name;
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
            for (int i=0;i < watch_list.Count;i++)
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
        {dgvList.Rows.Add(column_values);}

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
            
            if(!string.IsNullOrWhiteSpace(key_word)){
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
                              em.folder_name.ToLower().Contains(key_word));
                        }
                       

                        if (match)
                        {
                             dgvList.Rows.Add(new object[]{ em.email_name ,
                                em.email_add,
                                em.folder_name, 
                                em.destination_folder,
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
                                em.folder_name, 
                                em.destination_folder,
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
            txtRuleName.Text = Properties.Settings.Default.RuleName_Prefix;
            txtFolder.Text = Properties.Settings.Default.WatchFolder_Prefix;
            txtRecipient.Text = Properties.Settings.Default.Recipient;
            txtCategoryPrefix.Text = Properties.Settings.Default.CategoryRulePrefix;
            btnRefresh.PerformClick();
        }

        private void btnDelete_Click(object sender, EventArgs e)
        {
            DataGridViewSelectedRowCollection selected_rows= dgvList.SelectedRows;
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
                            if ( Globals.ThisAddIn.OutlookRules.fnRemoveEmailFromRule(
                                itm.Cells[4].Value.ToString(),
                                itm.Cells[1].Value.ToString()))
                            {
                                remove_count += 1;
                                DeleteWatchItem(itm.Cells[1].Value.ToString(), itm.Cells[4].Value.ToString());
                            }
                        }

                        if (remove_count > 0)
                        {
                            Globals.ThisAddIn.OutlookRules.Save(true);
                            MessageBox.Show("Deleted Email/s [" + remove_count + "] !", "FarCap Outlook Add-In");
                            btnSearch.PerformClick();
                        }
                    }
                    catch ( Exception ex){
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

            if (rule_prefix != "" && folder_prefix != "" && err_recipients != "" && cat_prefix != "")
            {
                if (MessageBox.Show("Confirm to UPDATE the configuration?", "Confirm Update - FarCap Outlook Add-In", MessageBoxButtons.YesNo, MessageBoxIcon.Question, MessageBoxDefaultButton.Button2) == System.Windows.Forms.DialogResult.Yes)
                {
                    Properties.Settings.Default.WatchFolder_Prefix = folder_prefix;
                    Properties.Settings.Default.RuleName_Prefix = rule_prefix;
                    Properties.Settings.Default.Recipient = err_recipients;
                    Properties.Settings.Default.CategoryRulePrefix = cat_prefix;
                    
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
                MessageBox.Show("Sent!");
            else
                MessageBox.Show("Failed to send!");

        }

        private void backgroundWorker1_DoWork(object sender, DoWorkEventArgs e)
        {
            UpdateWatchList(true);
            foreach (myWatchEmail em in watch_list)
            {
                this.Invoke(new del_AddRow(AddRow),
                 new object[] { new object[]{ em.email_name ,
                        em.email_add,
                        em.folder_name, 
                        em.destination_folder,
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
