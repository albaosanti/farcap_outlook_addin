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
    public partial class frmEditCategory : Form
    {
        public List<DataGridViewRow> selected_emails = null;
        
        public frmEditCategory()
        {
            InitializeComponent();
        }
        private void frmEditTarget_Load(object sender, EventArgs e)
        {
            initList();
            LoadCategory();
        }

        #region Functions & Procedures
       
        private void initList()
        {
            lblCount.Text = selected_emails.Count.ToString();
        }
        private void LoadCategory()
        {
            try
            {
                cmbTarget.Items.Clear();
                foreach (Outlook.Category _cat in Globals.ThisAddIn.Application.Session.Categories)
                    cmbTarget.Items.Add(_cat.Name);
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message + ex.StackTrace, "Error Loading Drag & Drop AddIn");
            }
        }

       
        #endregion

        private void btnClose_Click(object sender, EventArgs e)
        {
            this.DialogResult = System.Windows.Forms.DialogResult.Cancel;
            this.Close();
        }

        private void btnChange_Click(object sender, EventArgs e)
        {
            string tar_rulename = "";
            string src_rulename = "";

            string sender_address = "";
            string recipient_address;

            bool eadd_exist = false;
            bool has_changed = false;

            if (cmbTarget.SelectedIndex > -1)
            {
                if (MessageBox.Show("Are you to change the category to " + cmbTarget.Text + "?", 
                    "Confirm", MessageBoxButtons.YesNo, MessageBoxIcon.Question) == System.Windows.Forms.DialogResult.Yes)
                {
                    try
                    {
                        tar_rulename = Globals.ThisAddIn.CAT_RULE_PREFIX + cmbTarget.Text;
                        Outlook.Rule tar_rule = Globals.ThisAddIn.fnFindRuleByName(tar_rulename);
                        Outlook.Rule src_rule = null;
                        
                        //CREATE RULE 
                        if (tar_rule == null)
                        {
                            //CREATE NEW RULE
                            tar_rule = Globals.ThisAddIn.GlobalRules.Create(tar_rulename, Outlook.OlRuleType.olRuleReceive);
                            tar_rule.Actions.AssignToCategory.Categories = new object[] {cmbTarget.Text};
                            tar_rule.Actions.AssignToCategory.Enabled = true;
                        }

                        //CHECK EACH SENDER_ADDRESS
                        foreach (DataGridViewRow row in selected_emails)
                        {
                            sender_address = row.Cells[1].Value.ToString().Trim();//email address
                            src_rulename = row.Cells[3].Value.ToString();//source rule

                            eadd_exist = false;

                            if (sender_address != "" &&
                                    row.Cells[3].Value.ToString().ToLower() !=
                                        cmbTarget.Text.ToLower())
                            {
                                //DELETE THE EMAIL FROM THE PREVIOUS RULE
                                src_rule = Globals.ThisAddIn.fnFindRuleByName(src_rulename);
                                if (src_rule != null)
                                {
                                    foreach (Outlook.Recipient _recipient in src_rule.Conditions.From.Recipients)
                                    {
                                        recipient_address = Globals.ThisAddIn.fnGetSenderAddress(_recipient);
                                        if (!string.IsNullOrEmpty(recipient_address))
                                        {
                                            if (recipient_address.ToLower() == sender_address.ToLower())
                                            {
                                                _recipient.Delete();
                                                _recipient.Resolve();
                                                has_changed = true;
                                                break;
                                            }
                                        }
                                    }
                                    if (src_rule.Conditions.From.Recipients.Count == 0)
                                        Globals.ThisAddIn.GlobalRules.Remove(src_rulename);
                                }


                                //ADD THE EMAIL TO THE NEW RULE
                                if (tar_rule.Conditions.From.Recipients.Count > 0)
                                {
                                    foreach (Outlook.Recipient _recipient in tar_rule.Conditions.From.Recipients)
                                    {
                                        recipient_address = Globals.ThisAddIn.fnGetSenderAddress(_recipient);
                                        if (!string.IsNullOrEmpty(recipient_address))
                                        {
                                            eadd_exist = (recipient_address.ToLower() == row.Cells[1].Value.ToString().ToLower());
                                            if (eadd_exist) break;
                                        }
                                    }
                                }

                                //ADD FROM EMAIL
                                if (!eadd_exist)
                                {
                                    tar_rule.Conditions.From.Recipients.Add(sender_address);
                                    tar_rule.Conditions.From.Recipients.ResolveAll();
                                    tar_rule.Conditions.From.Enabled = true;
                                    has_changed = true;
                                }
                            }
                        }
                        if (has_changed && Globals.ThisAddIn.GlobalRules != null)
                        {
                            Globals.ThisAddIn.GlobalRules.Save(true);
                        }

                        this.DialogResult = System.Windows.Forms.DialogResult.OK;
                        this.Close();
                    }
                    catch (Exception ex)
                    {
                        MessageBox.Show(ex.Message + ex.StackTrace);
                    }                    
                }                
            }
        }
             
    }
}
