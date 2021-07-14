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
        
        private string rule_name;
        private List<SenderData> scanned_senders;
        private bool cancelled = false;
        private int scan_period;
        #endregion 
                
        #region Methods
        private int CountUniqueSenders()
        {
            var cnt = 0;
            var rule = Globals.ThisAddIn.OutlookRules.fnFindRuleByName(rule_name);
            if (rule != null)
                cnt = rule.Conditions.From.Recipients.Count;

            return cnt;
        }

        private string getFilterScanToDate()
        {
            DateTime scan_to = DateTime.Now;
            string filter = "";
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

        private void initLoad()
        {
            rule_name = Properties.Settings.Default.RuleName_Prefix + parent_folder.Name;
            lblFolderName.Text = parent_folder.Name;
            lblRuleSenders.Text = string.Format("{0} unique sender/s.",CountUniqueSenders());            

            numScan.Value = 3;
            cmbPeriod.SelectedIndex = 0;

            cancelled = false;
            chkAll.Checked = false;

            btnProcess.Text = "Start";
            btnSave.Text = "Save";

            btnProcess.Enabled = true;
            btnSave.Enabled = false;
        }

        private void btnProcess_Click(object sender, EventArgs e)
        {
            if (btnProcess.Text.ToLower() == "start")
            {
                cancelled = false;
                btnProcess.Text = "Stop";
                lblFoundSenders.Text = "Scanning..";
                lblStatus.Text = "Started..";

                chkAll.Enabled = false;
                numScan.Enabled = false;
                cmbPeriod.Enabled = false;
                scan_period = cmbPeriod.SelectedIndex;
                                
                bgwProcess.RunWorkerAsync();
            }
            else
            {
                if (MessageBox.Show("Are you sure to STOP the running process?", "Confirm Exit",
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
            
            var stopWatch = System.Diagnostics.Stopwatch.StartNew();
            var uniqueEmails = new HashSet<string>();
            var progress = 0;
            var count = 0;
            
            try
            {
                scanned_senders = new List<SenderData>();
                items = parent_folder.GetTable(getFilterScanToDate());
                count = items.GetRowCount();

                items.Columns.RemoveAll();
                items.Columns.Add("SenderName");
                items.Columns.Add("SenderEmailAddress");
                items.Columns.Add("SenderEmailType");
                items.Columns.Add("ReceivedTime");

                for (int index = 0; index < count; index++)
                {
                    if (cancelled) break;
                    
                    var mailItem = (Outlook.Row)items.GetNextRow();
                    var name = (string)mailItem["SenderName"];
                    var senderType = (string)mailItem["SenderEmailType"];
                    var emailAddress = (string)mailItem["SenderEmailAddress"];

                    progress = Convert.ToInt32(((Convert.ToDouble(index+1) / Convert.ToDouble(count)) * 100));
                    bgwProcess.ReportProgress(progress, "Processing " + (index+1) + "/" + count);                   
                    
                    if (!string.IsNullOrEmpty(emailAddress) && uniqueEmails.Add(emailAddress))
                        scanned_senders.Add(new SenderData(parent_folder.Name, name, emailAddress, senderType));

                }

            }
            catch (Exception ex)
            {
                Globals.ThisAddIn.Error_Sender.SendNotification( "@bgwProcess_SyncRule>> " + ex.Message + ex.StackTrace);
                bgwProcess.ReportProgress(0, "Error: " + ex.Message);
                scanned_senders = null;
            }
            stopWatch.Stop();

            if (cancelled) 
                scanned_senders = null;
            else
                bgwProcess.ReportProgress(0, string.Format("Getting unique emails in folder took: {0} ms.", stopWatch.ElapsedMilliseconds));
        }

        private void bgwProcess_ProgressChanged(object sender, ProgressChangedEventArgs e)
        {
            if (e.ProgressPercentage >= 1 && e.ProgressPercentage <= 100)
                progressBar1.Value = e.ProgressPercentage;

            lblStatus.Text =  e.UserState.ToString(); 
        }

        private void bgwProcess_RunWorkerCompleted(object sender, RunWorkerCompletedEventArgs e)
        {
            chkAll.Enabled = true;
            numScan.Enabled = true;
            cmbPeriod.Enabled = true;
            btnProcess.Enabled = true;
            btnProcess.Text = "Start";

            if (!cancelled  && scanned_senders!=null)
            {
                btnSave.Enabled = true;
                lblFoundSenders.Text = string.Format("Found {0} unique sender/s. ", scanned_senders.Count());

                btnSave.PerformClick();
            }
            else
            {
                lblFoundSenders.Text = "Stopped.";
                if(cancelled) lblStatus.Text = "Stopped.";            
            }

            
        }

        private void frmSyncRule_Load(object sender, EventArgs e)
        {initLoad();}

        private void chkAll_CheckedChanged(object sender, EventArgs e)
        {
            cmbPeriod.Enabled = !chkAll.Checked;
            numScan.Enabled = !chkAll.Checked;
        }
        
        private void btnSave_Click(object sender, EventArgs e)
        {
            
            if (scanned_senders != null)
            {
                if (scanned_senders.Count > 0)
                {
                    if(MessageBox.Show("Existing sender/s on this rule will be replace with "+ scanned_senders.Count.ToString()+ " unique sender/s.\n\nDo you want to continue???","Confirm - FarCap",
                        MessageBoxButtons.YesNo,MessageBoxIcon.Question,MessageBoxDefaultButton.Button2)==System.Windows.Forms.DialogResult.Yes)
                    {
                        try
                        {
                            var rule = Globals.ThisAddIn.OutlookRules.fnFindRuleByName(rule_name);
                            if (rule != null)
                            {
                                Globals.ThisAddIn.OutlookRules.Remove(rule_name);
                                rule = null;
                            }
                            rule = Globals.ThisAddIn.OutlookRules.Create(rule_name, Outlook.OlRuleType.olRuleReceive);
                            rule.Actions.MoveToFolder.Folder = (parent_folder);
                            rule.Actions.MoveToFolder.Enabled = true;

                            //INSERT ALL EMAIL ADDRESS
                            foreach (var sndr in scanned_senders)
                            {
                                rule.Conditions.From.Recipients.Add(sndr.EmailAddress);
                                rule.Conditions.From.Recipients.ResolveAll();
                                rule.Conditions.From.Enabled = true;
                            }

                            //SAVE                            
                            Globals.ThisAddIn.OutlookRules.Save(true);
                            
                            lblFoundSenders.Text = "Please re-scan folder.";
                            lblRuleSenders.Text = string.Format("{0} unique sender/s.", scanned_senders.Count);
                            lblStatus.Text = "Rule was updated!";

                            scanned_senders = null;
                            btnSave.Enabled = false;

                            MessageBox.Show("Done!");
                        }
                        catch (Exception ex)
                        { Globals.ThisAddIn.Error_Sender.SendNotification("@btnSave >> " + ex.Message + ex.StackTrace);}                        
                    }
                }
            }
        }
    }
}
