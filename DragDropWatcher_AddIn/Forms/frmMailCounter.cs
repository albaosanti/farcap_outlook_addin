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
  public partial class frmMailCounter : Form
  {
    #region Variables
    public Outlook.Folder parent_folder;

    private delegate void delFillGrid();
    private List<GroupSender> scanned_senders;
    private int scan_period = 0;
    private bool cancelled = false;
    private class GroupSender
    {
      public SenderData sender;
      public int count;
      public GroupSender(SenderData _sender, int _count)
      {
        this.sender = _sender;
        this.count = _count;
      }
    }


    #endregion

    #region Functions & Procedures
    private void fnFillGrid()
    {
      dgvList.Rows.Clear();
      foreach (var row in scanned_senders)
        dgvList.Rows.Add(new object[] { row.sender.Name, row.sender.EmailAddress, row.count });
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

    public frmMailCounter()
    {
      InitializeComponent();
    }

    private void frmMailCounter_Load(object sender, EventArgs e)
    {
      if (parent_folder != null)
      {
        lblFolderName.Text = parent_folder.Name;
        btnProcess.Text = "Start";
        btnProcess.Enabled = true;
      }
      else
      {
        lblFolderName.Text = "<Not Set>";
        btnProcess.Enabled = false;
      }
      cmbPeriod.SelectedIndex = 0;
      lblSender.Text = "<0>";
      dgvList.Rows.Clear();

    }

    private void dgvList_CellContentClick(object sender, DataGridViewCellEventArgs e)
    {

    }

    private void chkAll_CheckedChanged(object sender, EventArgs e)
    {
      cmbPeriod.Enabled = !chkAll.Checked;
      numScan.Enabled = !chkAll.Checked;
    }

    private void btnProcess_Click(object sender, EventArgs e)
    {
      if (btnProcess.Text.Equals("Start"))
      {
        if (MessageBox.Show($"Proceed scanning mails from folder [{lblFolderName.Text}] ?",
            "FarCap - Mail Counter", MessageBoxButtons.YesNo,
            MessageBoxIcon.Question, MessageBoxDefaultButton.Button2) == DialogResult.Yes)
        {
          dgvList.Rows.Clear();
          dgvList.Enabled = false;
          lblSender.Text = "Processing..";
          btnProcess.Text = "Stop";

          chkAll.Enabled = false;
          cmbPeriod.Enabled = false;
          numScan.Enabled = false;
          scan_period = cmbPeriod.SelectedIndex;
          cancelled = false;

          bgProcess.RunWorkerAsync();
        }
      }
      else
      {
        if (MessageBox.Show("Cancel the running process?",
            "FarCap - Mail Counter", MessageBoxButtons.YesNo,
            MessageBoxIcon.Question, MessageBoxDefaultButton.Button2) == DialogResult.Yes)
        {
          btnProcess.Text = "Cancelling..";
          btnProcess.Enabled = false;
          cancelled = true;
        }
      }
    }

    private void bgProcess_DoWork(object sender, DoWorkEventArgs e)
    {
      Outlook.Table items;
      Outlook.Search result = null;

      var stopWatch = System.Diagnostics.Stopwatch.StartNew();
      var uniqueEmails = new HashSet<string>();
      var progress = 0;
      var count = 0;
      var search_scope=parent_folder.FolderPath;      

      try
      {
        //items = parent_folder.GetTable(getFilterScanToDate());
        Globals.ThisAddIn.Error_Sender.WriteLog("Scope>>", search_scope);

        bgProcess.ReportProgress(0, "Searching mails...");
        scanned_senders = new List<GroupSender>();
        result = Globals.ThisAddIn.Application.AdvancedSearch(search_scope, getFilterScanToDate(), true,"NewSearch");
        items = result.GetTable();
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

          progress = Convert.ToInt32(((Convert.ToDouble(index + 1) / Convert.ToDouble(count)) * 100));
          bgProcess.ReportProgress(progress, "Processing " + (index + 1) + "/" + count);

          if (!string.IsNullOrEmpty(emailAddress))
          {
            emailAddress = emailAddress.ToLower();
            var senderdata = new SenderData(parent_folder.Name, name, emailAddress, senderType);
            if (uniqueEmails.Add(emailAddress))
              scanned_senders.Add(new GroupSender(senderdata, 1));
            else
            {
              var idx = scanned_senders.FindIndex(itm => String.Equals(itm.sender.EmailAddress, emailAddress, StringComparison.OrdinalIgnoreCase));
              if (idx > -1)
                scanned_senders[idx].count++;
            }
          }
        }
        bgProcess.ReportProgress(0, "Populating grid..");
        scanned_senders = scanned_senders.OrderByDescending(itm => itm.count).ToList();
        this.Invoke(new delFillGrid(fnFillGrid));
      }
      catch (Exception ex)
      {
        Globals.ThisAddIn.Error_Sender.SendNotification("@bgwProcess_MailCount>> " + ex.Message + ex.StackTrace);
        bgProcess.ReportProgress(0, "Error: " + ex.Message);
        scanned_senders = null;
      }
      stopWatch.Stop();

      if (cancelled)
        scanned_senders = null;
      else
        bgProcess.ReportProgress(0, $"Getting mail count took: {stopWatch.ElapsedMilliseconds} ms.");
    }

    private void bgProcess_ProgressChanged(object sender, ProgressChangedEventArgs e)
    {
      if (e.ProgressPercentage >= 1 && e.ProgressPercentage <= 100)
        progressBar1.Value = e.ProgressPercentage;

      lblStatus.Text = e.UserState.ToString();
    }

    private void bgProcess_RunWorkerCompleted(object sender, RunWorkerCompletedEventArgs e)
    {
      dgvList.Enabled = true;
      chkAll.Enabled = true;
      if (!chkAll.Checked)
      {
        numScan.Enabled = true;
        cmbPeriod.Enabled = true;
      }
      btnProcess.Enabled = true;
      btnProcess.Text = "Start";

      if (!cancelled && scanned_senders != null)
        lblSender.Text = $"Found {scanned_senders.Count()} unique sender/s. ";
      else
        lblSender.Text = "Stopped.";
    }
  }
}
