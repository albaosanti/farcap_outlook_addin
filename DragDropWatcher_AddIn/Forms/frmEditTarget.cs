using System;
using System.Collections.Generic;
using System.Reflection;
using System.Windows.Forms;

using Outlook = Microsoft.Office.Interop.Outlook;
using Office = Microsoft.Office.Core;

namespace DragDrapWatcher_AddIn
{
  public partial class frmEditTarget : Form
  {
    public List<DataGridViewRow> selected_emails = null;
    public List<string[]> ValidFolders = null;

    public frmEditTarget()
    {
      InitializeComponent();
    }
    private void frmEditTarget_Load(object sender, EventArgs e)
    {
      initList();
      LoadFolders();
    }

    #region Functions & Procedures
    // Returns Folder object based on folder path
    private Outlook.Folder fnGetFolder(string folderPath)
    {
      Outlook.Folder folder;
      string backslash = @"\";
      try
      {
        if (folderPath.StartsWith(@"\\"))
          folderPath = folderPath.Remove(0, 2);

        String[] folders = folderPath.Split(backslash.ToCharArray());
        folder = Globals.ThisAddIn.Application.Session.Folders[folders[0]] as Outlook.Folder;
        if (folder != null)
        {
          for (int i = 1; i <= folders.GetUpperBound(0); i++)
          {
            Outlook.Folders subFolders = folder.Folders;
            folder = subFolders[folders[i]] as Outlook.Folder;
            if (folder == null) return null;
          }
        }
        return folder;
      }
      catch { return null; }
    }

    private void initList()
    {
      lblCount.Text = selected_emails.Count.ToString();
    }
    private void LoadFolders()
    {
      try
      {
        Outlook._NameSpace outNS;
        Outlook.Application application = Globals.ThisAddIn.Application;
        //Get the MAPI namespace
        outNS = application.GetNamespace("MAPI");
        //Get UserName
        string profileName = outNS.CurrentUser.Name;

        Outlook.Folders folders = outNS.Folders;

        ValidFolders = new List<string[]>();
        cmbTarget.Items.Clear();

        if (folders.Count > 0)
          IterateFolder(folders);

      }
      catch (Exception ex)
      {
        MessageBox.Show(ex.Message + ex.StackTrace, "Error Loading Drag & Drop AddIn");
      }
    }

    private void IterateFolder(Outlook.Folders parent_folder)
    {
      foreach (Outlook.Folder sub_fldr in parent_folder)
      {
        if (sub_fldr.Name.ToLower() == "deleted items" ||
            sub_fldr.Name.StartsWith("Vault", StringComparison.InvariantCultureIgnoreCase) ||
            sub_fldr.Name.StartsWith("Public Folder", StringComparison.InvariantCultureIgnoreCase))
        {
          continue;
        }

        if (sub_fldr.Name.ToLower().StartsWith(Properties.Settings.Default.WatchFolder_Prefix.ToLower()))
        {
          ValidFolders.Add(new string[] { sub_fldr.Name, sub_fldr.FolderPath });
          cmbTarget.Items.Add(sub_fldr.Name);
        }

        if (sub_fldr.Folders.Count > 0)
          IterateFolder(sub_fldr.Folders);
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
      string folder_path = string.Empty;
      string tar_rulename = string.Empty;
      bool has_changed = false;
      string loggerPrefix = $"{this.GetType().Name}->{MethodBase.GetCurrentMethod().Name} ::";
      Globals.ThisAddIn.Error_Sender.WriteLog($"{loggerPrefix}  Triggered");

      if (cmbTarget.SelectedIndex > -1)
      {
        folder_path = ValidFolders[cmbTarget.SelectedIndex][1];

        if (MessageBox.Show("Are you to change the target folder to " + cmbTarget.Text + "?",
            "Confirm", MessageBoxButtons.YesNo, MessageBoxIcon.Question) == System.Windows.Forms.DialogResult.Yes)
        {
          try
          {
            Outlook.Folder tar_folder = fnGetFolder(folder_path);
            tar_rulename = Properties.Settings.Default.RuleName_Prefix + tar_folder.Name;
            foreach (DataGridViewRow row in selected_emails)
            {
              if (Globals.ThisAddIn.OutlookRules.AddEmailToRule(
                    tar_rulename,
                    row.Cells[1].Value.ToString().Trim(),
                    row.Cells[0].Value.ToString().Trim(),
                    tar_folder))
                has_changed = true;
            }
            if (has_changed && Globals.ThisAddIn.OutlookRules != null)
              Globals.ThisAddIn.OutlookRules.Save(true);

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
