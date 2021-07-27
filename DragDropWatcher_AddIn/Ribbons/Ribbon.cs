using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Reflection;
using System.Runtime.InteropServices;
using System.Text;
using System.Windows.Forms;

using Outlook = Microsoft.Office.Interop.Outlook;
using Office = Microsoft.Office.Core;
using System.Collections.Specialized;



// TODO:  Follow these steps to enable the Ribbon (XML) item:

// 1: Copy the following code block into the ThisAddin, ThisWorkbook, or ThisDocument class.

//  protected override Microsoft.Office.Core.IRibbonExtensibility CreateRibbonExtensibilityObject()
//  {
//      return new Ribbon();
//  }

// 2. Create callback methods in the "Ribbon Callbacks" region of this class to handle user
//    actions, such as clicking a button. Note: if you have exported this Ribbon from the Ribbon designer,
//    move your code from the event handlers to the callback methods and modify the code to work with the
//    Ribbon extensibility (RibbonX) programming model.

// 3. Assign attributes to the control tags in the Ribbon XML file to identify the appropriate callback methods in your code.  

// For more information, see the Ribbon XML documentation in the Visual Studio Tools for Office Help.


namespace DragDrapWatcher_AddIn
{
  [ComVisible(true)]
  public class Ribbon : Office.IRibbonExtensibility
  {
    private Office.IRibbonUI ribbon;

    public Ribbon()
    {
    }

    #region IRibbonExtensibility Members

    public string GetCustomUI(string ribbonID)
    {
      return GetResourceText("DragDrapWatcher_AddIn.Ribbons.Ribbon.xml");
    }

    #endregion

    #region Ribbon Callbacks
    //Create callback methods here. For more information about adding callback methods, visit http://go.microsoft.com/fwlink/?LinkID=271226
    public bool mnuFarCapEnable(Office.IRibbonControl control)
    {
      Outlook.Folder folder = (Outlook.Folder)control.Context;
      return folder.Name.ToLower().StartsWith(Properties.Settings.Default.WatchFolder_Prefix);
    }

    public System.Drawing.Bitmap getSenderImage(Office.IRibbonControl control)
    {
      return Properties.Resources.user;
    }

    public System.Drawing.Bitmap getFarCapImage(Office.IRibbonControl control)
    {
      return Properties.Resources.farcap;
    }

    public string getContent_FarCapCategory(Office.IRibbonControl control)
    {
      Outlook.Categories cat_collection = null;
      string sub_buttons = "";
      int ctr = 1;

      sub_buttons = "<button id=\"btnCat_Manage\" label=\"Manage Rule Categories\" onAction=\"Controls_OnAction\" imageMso=\"CategorizeMenu\" />";
      sub_buttons += "<button id=\"btnCat_Clear\" label=\"Clear Selected\" onAction=\"btnCat_Clear_OnAction\" getEnabled=\"btnCatClear_getEnable\" imageMso=\"Clear\" />";
      sub_buttons += "<menuSeparator id=\"cat_separator\" />";

      cat_collection = Globals.ThisAddIn.Application.Session.Categories;
      foreach (Outlook.Category cat_item in cat_collection)
      {
        sub_buttons += "<checkBox id=\"chkfarcap_cat" + ctr + "\" label=\"" + cat_item.Name + "\" onAction=\"CheckBox_OnAction\" getPressed=\"getPressed_CheckBox\" tag=\"" + cat_item.Name + "\" />";
        ctr += 1;
      }

      return "<menu xmlns=\"http://schemas.microsoft.com/office/2009/07/customui\">"
        + sub_buttons
        + "</menu>";
    }

    public void btnCat_Clear_OnAction(Office.IRibbonControl control)
    {
      Outlook.MailItem item = null;
      Outlook.Selection selected = null;

      string categories = null;
      bool changed = false;

      if (control.Context is Outlook.Selection)
      {
        try
        {
          selected = (Outlook.Selection)control.Context;
          if (selected.Count > 0)
          {
            if (selected[1] is Outlook.MailItem)
            {
              item = (Outlook.MailItem)selected[1];
              categories = item.Categories;
              if (!string.IsNullOrEmpty(categories))
              {
                if (MessageBox.Show("Are you sure to REMOVE all the category rule/s from this sender?"
                    , "Confirm - FarCap", MessageBoxButtons.YesNo, MessageBoxIcon.Question, MessageBoxDefaultButton.Button2) == DialogResult.Yes)
                {
                  string[] arr = categories.Split(new char[] { ',' });
                  string rule_name = null;
                  string senderaddress = null;

                  foreach (string category in arr)
                  {

                    if (!string.IsNullOrWhiteSpace(category))
                    {
                      rule_name = Globals.ThisAddIn.CAT_RULE_PREFIX + category.Trim();
                      senderaddress = Globals.ThisAddIn.fnGetSenderAddress(item.Sender);
                      if (senderaddress != null)
                      {
                        Globals.ThisAddIn.fnRemoveMailCategory(ref categories, category.Trim());
                        Globals.ThisAddIn.OutlookRules.fnRemoveEmailFromRule(rule_name, senderaddress);
                        changed = true;
                      }
                    }
                  }

                  if (changed)
                  {
                    item.Categories = categories;
                    item.Save();
                    if (Globals.ThisAddIn.OutlookRules != null)
                      Globals.ThisAddIn.OutlookRules.Save(true);

                    MessageBox.Show("Selected category cleared!");
                  }
                }
              }
            }
          }
        }
        catch (Exception ex)
        { Globals.ThisAddIn.Error_Sender.SendNotification("@btnCat_Clear_OnAction >> " + ex.Message + ex.StackTrace); }
      }

    }

    public string getPressed_CheckBox(Office.IRibbonControl control)
    {
      Outlook.MailItem item = null;
      Outlook.Selection selected = null;

      string category = null;
      string pressed = "false";
      string category_name = control.Tag;

      if (control.Context is Outlook.Selection)
      {
        try
        {
          selected = (Outlook.Selection)control.Context;
          if (selected.Count > 0)
          {
            if (selected[1] is Outlook.MailItem)
            {
              item = (Outlook.MailItem)selected[1];
              category = item.Categories;
              if (!string.IsNullOrEmpty(category))
              {
                if (category.ToLower().Contains(category_name.ToLower()))
                  pressed = "true";
              }
            }
          }
        }
        catch (Exception ex)
        { Globals.ThisAddIn.Error_Sender.SendNotification("@InitContextCategory >> " + ex.Message + ex.StackTrace); }
      }

      return pressed;
    }

    public string btnCatClear_getEnable(Office.IRibbonControl control)
    {
      Outlook.MailItem item = null;
      Outlook.Selection selected = null;

      string category = null;
      string enable = "false";
      string category_name = control.Tag;

      if (control.Context is Outlook.Selection)
      {
        try
        {
          selected = (Outlook.Selection)control.Context;
          if (selected.Count > 0)
          {
            if (selected[1] is Outlook.MailItem)
            {
              item = (Outlook.MailItem)selected[1];
              category = item.Categories;
              if (!string.IsNullOrEmpty(category))
                enable = "true";
            }
          }
        }
        catch (Exception ex)
        { Globals.ThisAddIn.Error_Sender.SendNotification("@btnCatClear_getEnable >> " + ex.Message + ex.StackTrace); }
      }

      return enable;
    }

    public void CheckBox_OnAction(Office.IRibbonControl control, bool pressed)
    {
      Outlook.MailItem item = null;
      Outlook.Selection selected = null;

      string alert_ = null;
      string selected_category = null;
      string item_categories = null;
      string sender_address = null;
      string target_rulename = null;

      bool ok = false;

      if (control.Context is Outlook.Selection)
      {
        selected = (Outlook.Selection)control.Context;
        if (selected.Count > 0)
        {
          if (selected[1] is Outlook.MailItem)
          {
            try
            {
              selected_category = control.Tag;
              item = (Outlook.MailItem)selected[1];
              item_categories = item.Categories;


              if (pressed)
              {
                if (Globals.ThisAddIn.fnAppendToMailCategory(ref item_categories, selected_category))
                  alert_ = "ADD";
              }
              else
              {
                if (Globals.ThisAddIn.fnRemoveMailCategory(ref item_categories, selected_category))
                  alert_ = "REMOVE";
              }

              if (alert_ != null)
              {
                target_rulename = Globals.ThisAddIn.CAT_RULE_PREFIX + selected_category;
                sender_address = Globals.ThisAddIn.fnGetSenderAddress(item.Sender);

                if (MessageBox.Show("Are you sure to [" + alert_ + "] this sender from category rule?",
                  "FarCap Outlook Add-In", MessageBoxButtons.YesNo,
                  MessageBoxIcon.Question, MessageBoxDefaultButton.Button2) == DialogResult.Yes)
                {
                  item.Categories = item_categories;
                  item.Save();

                  if (alert_.Equals("ADD"))
                    ok = Globals.ThisAddIn.fnAddEmailToRule_Category(target_rulename, sender_address, selected_category);
                  else
                    ok = Globals.ThisAddIn.OutlookRules.fnRemoveEmailFromRule(target_rulename, sender_address);

                  //Save rules
                  if (Globals.ThisAddIn.OutlookRules != null && ok)
                  {
                    Globals.ThisAddIn.OutlookRules.Save(true);
                    MessageBox.Show("Done!");
                  }
                }
              }
            }
            catch (Exception ex)
            { Globals.ThisAddIn.Error_Sender.SendNotification("@CheckContextCategory >> " + ex.Message + ex.StackTrace); }

          }
        }
      }
    }

    public void Controls_OnAction(Office.IRibbonControl control)
    {
      switch (control.Id.ToLower())
      {
        /*RIBBON BUTTONS*/
        case "btnmanagesender":
          frmManager manager = new frmManager();
          manager.ShowDialog();
          break;
       
          /*MAIL FOLDER CONTEXT*/
        case "btnsyncrule":
          frmSyncRule sync = new frmSyncRule();
          sync.parent_folder = (Outlook.Folder)control.Context;
          sync.ShowDialog();
          break;
      
        case "btncat_manage":
          frmCategoryManager cat_manager = new frmCategoryManager();
          cat_manager.ShowDialog();
          break;
                  
        default:
          break;

      }
    }

    public void Ribbon_Load(Office.IRibbonUI ribbonUI)
    {
      this.ribbon = ribbonUI;
    }

    #endregion

    #region Helpers

    private static string GetResourceText(string resourceName)
    {
      Assembly asm = Assembly.GetExecutingAssembly();
      string[] resourceNames = asm.GetManifestResourceNames();
      for (int i = 0; i < resourceNames.Length; ++i)
      {
        if (string.Compare(resourceName, resourceNames[i], StringComparison.OrdinalIgnoreCase) == 0)
        {
          using (StreamReader resourceReader = new StreamReader(asm.GetManifestResourceStream(resourceNames[i])))
          {
            if (resourceReader != null)
            {
              return resourceReader.ReadToEnd();
            }
          }
        }
      }
      return null;
    }

    #endregion


  }
}
