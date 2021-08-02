using System;
using System.Collections.Generic;
using System.IO;
using System.Runtime.InteropServices;
using System.Text.RegularExpressions;
using System.Windows.Forms;

namespace DragDrapWatcher_AddIn
{
  public class ClsSendNotif
  {
    private string local_log_path = $"C:\\FarCap_Outlook_AddIn\\Error_{DateTime.Today:yyyyMMdd}.log";
    private const string Subject = "FarCap Outlook Add-In";

    public bool SendNotification(string str_message)
    {
      bool ok_sent = false;

      Microsoft.Office.Interop.Outlook.MailItem mail = null;
      Microsoft.Office.Interop.Outlook.Recipients mailRecipients = null;
      Microsoft.Office.Interop.Outlook.Recipient mailRecipient = null;

      List<string> recipients = Split_Recipients(Properties.Settings.Default.Recipient);
      string ex_msg = string.Empty;
      try
      {
        if (recipients.Count > 0)
        {
          mail = (Microsoft.Office.Interop.Outlook.MailItem)Globals.ThisAddIn.Application.CreateItem(Microsoft.Office.Interop.Outlook.OlItemType.olMailItem);
          mail.Subject = Subject + " Error";
          mail.Body = "An exception has occured in the code of add-in.";
          mail.Body += "\n\n" + str_message;
          mailRecipients = mail.Recipients;

          foreach (string eadd in recipients)
          {
            mailRecipient = mailRecipients.Add(eadd);
            mailRecipient.Resolve();
          }
          if (mailRecipient.Resolved)
          {
            ((Microsoft.Office.Interop.Outlook._MailItem)mail).Send();
            ok_sent = true;
          }
          else
          {
            ex_msg = "Unable to send the error notification";
          }
        }
        else
        {
          ex_msg = "No recipient.";
        }
      }
      catch (System.Exception ex)
      {
        ex_msg = ex.Message + ex.StackTrace;
      }
      finally
      {
        if (mailRecipient != null)
          Marshal.ReleaseComObject(mailRecipient);
        if (mailRecipients != null)
          Marshal.ReleaseComObject(mailRecipients);
        if (mail != null)
          Marshal.ReleaseComObject(mail);
      }

      if (!ok_sent)
        WriteLog(ex_msg, str_message);

      return ok_sent;

    }

    public bool SendTestNotification(string str_message, string str_recipients)
    {
      bool ok_sent = false;

      Microsoft.Office.Interop.Outlook.MailItem mail = null;
      Microsoft.Office.Interop.Outlook.Recipients mailRecipients = null;
      Microsoft.Office.Interop.Outlook.Recipient mailRecipient = null;

      List<string> recipients = Split_Recipients(str_recipients);
      string ex_msg = string.Empty;
      try
      {
        if (recipients.Count > 0)
        {
          mail = (Microsoft.Office.Interop.Outlook.MailItem)Globals.ThisAddIn.Application.CreateItem(Microsoft.Office.Interop.Outlook.OlItemType.olMailItem);
          mail.Subject = Subject + " Error";
          mail.Body = "An exception has occured in the code of add-in.";
          mail.Body += "\n\n" + str_message;
          mailRecipients = mail.Recipients;
          foreach (string eadd in recipients)
          {
            mailRecipient = mailRecipients.Add(eadd);
            mailRecipient.Resolve();
          }
          if (mailRecipient.Resolved)
          {
            ((Microsoft.Office.Interop.Outlook._MailItem)mail).Send();
            ok_sent = true;
          }
          else
          {
            ex_msg = "Unable to send the error notification";
          }
        }
        else
        {
          ex_msg = "No recipient.";
        }
      }
      catch (System.Exception ex)
      {
        ex_msg = ex.Message + ex.StackTrace;
      }
      finally
      {
        if (mailRecipient != null)
          Marshal.ReleaseComObject(mailRecipient);
        if (mailRecipients != null)
          Marshal.ReleaseComObject(mailRecipients);
        if (mail != null)
          Marshal.ReleaseComObject(mail);
      }

      if (!ok_sent)
        WriteLog(ex_msg, str_message);

      return ok_sent;
    }

    public void WriteLog(string ex_msg, string str_message)
    {

      StreamWriter writer = null;
      try
      {
        if (!Directory.Exists(Path.GetDirectoryName(local_log_path)))
          Directory.CreateDirectory(Path.GetDirectoryName(local_log_path));

        writer = new StreamWriter(local_log_path, true);
        if (!string.IsNullOrWhiteSpace(ex_msg))
        {
          writer.WriteLine("Unsend Notification Error: " + ex_msg);
        }
        writer.WriteLine($"Timestamp: {DateTime.Now:dd-MM-yyyy hh:mm:ss}, Message: {str_message}");
        writer.Close();
        writer.Dispose();
        writer = null;
      }
      catch (System.Exception ex)
      {
        MessageBox.Show("Failed to write log!\nException: " + ex.Message +
                        "\n\nMessage:" + str_message, "FarCap Outlook Add-in");
      }
      finally
      {
        if (writer != null)
        {
          writer.Close();
          writer.Dispose();
        }
      }
    }

    public bool IsValidEmailAdd(string email_add)
    {
      Regex regex = new Regex(@"^([a-zA-Z0-9_\-\.]+)@((\[[0-9]{1,3}\.[0-9]{1,3}\.[0-9]{1,3}\.)|(([a-zA-Z0-9\-]+\.)+))([a-zA-Z]{2,4}|[0-9]{1,3})(\]?)$",
        RegexOptions.CultureInvariant | RegexOptions.Singleline);

      if (string.IsNullOrWhiteSpace(email_add))
        return false;

      return regex.IsMatch(email_add.Trim());
    }

    private List<string> Split_Recipients(string str_recipients)
    {
      string[] sp = str_recipients.Split(new char[] { ';' });
      List<string> rec = new List<string>();

      foreach (string s1 in sp)
      {
        if (IsValidEmailAdd(s1))
        {
          rec.Add(s1.Trim());
        }
      }
      return rec;
    }
  }
}