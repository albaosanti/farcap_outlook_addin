namespace DragDrapWatcher_AddIn
{
  public class FarCapSender
  {
    public string rulename;
    public string sender_name;
    public string sender_email;
    public string folder_name;
    public string folder_path;
    public int rule_number;

    public FarCapSender(string _rulename, string _email, string _name, string _folder)
    {
      this.rulename = _rulename;
      this.sender_email = _email;
      this.sender_name = _name;
      this.folder_name = _folder;
      this.rule_number = 0;
      //GET FOLDER INDEX
      var idx = rulename.LastIndexOf('_');
      if (idx > -1 && idx < rulename.Length - 1)
        int.TryParse(rulename.Substring(idx + 1), out rule_number);
    }
  }
}