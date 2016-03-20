namespace XEws.Cmdlet
{
    using System.Management.Automation;
    using Microsoft.Exchange.WebServices.Data;
    using XEws.CmdletAbstract;

    /*

        Cmdlet is returning Folder object of WellKnownFolderId. This 
        is usefull if one want get Inbox, SentItems, etc... folder
        object right away.

        Command is completed together with help.

    */
    // Removing command from the module, new set of Ews command will be
    // introduced.
    //[Cmdlet(VerbsCommon.Get, "XEwsWellKnownFolder")]
    public sealed class GetXEwsWellKnownFolder : XEwsCmdlet
    {
        [Parameter(Mandatory = true, Position = 0)]
        public WellKnownFolderName WellKnownFolder
        {
            get;
            set;
        }

        protected override void ProcessRecord()
        {
            Folder wellKnownFolder = XEwsFolderCmdlet.GetWellKnownFolder(this.WellKnownFolder, this.GetSessionVariable());

            WriteObject(wellKnownFolder);
        }
    }
}