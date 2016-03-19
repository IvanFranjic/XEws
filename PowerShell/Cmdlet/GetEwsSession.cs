namespace XEws.PowerShell.Cmdlet
{
    using System.Management.Automation;

    [Cmdlet(VerbsCommon.Get, "EwsSession")]
    public class GetEwsSession : EwsSessionCmdlet
    {
        #region Properties
        
        #endregion

        #region Fields

        #endregion

        #region Constructors

        #endregion

        #region Public Methods

        #endregion

        #region Private Methods

        #endregion

        #region Override Methods

        protected override void ProcessRecord()
        {
            WriteObject(this.EwsSession);
        }

        #endregion
    }
}
