namespace XEws.PowerShell.Cmdlet
{
    using System.Management.Automation;

    [Cmdlet(VerbsCommon.Remove, "EwsSession")]
    public class RemoveEwsSession : EwsSessionCmdlet
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
            this.RemovePowerShellVariable(EwsPowerShellVariable.ExchangeService);
        }

        #endregion
    }
}
