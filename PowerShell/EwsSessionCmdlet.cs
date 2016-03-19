namespace XEws.PowerShell
{
    using System.Security;
    using System.Management.Automation;
    using Microsoft.Exchange.WebServices.Data;

    public class EwsSessionCmdlet : EwsCmdletAbstract
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

        /// <summary>
        /// Sets current ews session impersonatedEmail.
        /// </summary>
        /// <param name="impersonatedEmailAddress">Email address of user to impersonate.</param>
        internal void SetEwsSession(string impersonatedEmailAddress)
        {
            ExchangeService ewsSession = (ExchangeService)
                this.GetPowerShellVariable(EwsPowerShellVariable.ExchangeService).Value;

            // Are we clearing or setting impersonation?
            if (string.IsNullOrEmpty(impersonatedEmailAddress))
            {
                ewsSession.ImpersonatedUserId = null;
            }
            else
            {
                ewsSession.ImpersonatedUserId = new ImpersonatedUserId(ConnectingIdType.SmtpAddress, impersonatedEmailAddress);
                ewsSession.HttpHeaders.Add("X-AnchorMailbox", impersonatedEmailAddress);
            }

            this.SetPowerShellVariable(EwsPowerShellVariable.ExchangeService, ewsSession);
        }

        /// <summary>
        /// Validate redirect Autodiscover url. If it's https it will
        /// return true, otherwise it will return false.
        /// </summary>
        /// <param name="url">Autodiscover redirect URL</param>
        /// <returns>bool</returns>
        internal bool AutoDiscoverRedirectValidation(string url)
        {
            // TODO: Implement verbose strings.
            this.WriteVerbose("Autodiscover will be redirected to " + url);
            return url.StartsWith("https://");
        }

        #endregion

        #region Override Methods

        #endregion
    }
}
