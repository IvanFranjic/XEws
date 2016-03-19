namespace XEws.PowerShell
{
    using System;
    using XEws.Model;
    using System.Security;
    using System.Management.Automation;
    using System.Text.RegularExpressions;
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

        /// <summary>
        /// Validates which of provided fields will be used for username.
        /// </summary>
        /// <param name="emailAddress">emailAddress field.</param>
        /// <param name="userName">userName field.</param>
        /// <param name="userIdentity">returns identity between two.</param>
        internal void ValidateUserName(string emailAddress, string userName, out string userIdentity)
        {
            if (!string.IsNullOrEmpty(userName))
            {
                try
                {
                    this.ValidateEmailAddress(userName);
                }
                catch
                {
                    Regex samAccountNameFormat = new Regex(".+\\\\.+");
                    Match samAccountNameMatch = samAccountNameFormat.Match(userName);

                    if (!samAccountNameMatch.Success)
                        throw new InvalidOperationException(String.Format("Provided username '{0}' is not in correct format. Please use one of the following format: 'domain\\username' or 'username@domain.com'", userName));
                }
                userIdentity = userName;
            }
            else
                userIdentity = emailAddress;
        }

        #endregion

        #region Override Methods

        #endregion
    }
}
