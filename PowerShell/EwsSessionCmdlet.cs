namespace XEws.PowerShell
{
    using System;
    using System.Net;
    using XEws.Model;
    using System.Security;
    using System.Net.Security;
    using System.Management.Automation;
    using System.Text.RegularExpressions;
    using Microsoft.Exchange.WebServices.Data;
    using System.Security.Cryptography.X509Certificates;

    public class EwsSessionCmdlet : EwsCmdletAbstract
    {
        #region Properties

        private bool trustAllCertificates = false;
        [Parameter(Mandatory = false, Position = 9)]
        public bool TrustAllCertificates
        {
            get
            {
                return this.trustAllCertificates;
            }
            set
            {
                this.trustAllCertificates = value;
            }
        }

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

        /// <summary>
        /// Validate if connection to ssl site is secure (certificate is valid) and returns true
        /// or false, depending on result. At this point it will always return true which means
        /// that all certificates will be considered as valid.
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="certificateToVerify"></param>
        /// <param name="certificateChain"></param>
        /// <param name="sslErrors"></param>
        /// <returns></returns>
        internal bool ValidateServerCertificateCallBack(object sender, X509Certificate certificateToVerify, X509Chain certificateChain, SslPolicyErrors sslErrors)
        {
            // At this point only dont connect if certificate expired.

            DateTime certificateExpiration = GetDateFromString(certificateToVerify.GetExpirationDateString());

            // lets fail if certificate expired
            if (certificateExpiration < DateTime.Now)
                return false;

            return true;
        }

        /// <summary>
        /// Method tries to convert string to date time. If it's not able to parse string
        /// returns datetime +5 mins from now.
        /// </summary>
        /// <param name="dateTime">string representation that needs to be converted to datetime value</param>
        /// <returns></returns>
        internal DateTime GetDateFromString(string dateTime)
        {
            DateTime parsedResult;

            if (DateTime.TryParse(dateTime, out parsedResult))
            {
                return parsedResult;
            }

            // If we cannot parse date, just return 5 minutes
            // from now to avoid fail
            return DateTime.Now.AddMinutes(5);
        }

        #endregion

        #region Override Methods

        protected override void BeginProcessing()
        {
            // 
            if (this.trustAllCertificates)
            {
                if (this.ShouldContinue(
                    "You chose to trust all certificates. Please confirm before continuing. If you choose YES connection will trust all certificates.",
                    string.Empty
                ))
                    ServicePointManager.ServerCertificateValidationCallback = this.ValidateServerCertificateCallBack;
            }
                
        }

        #endregion
    }
}
