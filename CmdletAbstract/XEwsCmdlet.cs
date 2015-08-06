using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Management.Automation;
using Microsoft.Exchange.WebServices.Data;
using System.Net;
using XEws.Tracer;

namespace XEws.CmdletAbstract
{
    /*

        Class should be used as starting point to build
        cmdlets. Cmdlet shouldn't directly inherit from 
        this class, instead new more direct abstract class
        should be created which interits from this one.

    */
    public abstract class XEwsCmdlet : PSCmdlet
    {
        #region Properties

        internal ExchangeService EwsSession
        {
            get;
            private set;
        }

        #endregion
        

        #region Methods

        /// <summary>
        /// Method retrieves ExchangeService object from $EwsSession variable
        /// located in current powershell session.
        /// </summary>
        /// <returns>ExchangeService</returns>
        internal virtual ExchangeService GetSessionVariable()
        {
            PSVariable ewsSession = this.SessionState.PSVariable.Get("EwsSession");

            if (ewsSession == null)
                throw new InvalidOperationException("Could not find session. Please use Import-XEwsSession and try again.");

            ExchangeService session = (ExchangeService)ewsSession.Value;
            
            return session;
        }

        /// <summary>
        /// Sets powershell session $EwsSession variable with ExchangeService object.
        /// </summary>
        /// <param name="userName">Username of the user connecting to the Ews. It should be in email address format (UPN).</param>
        /// <param name="password">Password of the user connecting to the Ews.</param>
        /// <param name="ewsUrl">Ews url. It usualy looks like https://host.domain.com/EWS/Exchange.asmx</param>
        internal void SetSessionVariable(string userName, string password, Uri ewsUrl)
        {
            ValidateUserName(userName);

            ExchangeService ewsService = new ExchangeService();
            NetworkCredential credentials = new NetworkCredential(userName, password);

            ewsService.Url = ewsUrl;
            ewsService.Credentials = credentials;
            
            this.SessionState.PSVariable.Set("EwsSession", ewsService);
        }

        /// <summary>
        /// Method for setting $EwsSession session context variable.
        /// </summary>
        /// <param name="userName">Connecting username.</param>
        /// <param name="password">Connecting password.</param>
        /// <param name="ewsUrl">Ews endpoint.</param>
        /// <param name="impersonateEmail">Email address of the impersonated user.</param>
        /// <param name="traceEnabled">Enable ews tracing.</param>
        /// <param name="traceFolder">Location where trace items will be saved.</param>
        /// <param name="traceFlags">Options what to trace.</param>
        internal void SetSessionVariable(string userName, string password, Uri ewsUrl, string impersonateEmail, bool traceEnabled, string traceFolder, TraceFlags traceFlags)
        {
            this.SetSessionVariable(userName, password, ewsUrl);
            this.SetSessionVariable(impersonateEmail, traceEnabled, traceFolder, traceFlags);
        }

        /// <summary>
        /// Helper method for setting already imported session. 
        /// </summary>
        /// <param name="impersonateEmail">Email address of the impersonated user.</param>
        /// <param name="traceEnabled">Enable ews tracing.</param>
        /// <param name="traceFolder">Location where trace items will be saved.</param>
        /// <param name="traceFlags">Options what to trace.</param>
        internal void SetSessionVariable(string impersonateEmail, bool traceEnabled, string traceFolder, TraceFlags traceFlags)
        {
            if (!String.IsNullOrEmpty(impersonateEmail))
                this.SetSessionImpersonation(impersonateEmail);
            else
                this.SetSessionImpersonation(null);

            if (traceEnabled)
            {
                ExchangeService ewsService = this.GetSessionVariable();
                ewsService.TraceListener = new TraceListener(traceFolder);
                ewsService.TraceEnabled = traceEnabled;
                ewsService.TraceFlags = traceFlags;

                this.SessionState.PSVariable.Set("EwsSession", ewsService);
            }

        }

        /// <summary>
        /// Sets powershell session $EwsSession with ImpersonateUserId parameter.
        /// </summary>
        /// <param name="impersonateEmail">Email address of the Impersonating user.</param>
        internal void SetSessionImpersonation(string impersonateEmail)
        {
            ExchangeService ewsService = this.GetSessionVariable();

            if (impersonateEmail == null)
            {
                ewsService.ImpersonatedUserId = null;
                return;
            }

            ValidateUserName(impersonateEmail);

            if (!String.IsNullOrEmpty(impersonateEmail))
                ewsService.ImpersonatedUserId = new ImpersonatedUserId(ConnectingIdType.SmtpAddress, impersonateEmail);

        }

        /// <summary>
        /// Helper method for validating if the username is in correct UPN format.
        /// </summary>
        /// <param name="userName">username to check.</param>
        private void ValidateUserName(string userName)
        {
            try
            {
                EmailAddress upnTest = new EmailAddress(userName);
            }
            catch
            {
                throw new InvalidOperationException("Impersonate email is not in correct format.Please use user@domain.com format.");
            }
        }

        #endregion
    }
}
