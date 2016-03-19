namespace XEws.PowerShell
{
    using System.Net.Mail;
    using System.Security;
    using System.Management.Automation;
    using Microsoft.Exchange.WebServices.Data;
    using Microsoft.Exchange.WebServices.Autodiscover;

    public abstract class EwsCmdletAbstract : PSCmdlet
    {
        #region Properties

        /// <summary>
        /// Returns EwsSession inside current runspace.
        /// </summary>
        internal ExchangeService EwsSession
        {
            get
            {
                PSVariable exchangeService = this.GetPowerShellVariable(EwsPowerShellVariable.ExchangeService);
                if (exchangeService == null)
                    return null;
                return (ExchangeService)exchangeService.Value;
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
        /// Sets global variable inside PowerShell session.
        /// </summary>
        /// <param name="variableName">Predefined variable names.</param>
        /// <param name="variableValue">Value that will be held by variable.</param>
        internal void SetPowerShellVariable(EwsPowerShellVariable variableName, object variableValue)
        {
            this.SessionState.PSVariable.Set(variableName.ToString(), variableValue);
        }

        /// <summary>
        /// Gets global variable from current PowerShell session. Returns
        /// null if no variable set.
        /// </summary>
        /// <param name="variableName">Name of variable that we want value from.</param>
        /// <returns>PSVariable</returns>
        internal PSVariable GetPowerShellVariable(EwsPowerShellVariable variableName)
        {
            return this.SessionState.PSVariable.Get(variableName.ToString());
        }

        /// <summary>
        /// Removes global variable from current PowerShell session.
        /// </summary>
        /// <param name="variableName">Name of the variable we want to delete.</param>
        internal void RemovePowerShellVariable(EwsPowerShellVariable variableName)
        {
            PSVariable variableToRemove = this.GetPowerShellVariable(variableName);
            this.SessionState.PSVariable.Remove(variableToRemove);
        }

        /// <summary>
        /// Validate if provided string is email address. Return void if true,
        /// otherwise throw error.
        /// </summary>
        /// <param name="emailAddress">Email address to validate.</param>
        internal void ValidateEmailAddress(string emailAddress)
        {
            try
            {
                MailAddress mailAddress = new MailAddress(emailAddress);
            }
            catch (System.Exception)
            {
                // TODO: implement Exception strings.
                throw new PSInvalidOperationException("Provided email address is not valid: " + emailAddress);
            }
        }

        #endregion

        #region Override Methods

        protected override void StopProcessing()
        {
            this.WriteWarning("Stopping command");
        }

        #endregion
    }
}
