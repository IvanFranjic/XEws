namespace XEws.PowerShell
{
    using System;

    public abstract class EwsCmdlet : EwsCmdletAbstract
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
        
        protected override void BeginProcessing()
        {
            // TODO: error strings
            // no point of executing module if session is not loaded.
            if (this.EwsSession == null)
                throw new InvalidOperationException("Session not loaded.");
        }

        #endregion
    }
}
