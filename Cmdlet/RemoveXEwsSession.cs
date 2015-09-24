namespace XEws.Cmdlet
{
    using System.Management.Automation;

    [Cmdlet( VerbsCommon.Remove , "XEwsSession" )]
    public sealed class RemoveXEwsSession : PSCmdlet
    {
        protected override void ProcessRecord ( )
        {
            if (this.SessionVariableExist())
            {
                WriteWarning( string.Format( "Removing session variable." ) );
                this.RemoveSessionVariable( );
            }
        }

        internal bool SessionVariableExist()
        {
            if ( this.SessionState.PSVariable.Get( "EwsSession" ) != null )
                return true;

            return false;
        }

        internal void RemoveSessionVariable()
        {
            this.SessionState.PSVariable.Remove( "EwsSession" );
        }
    }
}
