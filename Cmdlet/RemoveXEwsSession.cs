namespace XEws.Cmdlet
{
    using System.Management.Automation;
    using XEws.CmdletAbstract;
    // Removing command from the module, new set of Ews command will be
    // introduced.
    //[Cmdlet( VerbsCommon.Remove , "XEwsSession" )]
    public sealed class RemoveXEwsSession : XEwsCmdlet
    {
        protected override void ProcessRecord ( )
        {
            try
            {
                var sessionVariable = this.GetSessionVariable( );
                WriteWarning( string.Format( "Removing session variable." ) );

                this.RemoveSessionVariable( );
            }
            catch ( System.Exception )
            {
            }
        }
        
        internal void RemoveSessionVariable()
        {
            this.SessionState.PSVariable.Remove( "EwsSession" );
        }
    }
}
