namespace XEws.PowerShell.Cmdlet
{
    using System.Management.Automation;
    using Microsoft.Exchange.WebServices.Data;

    [Cmdlet(VerbsCommon.Get, "EwsExtendedPropertySet")]
    public class GetEwsExtendedPropertySet : EwsCmdletAbstract
    {
        #region Properties

        [Parameter(Mandatory = true)]
        public int PropertyId
        {
            get;
            set;
        }

        [Parameter(Mandatory = true)]
        public MapiPropertyType MapiProperty
        {
            get;
            set;
        }
        
        [Parameter(Mandatory = false, ParameterSetName = "DefaultExtendedProperty")]
        public DefaultExtendedPropertySet DefaultExtendedProperty
        {
            get;
            set;
        }


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
            ExtendedPropertyDefinition extendedProperty;

            switch (this.ParameterSetName)
            {
                case "DefaultExtendedProperty":
                    extendedProperty = new ExtendedPropertyDefinition(this.DefaultExtendedProperty, this.PropertyId, this.MapiProperty);
                    break;

                default:
                    extendedProperty = new ExtendedPropertyDefinition(this.PropertyId, this.MapiProperty);
                    break;
            }
            
            this.WriteObject(extendedProperty);
        }

        #endregion
    }
}
