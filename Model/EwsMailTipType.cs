namespace XEws.Model
{
    public enum EwsMailTipType
    {
        // https://msdn.microsoft.com/en-us/library/office/dd899452%28v=exchg.150%29.aspx

        OutOfOfficeMessage,
        All,
        MailboxFullStatus,
        CustomMailTip,
        ExternalMemberCount,
        TotalMemberCount,
        MaxMessageSize,
        DeliveryRestriction,
        ModerationStatus,
        InvalidRecipient,
    }
}
