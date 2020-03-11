
namespace SaveOutlookAttachments
{
    internal static class Constants
    {
        internal const string MapiNamespace = "MAPI";
        internal const string Note = "IPM.Note";
        internal const string MultipartSignedNote = "IPM.Note.SMIME.MultipartSigned";

        internal const string PROPTAG = "http://schemas.microsoft.com/mapi/proptag/0x";
        internal const string PT_BOOLEAN = "000B";
        internal const string PT_STRING8 = "001E";
        internal const string PT_UNICODE = "001F";
        internal const string PR_TRANSPORT_MESSAGE_HEADERS = PROPTAG + "007D" + PT_STRING8;
        internal const string PR_ATTACH_CONTENT_ID = PROPTAG + "3712" + PT_STRING8;
        internal const string PR_ATTACH_HIDDEN = PROPTAG + "7FFE" + PT_BOOLEAN;
    }
}
