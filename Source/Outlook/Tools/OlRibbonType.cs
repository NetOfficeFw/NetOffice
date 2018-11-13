using System;

namespace NetOffice.OutlookApi.Tools
{
    /// <summary>
    /// Outlook Ribbon Types
    /// </summary>
    /// <remarks>
    /// https://docs.microsoft.com/en-us/office/vba/outlook/how-to/office-fluent-ui-extensibility/implementing-the-iribbonextensibility-interface
    /// </remarks>
    public enum OlRibbonType
    {
        /// <summary>
        /// Microsoft.Outlook.Appointment
        /// </summary>
        Microsoft_Outlook_Appointment = 0,

        /// <summary>
        /// Microsoft.Outlook.Contact
        /// </summary>
        Microsoft_Outlook_Contact = 1,

        /// <summary>
        /// Microsoft.Outlook.DistributionList
        /// </summary>
        Microsoft_Outlook_DistributionList = 2,

        /// <summary>
        /// Microsoft.Outlook.Explorer
        /// </summary>
        Microsoft_Outlook_Explorer = 3,

        /// <summary>
        /// Microsoft.Outlook.Journal
        /// </summary>
        Microsoft_Outlook_Journal = 4,

        /// <summary>
        /// Microsoft.Outlook.Mail.Compose
        /// </summary>
        Microsoft_Outlook_Mail_Compose = 5,

        /// <summary>
        /// Microsoft.Outlook.Mail.Read
        /// </summary>
        Microsoft_Outlook_Mail_Read = 6,

        /// <summary>
        /// Microsoft.Outlook.MeetingRequest.Read
        /// </summary>
        Microsoft_Outlook_MeetingRequest_Read = 7,

        /// <summary>
        /// Microsoft.Outlook.MeetingRequest.Send
        /// </summary>
        Microsoft_Outlook_MeetingRequest_Send = 8,

        /// <summary>
        /// Microsoft.Outlook.Post.Compose
        /// </summary>
        Microsoft_Outlook_Post_Compose = 9,

        /// <summary>
        /// Microsoft.Outlook.Post.Read
        /// </summary>
        Microsoft_Outlook_Post_Read = 10,

        /// <summary>
        /// Microsoft.Outlook.Report
        /// </summary>
        Microsoft_Outlook_Report = 11,

        /// <summary>
        /// Microsoft.Outlook.Resend
        /// </summary>
        Microsoft_Outlook_Resend = 12,

        /// <summary>
        /// Microsoft.Outlook.Response.Compose
        /// </summary>
        Microsoft_Outlook_Response_Compose = 13,

        /// <summary>
        /// Microsoft.Outlook.Response.CounterPropose
        /// </summary>
        Microsoft_Outlook_Response_CounterPropose = 14,

        /// <summary>
        /// Microsoft.Outlook.Response.Read
        /// </summary>
        Microsoft_Outlook_Response_Read = 15,

        /// <summary>
        /// Microsoft.Outlook.RSS
        /// </summary>
        Microsoft_Outlook_RSS = 16,

        /// <summary>
        /// Microsoft.Outlook.Sharing.Compose
        /// </summary>
        Microsoft_Outlook_Sharing_Compose = 17,

        /// <summary>
        /// Microsoft.Outlook.Sharing.Read
        /// </summary>
        Microsoft_Outlook_Sharing_Read = 18,

        /// <summary>
        /// Microsoft.Outlook.Task
        /// </summary>
        Microsoft_Outlook_Task = 19
    }
}
