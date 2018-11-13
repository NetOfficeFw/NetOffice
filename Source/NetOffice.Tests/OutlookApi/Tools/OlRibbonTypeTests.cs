using System;
using NUnit.Framework;
using NetOffice.OutlookApi.Tools;

namespace NetOffice.Tests.OutlookApi.Tools
{
    [TestFixture]
    public class OlRibbonTypeTests
    {
        [Test]
        [TestCase(OlRibbonType.Microsoft_Outlook_Appointment, "Microsoft.Outlook.Appointment")]
        [TestCase(OlRibbonType.Microsoft_Outlook_Contact, "Microsoft.Outlook.Contact")]
        [TestCase(OlRibbonType.Microsoft_Outlook_DistributionList, "Microsoft.Outlook.DistributionList")]
        [TestCase(OlRibbonType.Microsoft_Outlook_Journal, "Microsoft.Outlook.Journal")]
        [TestCase(OlRibbonType.Microsoft_Outlook_Mail_Compose, "Microsoft.Outlook.Mail.Compose")]
        [TestCase(OlRibbonType.Microsoft_Outlook_Mail_Read, "Microsoft.Outlook.Mail.Read")]
        [TestCase(OlRibbonType.Microsoft_Outlook_MeetingRequest_Read, "Microsoft.Outlook.MeetingRequest.Read")]
        [TestCase(OlRibbonType.Microsoft_Outlook_MeetingRequest_Send, "Microsoft.Outlook.MeetingRequest.Send")]
        [TestCase(OlRibbonType.Microsoft_Outlook_Post_Compose, "Microsoft.Outlook.Post.Compose")]
        [TestCase(OlRibbonType.Microsoft_Outlook_Post_Read, "Microsoft.Outlook.Post.Read")]
        [TestCase(OlRibbonType.Microsoft_Outlook_Report, "Microsoft.Outlook.Report")]
        [TestCase(OlRibbonType.Microsoft_Outlook_Resend, "Microsoft.Outlook.Resend")]
        [TestCase(OlRibbonType.Microsoft_Outlook_Response_Compose, "Microsoft.Outlook.Response.Compose")]
        [TestCase(OlRibbonType.Microsoft_Outlook_Response_CounterPropose, "Microsoft.Outlook.Response.CounterPropose")]
        [TestCase(OlRibbonType.Microsoft_Outlook_Response_Read, "Microsoft.Outlook.Response.Read")]
        [TestCase(OlRibbonType.Microsoft_Outlook_RSS, "Microsoft.Outlook.RSS")]
        [TestCase(OlRibbonType.Microsoft_Outlook_Sharing_Compose, "Microsoft.Outlook.Sharing.Compose")]
        [TestCase(OlRibbonType.Microsoft_Outlook_Sharing_Read, "Microsoft.Outlook.Sharing.Read")]
        [TestCase(OlRibbonType.Microsoft_Outlook_Task, "Microsoft.Outlook.Task")]
        [TestCase(OlRibbonType.Microsoft_Outlook_Explorer, "Microsoft.Outlook.Explorer")]
        public void OlRibbonType_MemberValue_IsCorrectlyConvertedToRibbonID(OlRibbonType ribbonType, string expectedValue)
        {
            // Arrange
            var attribute = new OlCustomUIAttribute(ribbonType, "DummyValue.xml");

            // Act
            var actualValue = attribute.RibbonID;

            // Assert
            Assert.AreEqual(expectedValue, actualValue);
        }
    }
}