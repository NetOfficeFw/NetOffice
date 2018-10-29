using System;
using NUnit.Framework;
using NetOffice.OutlookApi.Tools;

namespace NetOffice.Tests.OutlookApi.Tools
{
    [TestFixture]
    public class OlCustomUIAttributeTests
    {
        /// <summary>
        /// Regression test for #223 (OlRibbonType.cs wrong enum for Microsoft.Outlook.Mail.Compose)
        /// </summary>
        [Test]
        public void OlRibbonType_MailComposeValue_IsCorrectlyConvertedToRibbonId()
        {
            // Arrange
            var attribute = new OlCustomUIAttribute(OlRibbonType.Microsoft_Outlook_Mail_Compose, "DummyValue.xml");

            // Act
            var actualRibbonId = attribute.RibbonID;

            // Assert
            Assert.AreEqual("Microsoft.Outlook.Mail.Compose", actualRibbonId);
        }
    }
}