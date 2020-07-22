using System;
using NUnit.Framework;
using NetOffice.ExcelApi;

namespace NetOffice.Tests.ExcelApi
{
    [TestFixture]
    [Category("IntegrationTests")]
    [Category("IntegrationTests_Excel")]
    public class ApplicationTests
    {
        public Application ExcelApp { get; set; }

        [SetUp]
        public void SetUp()
        {
            this.ExcelApp = new Application();
        }

        [TearDown]
        public void TearDown()
        {
            this.ExcelApp?.Quit();
            this.ExcelApp?.Dispose();
        }

        [Test]
        public void IsWithEventRecipients_WithNoSubscribers_ReturnsFalse()
        {
            // Arrange

            // Act
            var actualValue = ExcelApp.IsWithEventRecipients;

            // Assert
            Assert.IsFalse(actualValue);
        }

        [Test]
        public void IsWithEventRecipients_WithASubscriber_ReturnsTrue()
        {
            // Arrange
            ExcelApp.SheetActivateEvent += sh => { };

            // Act
            var actualValue = ExcelApp.IsWithEventRecipients;

            // Assert
            Assert.IsTrue(actualValue);
        }

        [Test]
        public void HasEventRecipients_WithNoSubscribers_ReturnsFalse()
        {
            // Arrange

            // Act
            var actualValue = ExcelApp.HasEventRecipients();

            // Assert
            Assert.IsFalse(actualValue);
        }

        [Test]
        public void HasEventRecipients_WithASubscriber_ReturnsTrue()
        {
            // Arrange
            ExcelApp.SheetActivateEvent += sh => { };

            // Act
            var actualValue = ExcelApp.HasEventRecipients();

            // Assert
            Assert.IsTrue(actualValue);
        }

        [Test]
        public void HasEventRecipients_EventNameParam_WithNoSubscribers_ReturnsFalse()
        {
            // Arrange

            // Act
            var actualValue = ExcelApp.HasEventRecipients("SheetChange");

            // Assert
            Assert.IsFalse(actualValue);
        }

        [Test]
        public void HasEventRecipients_EventNameParam_WithASubscriberOnAnotherEvent_ReturnsFalse()
        {
            // Arrange
            ExcelApp.SheetActivateEvent += sh => { };

            // Act
            var actualValue = ExcelApp.HasEventRecipients("SheetChange");

            // Assert
            Assert.IsFalse(actualValue);
        }

        [Test]
        public void HasEventRecipients_EventNameParam_WithASubscriberOnTheNamedEvent_ReturnsTrue()
        {
            // Arrange
            ExcelApp.SheetActivateEvent += sh => { };

            // Act
            var actualValue = ExcelApp.HasEventRecipients("SheetActivate");

            // Assert
            Assert.IsTrue(actualValue);
        }
    }
}
