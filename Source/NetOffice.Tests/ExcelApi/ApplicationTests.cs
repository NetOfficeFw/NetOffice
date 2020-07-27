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
        public void HasEventRecipients_NameofEventNameParam_WithNoSubscribers_ReturnsFalse()
        {
            // Arrange

            // Act
            var actualValue = ExcelApp.HasEventRecipients(nameof(ExcelApp.SheetChangeEvent));

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

        [Test]
        public void HasEventRecipients_EventNameParam_NonExistingEventName_ThrowsArgumentException()
        {
            // Arrange

            // Act & Assert
            var actualException = Assert.Throws<ArgumentOutOfRangeException>(
                () => ExcelApp.HasEventRecipients("NonExistingEventName")
            );

            Assert.AreEqual("eventName", actualException.ParamName);
            Assert.AreEqual("NonExistingEventName", actualException.ActualValue);
        }

        [Test]
        public void GetEventRecipients_WithSingleSubscriber_ReturnsTheSubsriber()
        {
            // Arrange
            ExcelApp.SheetActivateEvent += OnSheetActivateEventHandler1;

            // Act
            var actualRecipient = ExcelApp.GetEventRecipients("SheetActivate");

            // Assert
            CollectionAssert.IsNotEmpty(actualRecipient);
            var recipient = actualRecipient[0];
            Assert.AreEqual("OnSheetActivateEventHandler1", recipient.Method.Name);
        }

        [Test]
        public void GetEventRecipients_EventNameParam_NonExistingEventName_ThrowsArgumentException()
        {
            // Arrange

            // Act & Assert
            var actualException = Assert.Throws<ArgumentOutOfRangeException>(
                () => ExcelApp.GetEventRecipients("NonExistingEventName")
            );

            Assert.AreEqual("eventName", actualException.ParamName);
            Assert.AreEqual("NonExistingEventName", actualException.ActualValue);
        }

        [Test]
        public void GetCountOfEventRecipients_WithNoSubscribers_ReturnsZero()
        {
            // Arrange
            // ExcelApp.SheetActivateEvent += OnSheetActivateEventHandler;

            // Act
            var actualCount = ExcelApp.GetCountOfEventRecipients("SheetActivate");

            // Assert
            Assert.AreEqual(0, actualCount);
        }

        [Test]
        public void GetCountOfEventRecipients_WithMultipleSubscribers_ReturnsZero()
        {
            // Arrange
            ExcelApp.SheetActivateEvent += OnSheetActivateEventHandler1;
            ExcelApp.SheetActivateEvent += OnSheetActivateEventHandler2;
            ExcelApp.SheetActivateEvent += OnSheetActivateEventHandler3;

            // Act
            var actualCount = ExcelApp.GetCountOfEventRecipients("SheetActivate");

            // Assert
            Assert.AreEqual(3, actualCount);
        }

        [Test]
        public void GetCountOfEventRecipients_EventNameParam_NonExistingEventName_ThrowsArgumentException()
        {
            // Arrange

            // Act & Assert
            var actualException = Assert.Throws<ArgumentOutOfRangeException>(
                () => ExcelApp.GetCountOfEventRecipients("NonExistingEventName")
            );

            Assert.AreEqual("eventName", actualException.ParamName);
            Assert.AreEqual("NonExistingEventName", actualException.ActualValue);
        }

        void OnSheetActivateEventHandler1(ICOMObject sh)
        {
            // noop
        }

        void OnSheetActivateEventHandler2(ICOMObject sh)
        {
            // noop
        }

        void OnSheetActivateEventHandler3(ICOMObject sh)
        {
            // noop
        }
    }
}
