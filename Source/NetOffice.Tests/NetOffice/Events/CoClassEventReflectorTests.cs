using System;
using System.Collections.Generic;
using System.Text;
using NUnit.Framework;
using NetOffice.Events;

namespace NetOffice.Tests.NetOffice.Events
{
    [TestFixture]
    public class CoClassEventReflectorTests
    {
        [Test]
        public void HasEventRecipients_NoSubscriber_ReturnsFalse()
        {
            // Arrange
            var stub = new EventClassStub();

            // Act
            var actualValue = CoClassEventReflector.HasEventRecipients(stub, typeof(EventClassStub));

            // Assert
            Assert.IsFalse(actualValue);
        }

        [Test]
        public void HasEventRecipients_EventNameParam_NoSubscriber_ReturnsFalse()
        {
            // Arrange
            var stub = new EventClassStub();

            // Act
            var actualValue = CoClassEventReflector.HasEventRecipients(stub, typeof(EventClassStub), "Custom1");

            // Assert
            Assert.IsFalse(actualValue);
        }

        [Test]
        public void HasEventRecipients_EventNameParam_NoSubscriberForNamedEvent_ReturnsFalse()
        {
            // Arrange
            var stub = new EventClassStub();
            stub.Custom1Event += (obj, arg) => { };

            // Act
            var actualValue = CoClassEventReflector.HasEventRecipients(stub, typeof(EventClassStub), "Custom2");

            // Assert
            Assert.IsFalse(actualValue);
        }

        [Test]
        public void HasEventRecipients_EventWithSubscriber_ReturnsTrue()
        {
            // Arrange
            var stub = new EventClassStub();
            stub.Custom1Event += (sender, args) => { };

            // Act
            var actualValue = CoClassEventReflector.HasEventRecipients(stub, typeof(EventClassStub));

            // Assert
            Assert.IsTrue(actualValue);
        }

        [Test]
        public void HasEventRecipients_EventNameParam_EventWithSubscriber_ReturnsTrue()
        {
            // Arrange
            var stub = new EventClassStub();
            stub.Custom1Event += (sender, args) => { };

            // Act
            var actualValue = CoClassEventReflector.HasEventRecipients(stub, typeof(EventClassStub), "Custom1");

            // Assert
            Assert.IsTrue(actualValue);
        }

        [Test]
        public void HasEventRecipients_EventNameParam_NonExistingEventName_ThrowsArgumentException()
        {
            // Arrange
            var stub = new EventClassStub();

            // Act & Assert
            var actualException = Assert.Throws<ArgumentOutOfRangeException>(
                () => CoClassEventReflector.HasEventRecipients(stub, typeof(EventClassStub), "NonExistingEventName")
            );

            Assert.AreEqual("eventName", actualException.ParamName);
            Assert.AreEqual("NonExistingEventName", actualException.ActualValue);
        }

        [Test]
        public void GetEventRecipients_WithNoSubscribers_ReturnsEmptyCollection()
        {
            // Arrange
            var stub = new EventClassStub();

            // Act
            var actualRecipient = CoClassEventReflector.GetEventRecipients(stub, typeof(EventClassStub), "Custom1");

            // Assert
            Assert.IsNotNull(actualRecipient);
            CollectionAssert.IsEmpty(actualRecipient);
        }

        [Test]
        public void GetEventRecipients_WithSingleSubscriber_ReturnsTheSubsriber()
        {
            // Arrange
            var stub = new EventClassStub();
            stub.Custom1Event += OnCustom1EventHandler;

            // Act
            var actualRecipient = CoClassEventReflector.GetEventRecipients(stub, typeof(EventClassStub), "Custom1");

            // Assert
            CollectionAssert.IsNotEmpty(actualRecipient);
            var recipient = actualRecipient[0];
            Assert.AreEqual("OnCustom1EventHandler", recipient.Method.Name);
        }

        [Test]
        public void GetEventRecipients_WithMultipleSubscribers_ReturnsTheCorrectSubsriber()
        {
            // Arrange
            var stub = new EventClassStub();
            stub.Custom1Event += OnCustom1EventHandler;
            stub.Custom2Event += OnCustom2EventHandler;
            stub.Custom3Event += OnCustom3EventHandler;

            // Act
            var actualRecipient = CoClassEventReflector.GetEventRecipients(stub, typeof(EventClassStub), "Custom2");

            // Assert
            CollectionAssert.IsNotEmpty(actualRecipient);
            var recipient = actualRecipient[0];
            Assert.AreEqual("OnCustom2EventHandler", recipient.Method.Name);
        }

        [Test]
        public void GetEventRecipients_NonExistingEventName_ThrowsArgumentException()
        {
            // Arrange
            var stub = new EventClassStub();

            // Act & Assert
            var actualException = Assert.Throws<ArgumentOutOfRangeException>(
                () => CoClassEventReflector.GetEventRecipients(stub, typeof(EventClassStub), "NonExistingEventName")
            );

            Assert.AreEqual("eventName", actualException.ParamName);
            Assert.AreEqual("NonExistingEventName", actualException.ActualValue);
        }

        [Test]
        public void GetCountOfEventRecipients_NonExistingEventName_ThrowsArgumentException()
        {
            // Arrange
            var stub = new EventClassStub();

            // Act & Assert
            var actualException = Assert.Throws<ArgumentOutOfRangeException>(
                () => CoClassEventReflector.GetCountOfEventRecipients(stub, typeof(EventClassStub), "IncorrectEventName"));

            Assert.AreEqual("eventName", actualException.ParamName);
            Assert.AreEqual("IncorrectEventName", actualException.ActualValue);
        }

        [Test]
        public void GetCountOfEventRecipients_WithNoSubscribers_ReturnsZero()
        {
            // Arrange
            var stub = new EventClassStub();

            // Act
            var actualCount = CoClassEventReflector.GetCountOfEventRecipients(stub, typeof(EventClassStub), "Custom1");

            // Assert
            Assert.AreEqual(0, actualCount);
        }

        [Test]
        public void GetCountOfEventRecipients_WithMultipleSubscribers_ReturnsZero()
        {
            // Arrange
            var stub = new EventClassStub();
            stub.Custom1Event += OnCustom1EventHandler;

            stub.Custom2Event += OnCustom1EventHandler;
            stub.Custom2Event += OnCustom1EventHandler;
            
            stub.Custom3Event += OnCustom1EventHandler;
            stub.Custom3Event += OnCustom1EventHandler;
            stub.Custom3Event += OnCustom1EventHandler;

            // Act
            var actualCount = CoClassEventReflector.GetCountOfEventRecipients(stub, typeof(EventClassStub), "Custom2");

            // Assert
            Assert.AreEqual(2, actualCount);
        }

        private void OnCustom1EventHandler(object sender, EventArgs e)
        {
            // noop
        }

        private void OnCustom2EventHandler(object sender, EventArgs e)
        {
            // noop
        }

        private void OnCustom3EventHandler(object sender, EventArgs e)
        {
            // noop
        }
    }
}
