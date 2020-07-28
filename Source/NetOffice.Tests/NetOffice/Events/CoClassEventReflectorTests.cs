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
        public void HasEventRecipients_NameofEventNameParam_NoSubscriber_ReturnsFalse()
        {
            // Arrange
            var stub = new EventClassStub();

            // Act
            var actualValue = CoClassEventReflector.HasEventRecipients(stub, typeof(EventClassStub), nameof(EventClassStub.Custom1Event));

            // Assert
            Assert.IsFalse(actualValue);
        }

        [Test]
        public void HasEventRecipients_NameofEventNameParam_EventWithSubscriber_ReturnsTrue()
        {
            // Arrange
            var stub = new EventClassStub();
            stub.Custom1Event += (sender, args) => { };

            // Act
            var actualValue = CoClassEventReflector.HasEventRecipients(stub, typeof(EventClassStub), nameof(EventClassStub.Custom1Event));

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
        public void GetEventRecipients_NameofEventNameParam_NoSubscriber_ReturnsFalse()
        {
            // Arrange
            var stub = new EventClassStub();

            // Act
            var actualRecipient = CoClassEventReflector.GetEventRecipients(stub, typeof(EventClassStub), nameof(EventClassStub.Custom1Event));

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
        public void GetEventRecipients_NameofEventNameParam_EventWithSubscriber_ReturnsTrue()
        {
            // Arrange
            var stub = new EventClassStub();
            stub.Custom1Event += OnCustom1EventHandler;

            // Act
            var actualRecipient = CoClassEventReflector.GetEventRecipients(stub, typeof(EventClassStub), nameof(EventClassStub.Custom1Event));

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
        public void GetCountOfEventRecipients_NameofEventName_WithNoSubscribers_ReturnsZero()
        {
            // Arrange
            var stub = new EventClassStub();

            // Act
            var actualCount = CoClassEventReflector.GetCountOfEventRecipients(stub, typeof(EventClassStub), nameof(EventClassStub.Custom1Event));

            // Assert
            Assert.AreEqual(0, actualCount);
        }

        [Test]
        public void GetCountOfEventRecipients_WithMultipleSubscribers_ReturnsCorrectCount()
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

        [Test]
        public void GetCountOfEventRecipients_NameofEventName_WithMultipleSubscribers_ReturnsCorrectCount()
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
            var actualCount = CoClassEventReflector.GetCountOfEventRecipients(stub, typeof(EventClassStub), nameof(EventClassStub.Custom3Event));

            // Assert
            Assert.AreEqual(3, actualCount);
        }

        [Test]
        public void RaiseCustomEvent_WithNoSubscribers_ReturnsZero()
        {
            // Arrange
            var stub = new EventClassStub();
            var p = new object[] { };

            // Act
            var actualCount = CoClassEventReflector.RaiseCustomEvent(stub, typeof(EventClassStub), "Custom1", ref p);

            // Assert
            Assert.AreEqual(0, actualCount);
        }

        [Test]
        public void RaiseCustomEvent_WithTwoSubscribers_ReturnsTwo()
        {
            // Arrange
            var stub = new EventClassStub();
            stub.Custom2Event += OnCustom2EventHandler;
            stub.Custom2Event += OnCustom2EventHandler;

            var p = new object[] { this, new EventArgs() };

            // Act
            var actualCount = CoClassEventReflector.RaiseCustomEvent(stub, typeof(EventClassStub), "Custom2", ref p);

            // Assert
            Assert.AreEqual(2, actualCount);
        }

        [Test]
        public void RaiseCustomEvent_NameofEventName_WithThreeSubscribers_ReturnsCorrectInvocationsCount()
        {
            // Arrange
            var stub = new EventClassStub();
            stub.Custom3Event += OnCustom3EventHandler;
            stub.Custom3Event += OnCustom3EventHandler;
            stub.Custom3Event += OnCustom3EventHandler;

            var p = new object[] { this, new EventArgs() };

            // Act
            var actualCount = CoClassEventReflector.RaiseCustomEvent(stub, typeof(EventClassStub), nameof(EventClassStub.Custom3Event), ref p);

            // Assert
            Assert.AreEqual(3, actualCount);
        }

        [Test]
        public void RaiseCustomEvent_NonExistingEventName_ThrowsArgumentException()
        {
            // Arrange
            var stub = new EventClassStub();
            var p = new object[] { };

            // Act & Assert
            var actualException = Assert.Throws<ArgumentOutOfRangeException>(
                () => CoClassEventReflector.RaiseCustomEvent(stub, typeof(EventClassStub), "NonExistingEventName", ref p));

            Assert.AreEqual("eventName", actualException.ParamName);
            Assert.AreEqual("NonExistingEventName", actualException.ActualValue);
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
