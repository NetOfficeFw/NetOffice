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
    }
}
