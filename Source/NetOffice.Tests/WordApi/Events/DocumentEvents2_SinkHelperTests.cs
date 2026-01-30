// Copyright 2024 Cisco Systems, Inc.
// Licensed under MIT-style license (see LICENSE.txt file).

using System;
using NetOffice.WordApi;
using NetOffice.WordApi.Events;
using NetOffice.Tests.Helpers;
using NUnit.Framework;

namespace NetOffice.Tests.WordApi.Events
{
    [TestFixture]
    public class DocumentEvents2_SinkHelperTests
    {
        /// <summary>
        /// Regression test for #453 - Ensures ContentControlBeforeContentUpdate event uses correct event name
        /// in both Validate() and EventBinding.RaiseCustomEvent() calls
        /// </summary>
        [Test]
        public void ContentControlBeforeContentUpdate_EventRaised_CallsHandlerWithCorrectEventName()
        {
            // Arrange
            var eventBinder = new TestableComObjectStub();
            eventBinder.AddEventRecipient(nameof(DocumentEvents2_SinkHelper.ContentControlBeforeContentUpdate));
            var connectionPoint = new ConnectionPointStub();

            var events = new DocumentEvents2_SinkHelper(eventBinder, connectionPoint);
            var parameter1 = new FakeComObject();
            var content = new object();

            // Act
            events.ContentControlBeforeContentUpdate(parameter1, ref content);
            var actualParametersPassToEvent = eventBinder.LastRaisedEventParameters;

            // Assert
            Assert.AreEqual("ContentControlBeforeContentUpdate", eventBinder.LastRaisedEventName,
                "EventBinding.RaiseCustomEvent must be called with 'ContentControlBeforeContentUpdate' event name");

            CollectionAssert.IsNotEmpty(actualParametersPassToEvent);
            var actualParameter1 = actualParametersPassToEvent[0];
            Assert.IsInstanceOf<ContentControl>(actualParameter1,
                "Event ContentControlBeforeContentUpdate parameter must be of type ContentControl.");
        }
    }
}
