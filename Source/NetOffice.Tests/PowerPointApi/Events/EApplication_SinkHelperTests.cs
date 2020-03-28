using System;
using System.Collections;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using NetOffice.PowerPointApi;
using NetOffice.PowerPointApi.Events;
using NetOffice.Tests.Helpers;
using NUnit.Framework;

namespace NetOffice.Tests.PowerPointApi.Events
{
    [TestFixture]
    public class EApplication_SinkHelperTests
    {
        /// <summary>
        /// Regression test for #262 (ActivePowerPointApp.SlideShowBeginEvent doesn't work in versions 1.7.4.x)
        /// </summary>
        [Test]
        public void SlideShowBegin_EventRaised_CallsHandlerWithCorrectObject()
        {
            // Arrange
            var eventBinder = new TestableComObjectStub();
            var connectionPoint = new ConnectionPointStub();

            var events = new EApplication_SinkHelper(eventBinder, connectionPoint);
            var parameter1 = new FakeComObject();

            // Act
            events.SlideShowBegin(parameter1);
            var actualParametersPassToEvent = eventBinder.LastRaisedEventParameters;

            // Assert
            Assert.AreEqual("SlideShowBegin", eventBinder.LastRaisedEventName);

            CollectionAssert.IsNotEmpty(actualParametersPassToEvent);
            var actualParameter1 = actualParametersPassToEvent[0];
            Assert.IsInstanceOf<SlideShowWindow>(actualParameter1, "Event SlideShowBegin parameter must be of type SlideShowWindow.");
        }

        /// <summary>
        /// Regression test for #262 (ActivePowerPointApp.SlideShowBeginEvent doesn't work in versions 1.7.4.x)
        /// </summary>
        [Test]
        [TestCaseSource(nameof(PowerPointSlideShowEventsTestData))]
        public void PowerPointApplication_SlideShowEventIsRaised_CallsHandlerWithCorrectObject(string expectedEventName, Type expectedParameterType, Action<EApplication_SinkHelper> action)
        {
            // Arrange
            var eventBinder = new TestableComObjectStub();
            var connectionPoint = new ConnectionPointStub();

            var events = new EApplication_SinkHelper(eventBinder, connectionPoint);

            // Act
            action(events);
            var actualParametersPassToEvent = eventBinder.LastRaisedEventParameters;

            // Assert
            Assert.AreEqual(expectedEventName, eventBinder.LastRaisedEventName);

            CollectionAssert.IsNotEmpty(actualParametersPassToEvent);
            var actualParameter1 = actualParametersPassToEvent[0];
            Assert.IsInstanceOf(expectedParameterType, actualParameter1, $"Event '{expectedEventName}' parameter must be of type '{expectedParameterType.Name}'");
        }

        public static IEnumerable PowerPointSlideShowEventsTestData()
        {
            yield return new TestCaseData("SlideShowBegin", typeof(SlideShowWindow), new Action<EApplication_SinkHelper>((events) => events.SlideShowBegin(new FakeComObject())));
            yield return new TestCaseData("SlideShowNextBuild", typeof(SlideShowWindow), new Action<EApplication_SinkHelper>((events) => events.SlideShowNextBuild(new FakeComObject())));
            yield return new TestCaseData("SlideShowNextSlide", typeof(SlideShowWindow), new Action<EApplication_SinkHelper>((events) => events.SlideShowNextSlide(new FakeComObject())));
        }
    }
}
