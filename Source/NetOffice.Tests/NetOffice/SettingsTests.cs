using System;
using NUnit.Framework;

namespace NetOffice.Tests.NetOffice
{
    [TestFixture]
    public class SettingsTests
    {
        [Test]
        public void IsEqualTo_DefaultContructor_ReturnsTrue()
        {
            // Arrange
            var settings1 = new Settings();
            var settings2 = new Settings();

            // Act
            var actualResult = settings1.IsEqualTo(settings2);

            // Assert
            Assert.IsTrue(actualResult);
        }

        [Test]
        public void IsEqualTo_NullObject_ReturnsFalse()
        {
            // Arrange
            var settings1 = new Settings();

            // Act
            var actualResult = settings1.IsEqualTo(null);

            // Assert
            Assert.IsFalse(actualResult, "Settings object must not be equal to null object.");
        }

        [Test]
        public void IsEqualTo_EnableSafeModeValueIsDifferent_ReturnsFalse()
        {
            // Arrange
            var settings1 = new Settings();
            var settings2 = new Settings();

            settings1.EnableSafeMode = false;
            settings1.EnableSafeMode = true;

            // Act
            var actualResult = settings1.IsEqualTo(settings2);

            // Assert
            Assert.IsFalse(actualResult, "Settings objects must not be equal when EnableSafeMode value is different.");
        }
    }
}
