using System;
using System.Diagnostics.CodeAnalysis;
using NUnit.Framework;

namespace NetOffice.Tests.NetOffice
{
    [TestFixture]
    public class StringExTests
    {
        [Test]
        [TestCase("AssemblyName.resources")]
        [TestCase("AssemblyName.RESOURCES")]
        [TestCase("AssemblyName.resources, Version=1.0.0.0, Culture=en-US")]
        [SuppressMessage("ReSharper", "InvokeAsExtensionMethod")]
        public void ContainsIgnoreCase_MatchingString_ReturnsTrue(string text)
        {
            // Arrange

            // Act
            bool actualResult = StringEx.ContainsIgnoreCase(text, ".resources");

            // Assert
            Assert.IsTrue(actualResult);
        }

        [Test]
        public void ContainsIgnoreCase_ExtensionMethod_CanBeCalled()
        {
            // Arrange
            string text = "Sample Text";

            // Act
            bool actualResult = text.ContainsIgnoreCase("sample");

            // Assert
            Assert.IsTrue(actualResult);
        }

        [Test]
        [SuppressMessage("ReSharper", "InvokeAsExtensionMethod")]
        public void ContainsIgnoreCase_NullParameter_ReturnsFalse()
        {
            // Arrange

            // Act
            bool actualResult = StringEx.ContainsIgnoreCase(null, ".resources");

            // Assert
            Assert.IsFalse(actualResult);
        }

        [Test]
        [SuppressMessage("ReSharper", "InvokeAsExtensionMethod")]
        public void ContainsIgnoreCase_EmptyStringParameter_ReturnsFalse()
        {
            // Arrange

            // Act
            bool actualResult = StringEx.ContainsIgnoreCase(String.Empty, ".resources");

            // Assert
            Assert.IsFalse(actualResult);
        }
    }
}
