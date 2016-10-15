using System;
using System.Reflection;
using NUnit.Framework;

namespace NetOffice
{
    [TestFixture]
    public class CoreTests
    {
        private byte[] NetOfficePublicKey { get; set; }

        [OneTimeSetUp]
        public void OneTimeSetUp()
        {
            var name = this.GetType().Assembly.GetName();
            this.NetOfficePublicKey = name.GetPublicKeyToken();
        }

        [Test]
        [TestCase("WordApi, Version=1.7.3.0, Culture=neutral, PublicKeyToken=d0b2dc7c792d5ca6")]
        public void ContainsNetOfficePublicKeyToken_AssemblyNameWithValidNetOfficeToken_ReturnsTrue(string assemblyName)
        {
            // Arrange
            var name = new AssemblyName(assemblyName);
            var core = new Core();

            // Act
            var actualToken = core.ContainsNetOfficePublicKeyToken(name, this.NetOfficePublicKey);

            // Assert
            Assert.True(actualToken);
        }

        [Test]
        [TestCase("AssemblyA, Version=1.0.0.0, Culture=neutral")]
        public void ContainsNetOfficePublicKeyToken_AssemblyNameWithoutPublicKeyToken_ReturnsFalse(string assemblyName)
        {
            // Arrange
            var name = new AssemblyName(assemblyName);
            var core = new Core();

            // Act
            var actualContainsToken = core.ContainsNetOfficePublicKeyToken(name, this.NetOfficePublicKey);

            // Assert
            Assert.False(actualContainsToken);
        }

        [Test]
        [TestCase("AssemblyB, Version=1.0.0.0, Culture=neutral, PublicKeyToken=9589fa1be527eb6c")]
        public void ContainsNetOfficePublicKeyToken_AssemblyNameWithInvalidPublicKeyToken_ReturnsFalse(string assemblyName)
        {
            // Arrange
            var name = new AssemblyName(assemblyName);
            var core = new Core();

            // Act
            var actualContainsToken = core.ContainsNetOfficePublicKeyToken(name, this.NetOfficePublicKey);

            // Assert
            Assert.False(actualContainsToken);
        }
    }
}
