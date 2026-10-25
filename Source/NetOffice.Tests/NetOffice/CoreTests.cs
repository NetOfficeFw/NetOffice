using NetOffice.Tests.Helpers;
using NUnit.Framework;

namespace NetOffice.Tests.NetOffice
{
    [TestFixture]
    public class CoreTests
    {
        [Test]
        [TestCase(true)]
        [TestCase(false)]
        public void CreateNewProxyShare_IsEnumeratorArgument_PassesValueToCreateProxyShareEvent(bool isEnumerator)
        {
            // Arrange
            var factory = new Core();
            bool? actualIsEnumerator = null;
            factory.CreateProxyShare += (sender, args) => actualIsEnumerator = args.IsEnumerator;

            // Act
            var proxyShare = factory.CreateNewProxyShare(null, new FakeComObject(), isEnumerator);

            // Assert
            Assert.AreEqual(isEnumerator, actualIsEnumerator);
            Assert.AreEqual(isEnumerator, proxyShare.IsEnumerator);
        }
    }
}
