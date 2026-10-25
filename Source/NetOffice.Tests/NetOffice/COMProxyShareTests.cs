using NetOffice.Tests.Helpers;
using NUnit.Framework;

namespace NetOffice.Tests.NetOffice
{
    [TestFixture]
    public class COMProxyShareTests
    {
        [Test]
        public void Release_CalledMoreTimesThanAcquired_DoesNotDecrementCountBelowZero()
        {
            // Arrange
            var proxyShare = new COMProxyShare(new Core(), new FakeComObject());
            proxyShare.SuppressReleaseExceptions = true;

            // releasing the fake proxy fails because it is not a real COM object
            proxyShare.Release();

            var countChangedCount = 0;
            proxyShare.CountChanged += (sender) => countChangedCount++;

            // Act
            var released = proxyShare.Release();

            // Assert
            Assert.IsFalse(released);
            Assert.AreEqual(0, proxyShare.Count);
            Assert.AreEqual(0, countChangedCount, "CountChanged event must not be raised when the count does not change.");
        }
    }
}
