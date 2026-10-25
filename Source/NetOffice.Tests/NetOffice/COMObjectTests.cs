using NetOffice.Exceptions;
using NetOffice.Tests.Helpers;
using NUnit.Framework;

namespace NetOffice.Tests.NetOffice
{
    [TestFixture]
    public class COMObjectTests
    {
        [Test]
        public void Dispose_CalledTwiceWhileProxyShareIsShared_ReleasesProxyShareOnce()
        {
            // Arrange
            var factory = new Core();
            var first = new COMObject(factory, null, new FakeComObject());
            var proxyShare = ((ICOMProxyShareProvider)first).GetProxyShare();

            // second instance keeps the shared proxy acquired
            var second = new COMObject(factory, null, proxyShare);

            // Act
            first.Dispose();
            first.Dispose();

            // Assert
            Assert.AreEqual(1, proxyShare.Count, "Shared proxy must stay acquired by the second instance.");
            Assert.IsFalse(proxyShare.Released);
            Assert.IsTrue(first.IsDisposed);
            Assert.IsFalse(second.IsDisposed);
        }

        [Test]
        public void Dispose_InstanceIsAlreadyDisposed_DoesNotRaiseOnDisposeEvent()
        {
            // Arrange
            var factory = new Core();
            var first = new COMObject(factory, null, new FakeComObject());
            var proxyShare = ((ICOMProxyShareProvider)first).GetProxyShare();

            // second instance keeps the shared proxy acquired
            var second = new COMObject(factory, null, proxyShare);

            var onDisposeCount = 0;
            first.OnDispose += (eventArgs) => onDisposeCount++;

            // Act
            first.Dispose();
            first.Dispose();

            // Assert
            Assert.AreEqual(1, onDisposeCount);
        }

        [Test]
        public void Dispose_PreviousDisposeFailed_DisposesInstanceOnRetry()
        {
            // Arrange
            var factory = new Core();
            var instance = new COMObject(factory, null, new FakeComObject());

            // releasing the fake proxy fails because it is not a real COM object
            Assert.Throws<COMDisposeException>(() => instance.Dispose());

            // Act
            instance.Dispose();

            // Assert
            Assert.IsTrue(instance.IsDisposed);
            Assert.AreEqual(0, factory.ProxyCount);
        }
    }
}
