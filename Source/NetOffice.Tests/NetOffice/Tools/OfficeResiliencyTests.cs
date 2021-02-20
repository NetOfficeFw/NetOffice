using System;
using System.IO;
using NUnit.Framework;
using NetOffice.Tools;

namespace NetOffice.Tests.Tools
{
    [TestFixture]
    public class OfficeResiliencyTests
    {
        private static readonly string DataFolderPath = @"NetOffice\Tools\data";

        [Test]
        public void Parse_NetFrameworkAddin_ReturnsDisabledItem()
        {
            // Arrange
            var filepath = Path.Combine(TestContext.CurrentContext.TestDirectory, DataFolderPath, "NetFrameworkAddinResiliency.bin");
            var data = File.ReadAllBytes(filepath);

            // Act
            var actualResult = OfficeResiliency.Parse(data);

            // Assert
            Assert.IsNotNull(actualResult);

            Assert.AreEqual(@"d:\dev\netofficefw\resiliencyaddincrash\bin\debug\resiliencyaddincrash.dll", actualResult.Module);
            Assert.AreEqual("resiliencyaddincrash.resiliencyaddincrashconnect", actualResult.FriendlyName);
        }

        [Test]
        public void Parse_NativeAddin_ReturnsDisabledItem()
        {
            // Arrange
            var filepath = Path.Combine(TestContext.CurrentContext.TestDirectory, DataFolderPath, "NativeAddinResiliency.bin");
            var data = File.ReadAllBytes(filepath);

            // Act
            var actualResult = OfficeResiliency.Parse(data);

            // Assert
            Assert.IsNotNull(actualResult);

            Assert.AreEqual(@"d:\dev\netofficefw\sharedaddin.dll", actualResult.Module);
            Assert.AreEqual("sharedaddin.connect.1", actualResult.FriendlyName);
        }
    }
}
