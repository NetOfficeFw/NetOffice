using System;
using NetOffice.Diagnostics;
using NUnit.Framework;

namespace NetOffice.Tests.NetOffice
{
    [TestFixture]
    public class SelfDiagnosticsTests
    {
        [Test]
        public void AssemblyTitle_AssemblyWithNoAttribute_ReturnTitleBasedOnFilename()
        {
            // Arrange
            var addin = new NoAssemblyTitleAddin();
            var diag = new SelfDiagnostics(addin);

            // Act
            var title = diag.AssemblyTitle;

            // Assert
            Assert.AreEqual("NoAssemblyTitle", title);
        }
    }
}
