using System;
using NetOffice.Diagnostics;
using NetOffice.OfficeApi.Tools.Informations;
using NUnit.Framework;

namespace NetOffice.Tests.NetOffice
{
    [TestFixture]
    public class AssemblyInfoTests
    {
        [Test]
        public void AssemblyTitle_AssemblyWithNoAttribute_ReturnTitleBasedOnFilename()
        {
            // Arrange
            var addin = new NoAssemblyTitleAddin();
            var ownerAssembly = addin.GetType().Assembly;
            var diag = new AssemblyInfo(ownerAssembly);

            // Act
            var title = diag.AssemblyTitle;

            // Assert
            Assert.AreEqual("NoAssemblyTitle", title);
        }
    }
}
