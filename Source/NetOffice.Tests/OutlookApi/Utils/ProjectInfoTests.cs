using System;
using NUnit.Framework;
using OutlookApi.Utils;

namespace NetOffice.Tests.OutlookApi.Utils
{
    [TestFixture]
    public class ProjectInfoTests
    {
        [Test]
        public void AssemblyName_HasCorrectValue()
        {
            // Arrange
            var projectInfo = new ProjectInfo();

            // Act
            var actualAssemblyName = projectInfo.AssemblyName;

            // Assert
            Assert.AreEqual("OutlookApi", actualAssemblyName);
        }
    }
}
