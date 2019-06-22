using System;
using NUnit.Framework;
using AccessApi.Utils;

namespace NetOffice.Tests.AccessApi.Utils
{
    [TestFixture]
    public class ProjectInfoTests
    {
        /// <summary>
        /// Regression test for #231 (Access library ProjectInfo returns incorrect AssemblyName value)
        /// </summary>
        [Test]
        public void AssemblyName_HasCorrectValue()
        {
            // Arrange
            var projectInfo = new ProjectInfo();

            // Act
            var actualAssemblyName = projectInfo.AssemblyName;

            // Assert
            Assert.AreEqual("AccessApi", actualAssemblyName);
        }
    }
}
