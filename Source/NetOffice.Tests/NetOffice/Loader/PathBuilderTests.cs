using System;
using System.Collections.Generic;
using System.IO;
using NetOffice.Loader;
using NUnit.Framework;

namespace NetOffice.Tests.NetOffice.Loader
{
    [TestFixture]
    public class PathBuilderTests
    {
        [Test]
        [TestCaseSource(nameof(NetOfficeAssemblyNameTestCase))]
        public void BuildLocalPathFromAssemblyFileName_NetOfficeAssemblyName_ResolvesPathToAssemblyFile(string assemblyName)
        {
            // Arrange
            var factory = new Core();

            // Act
            var path = PathBuilder.BuildLocalPathFromAssemblyFileName(factory, assemblyName);

            // Assert
            StringAssert.EndsWith(assemblyName, path);
            Assert.IsTrue(Path.IsPathRooted(path));
        }

        public static IEnumerable<string> NetOfficeAssemblyNameTestCase
        {
            get
            {
                return Core.Default.CoreDomain.AssemblyNames;
            }
        }
    }
}