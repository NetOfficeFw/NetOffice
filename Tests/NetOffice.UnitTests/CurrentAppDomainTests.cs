using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Reflection;
using System.Text;
using System.Threading.Tasks;
using NUnit.Framework;

namespace NetOffice
{
    [TestFixture]
    public class CurrentAppDomainTests
    {
        [Test]
        public void LoadFrom_ValidAssemblyFullFilePath_LoadsAssembly()
        {
            // Arrange
            var fullPath = Path.Combine(TestContext.CurrentContext.TestDirectory, "NetOffice.dll");
            
            var core = new Core();
            var appDomain = core.CurrentAppDomain;

            // Act
            var assembly = appDomain.LoadFrom(fullPath);
            
            // Assert
            Assert.IsNotNull(assembly);
            Assert.AreEqual("NetOffice", assembly.GetName().Name);
        }
    }
}
