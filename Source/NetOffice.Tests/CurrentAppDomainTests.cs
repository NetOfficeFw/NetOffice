using System;
using NUnit.Framework;

namespace NetOffice.Tests
{
    public class CurrentAppDomainTests
    {
        [Test]
        public void CurrentDomain_AssemblyResolve_TypeNameWithoutComma_DoesNotThrowException()
        {
            // Arange
            var core = new Core();
            core.Console.Mode = DebugConsoleMode.MemoryList;

            // Act - will trigger CurrentDomain_AssemblyResolve handler
            var unityApplicationClass = Type.GetType("UnityEngine.Application, UnityEngine");

            // Assert
            Assert.AreEqual(0, core.Console.Messages.Length);
        }
    }
}
