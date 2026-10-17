using System;
using System.Reflection;
using NetOffice.ExcelApi.GlobalHelperModules;
using NUnit.Framework;

namespace NetOffice.Tests.ExcelApi
{
    [TestFixture]
    [NonParallelizable]
    public class GlobalModuleTests
    {
        [Test]
        public void DDERequest_UsesDdeMemberAndPreservesArguments()
        {
            var proxy = new DispatchProbe();

            object result = Invoke(proxy, () => GlobalModule.DDERequest(17, "prices"));

            Assert.Multiple(() =>
            {
                Assert.AreEqual("dde-result", result);
                Assert.AreEqual("DDERequest", proxy.Member);
                Assert.AreEqual(1, proxy.InvocationCount);
                CollectionAssert.AreEqual(new object[] { 17, "prices" }, proxy.Arguments);
            });
        }

        [TestCase(29)]
        [TestCase(30)]
        public void Run2_MaximumAndNeighboringOverloads_PreserveTargetAndArgumentOrder(int argumentCount)
        {
            var proxy = new DispatchProbe();
            var arguments = new object[argumentCount + 1];
            arguments[0] = "MacroUnderTest";
            for (int index = 1; index < arguments.Length; index++)
                arguments[index] = index;

            object result = Invoke(proxy, () => argumentCount == 29
                ? GlobalModule._Run2(arguments[0], arguments[1], arguments[2], arguments[3], arguments[4],
                    arguments[5], arguments[6], arguments[7], arguments[8], arguments[9], arguments[10],
                    arguments[11], arguments[12], arguments[13], arguments[14], arguments[15], arguments[16],
                    arguments[17], arguments[18], arguments[19], arguments[20], arguments[21], arguments[22],
                    arguments[23], arguments[24], arguments[25], arguments[26], arguments[27], arguments[28], arguments[29])
                : GlobalModule._Run2(arguments[0], arguments[1], arguments[2], arguments[3], arguments[4],
                    arguments[5], arguments[6], arguments[7], arguments[8], arguments[9], arguments[10],
                    arguments[11], arguments[12], arguments[13], arguments[14], arguments[15], arguments[16],
                    arguments[17], arguments[18], arguments[19], arguments[20], arguments[21], arguments[22],
                    arguments[23], arguments[24], arguments[25], arguments[26], arguments[27], arguments[28], arguments[29], arguments[30]));

            Assert.Multiple(() =>
            {
                Assert.AreEqual("run2-result", result);
                Assert.AreEqual("_Run2", proxy.Member);
                Assert.AreEqual(1, proxy.InvocationCount);
                CollectionAssert.AreEqual(arguments, proxy.Arguments);
            });
        }

        private static object Invoke(DispatchProbe proxy, Func<object> action)
        {
            var core = new Core();
            bool previousSafeMode = core.Settings.EnableSafeMode;
            core.Settings.EnableSafeMode = false;
            var owner = new COMObject(core, null, proxy);
            FieldInfo instance = typeof(GlobalModule).GetField("_instance", BindingFlags.Static | BindingFlags.NonPublic);
            object previousInstance = instance.GetValue(null);
            try
            {
                instance.SetValue(null, owner);
                return action();
            }
            finally
            {
                instance.SetValue(null, previousInstance);
                core.Settings.EnableSafeMode = previousSafeMode;
                // No RCW exists for this managed proxy; COM release is not applicable.
                GC.KeepAlive(owner);
            }
        }

        public sealed class DispatchProbe : MarshalByRefObject
        {
            public string Member { get; private set; }
            public object[] Arguments { get; private set; }
            public int InvocationCount { get; private set; }

            public object DDERequest(int channel, string item)
            {
                Record("DDERequest", new object[] { channel, item });
                return "dde-result";
            }

            public object Selection(int channel, string item)
            {
                Record("Selection", new object[] { channel, item });
                return "selection-result";
            }

            public object Run(params object[] arguments)
            {
                Record("Run", arguments);
                return "run-result";
            }

            public object _Run2(params object[] arguments)
            {
                Record("_Run2", arguments);
                return "run2-result";
            }

            private void Record(string member, object[] arguments)
            {
                Member = member;
                Arguments = arguments;
                InvocationCount++;
            }
        }
    }
}
