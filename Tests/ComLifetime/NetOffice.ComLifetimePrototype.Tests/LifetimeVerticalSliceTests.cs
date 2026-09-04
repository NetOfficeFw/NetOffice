using System;
using System.ComponentModel;
using System.IO;
using System.Runtime.InteropServices;
using NetOffice;
using NUnit.Framework;

namespace NetOffice.ComLifetimePrototype.Tests
{
    [TestFixture]
    [NUnit.Framework.Category("IntegrationTests")]
    public sealed class LifetimeVerticalSliceTests
    {
        private static readonly Guid ProbeClassId = new Guid("32D43B6B-6198-4E98-9386-49D6B576580A");

        [Test]
        [Apartment(System.Threading.ApartmentState.STA)]
        public void HarnessDetectsCurrentSharedRcwDisconnection()
        {
            NativeTelemetry.Reset();

            string manifest = Path.Combine(TestContext.CurrentContext.TestDirectory, "NativeLifetimeFixture.manifest");
            using (new ActivationContextScope(manifest))
            {
                Type probeType = Type.GetTypeFromCLSID(ProbeClassId, true);
                object rcw = Activator.CreateInstance(probeType);

                Assert.That(Marshal.IsComObject(rcw), Is.True, "Registration-free activation must produce an RCW.");
                Assert.That(NativeTelemetry.ConstructCount, Is.EqualTo(1));
                Assert.That(NativeTelemetry.LiveObjectCount, Is.EqualTo(1));

                var first = new ProbeWrapper(rcw);
                var second = new ProbeWrapper(rcw);

                Assert.That(first.Identity, Is.EqualTo(4242));
                Assert.That(first.Ping(), Is.EqualTo("pong"));
                Assert.That(second.Ping(), Is.EqualTo("pong"));

                int invokesBeforeDispose = NativeTelemetry.InvokeCount;
                first.Dispose();

                Assert.That(first.IsDisposed, Is.True);
                Assert.That(second.IsDisposed, Is.False, "The second NetOffice wrapper is independently live.");

                Exception danglingFailure = null;
                try
                {
                    second.Ping();
                }
                catch (Exception exception)
                {
                    danglingFailure = exception;
                }

                Assert.Multiple(() =>
                {
                    Assert.That(danglingFailure, Is.Not.Null,
                        "The current wrapper-local release should be detected by this negative-control slice.");
                    Assert.That(NativeTelemetry.InvokeCount, Is.EqualTo(invokesBeforeDispose),
                        "The dangling wrapper must fail before entering the destroyed native object.");
                    Assert.That(NativeTelemetry.DestroyCount, Is.EqualTo(1),
                        "Native destruction is the important lifetime event observed by the fixture.");
                    Assert.That(NativeTelemetry.LiveObjectCount, Is.EqualTo(0));
                    Assert.That(NativeTelemetry.ReleaseCount, Is.GreaterThan(0));
                });
            }
        }

        private sealed class ProbeWrapper : COMObject
        {
            public ProbeWrapper(object comProxy) : base(comProxy) { }

            public int Identity
            {
                get { return Factory.ExecuteInt32PropertyGet(this, "Identity"); }
            }

            public string Ping()
            {
                return Factory.ExecuteStringMethodGet(this, "Ping");
            }
        }

        private sealed class ActivationContextScope : IDisposable
        {
            private static readonly IntPtr InvalidHandleValue = new IntPtr(-1);
            private readonly IntPtr _handle;
            private readonly IntPtr _cookie;

            public ActivationContextScope(string manifestPath)
            {
                var context = new ActContext
                {
                    Size = Marshal.SizeOf(typeof(ActContext)),
                    Source = manifestPath
                };

                _handle = CreateActCtx(ref context);
                if (_handle == InvalidHandleValue)
                    throw new Win32Exception(Marshal.GetLastWin32Error(), "CreateActCtx failed for " + manifestPath);

                if (!ActivateActCtx(_handle, out _cookie))
                {
                    int error = Marshal.GetLastWin32Error();
                    ReleaseActCtx(_handle);
                    throw new Win32Exception(error, "ActivateActCtx failed for " + manifestPath);
                }
            }

            public void Dispose()
            {
                DeactivateActCtx(0, _cookie);
                ReleaseActCtx(_handle);
            }

            [StructLayout(LayoutKind.Sequential, CharSet = CharSet.Unicode)]
            private struct ActContext
            {
                public int Size;
                public uint Flags;
                [MarshalAs(UnmanagedType.LPWStr)] public string Source;
                public ushort ProcessorArchitecture;
                public ushort LanguageId;
                [MarshalAs(UnmanagedType.LPWStr)] public string AssemblyDirectory;
                public IntPtr ResourceName;
                [MarshalAs(UnmanagedType.LPWStr)] public string ApplicationName;
                public IntPtr Module;
            }

            [DllImport("kernel32.dll", CharSet = CharSet.Unicode, SetLastError = true)]
            private static extern IntPtr CreateActCtx(ref ActContext context);

            [DllImport("kernel32.dll", SetLastError = true)]
            private static extern bool ActivateActCtx(IntPtr handle, out IntPtr cookie);

            [DllImport("kernel32.dll", SetLastError = true)]
            private static extern bool DeactivateActCtx(uint flags, IntPtr cookie);

            [DllImport("kernel32.dll")]
            private static extern void ReleaseActCtx(IntPtr handle);
        }

        private static class NativeTelemetry
        {
            public static int LiveObjectCount { get { return LifetimeFixture_GetLiveObjectCount(); } }
            public static int ConstructCount { get { return LifetimeFixture_GetConstructCount(); } }
            public static int DestroyCount { get { return LifetimeFixture_GetDestroyCount(); } }
            public static int ReleaseCount { get { return LifetimeFixture_GetReleaseCount(); } }
            public static int InvokeCount { get { return LifetimeFixture_GetInvokeCount(); } }

            public static void Reset()
            {
                LifetimeFixture_ResetTelemetry();
            }

            [DllImport("NativeLifetimeFixture.dll", CallingConvention = CallingConvention.StdCall)]
            private static extern void LifetimeFixture_ResetTelemetry();

            [DllImport("NativeLifetimeFixture.dll", CallingConvention = CallingConvention.StdCall)]
            private static extern int LifetimeFixture_GetLiveObjectCount();

            [DllImport("NativeLifetimeFixture.dll", CallingConvention = CallingConvention.StdCall)]
            private static extern int LifetimeFixture_GetConstructCount();

            [DllImport("NativeLifetimeFixture.dll", CallingConvention = CallingConvention.StdCall)]
            private static extern int LifetimeFixture_GetDestroyCount();

            [DllImport("NativeLifetimeFixture.dll", CallingConvention = CallingConvention.StdCall)]
            private static extern int LifetimeFixture_GetReleaseCount();

            [DllImport("NativeLifetimeFixture.dll", CallingConvention = CallingConvention.StdCall)]
            private static extern int LifetimeFixture_GetInvokeCount();
        }
    }
}
