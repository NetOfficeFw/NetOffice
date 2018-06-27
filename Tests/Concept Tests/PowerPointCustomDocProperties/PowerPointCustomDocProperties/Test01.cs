using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using NetOffice;
using Office = NetOffice.OfficeApi;
using NetOffice.OfficeApi.Enums;
using PowerPoint = NetOffice.PowerPointApi;
using NetOffice.PowerPointApi.Enums;
using NetOffice.Extensions.Invoker;
using NetOffice.Exceptions;

namespace PowerPointCustomDocProperties
{
    public class Test01
    {
        public void Proceed(int testNumber)
        {
            using (var application = COMObject.Create<PowerPoint.Application>(COMObjectCreateOptions.CreateNewCore))
            {
                application.Settings.EnableAutomaticQuit = true;
                application.Settings.ExceptionMessageBehavior = ExceptionMessageHandling.DiagnosticsAndInnerMessage;
                application.Settings.ExceptionDiagnosticsMessage = "Failed to proceed {CallInstance}={CallType}=>{Name}{ParenthesizedArgs}.";

                application.Console.OnException += delegate(DebugConsole sender, Exception error)
                {
                    NetOfficeCOMException expo = error as NetOfficeCOMException;
                    if(null != expo)
                        Console.WriteLine("Test{0}, Console_OnException: {1} {2}", testNumber, expo.Message, expo.ApplicationVersion);
                };

                application.DisplayAlerts = PpAlertLevel.ppAlertsNone;
                application.Visible = MsoTriState.msoCTrue;

                var presentation = application.Presentations.Add();
                var properties = presentation.Property<Office.DocumentProperties>("CustomDocumentProperties");

                properties.Add(
                    "CustomProperty1",
                    false,
                    MsoDocProperties.msoPropertyTypeNumber,
                    new Random().Next(1, 999)
                    );

                application.Quit();
            }
        }
    }
}
