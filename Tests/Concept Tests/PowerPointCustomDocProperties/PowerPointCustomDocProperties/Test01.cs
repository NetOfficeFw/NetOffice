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

namespace PowerPointCustomDocProperties
{
    public class Test01
    {
        public NetOffice.Core Proceed(int testNumber)
        {
            using (var application = COMObject.Create<PowerPoint.Application>(COMObjectCreateOptions.CreateNewCore))
            {
                application.Console.OnException += delegate (DebugConsole arg1, Exception arg2)
                {
                    Console.WriteLine("Test{0}, Console_OnException: {1}", testNumber, arg2.Message);
                };

                application.DisplayAlerts = PpAlertLevel.ppAlertsNone;
                application.Visible = MsoTriState.msoCTrue;
                application.Settings.EnableAutomaticQuit = true;
                application.Settings.ExceptionMessageBehavior = ExceptionMessageHandling.DiagnosticsAndInnerMessage;

                var presentation = application.Presentations.Add();
                var properties = presentation.Property<Office.DocumentProperties>("CustomDocumentProperties");

                var randomValue = new Random().Next(1, 999);

                properties.Add("CustomProperty1", false, MsoDocProperties.msoPropertyTypeNumber, randomValue);

                return application.Factory;
            }
        }
    }
}
