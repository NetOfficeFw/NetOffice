using System;
using System.Collections.Generic;
using System.Linq;
using System.Runtime.InteropServices;
using NetOffice;
using NetOffice.Tools;
using Vbe = NetOffice.VBIDEApi;
using NetOffice.VBIDEApi.Tools;
using Office = NetOffice.OfficeApi;
using NetOffice.OfficeApi.Enums;

namespace VbeAddin
{
    /// <summary>
    /// Import/Export standard code modules from/to local file system.
    /// </summary>
    [COMAddin("Vbe Sample Addin CS4", "Simple Vbe Addin Showstopper", LoadBehavior.LoadAtStartup), Codebase]
    [ProgId("VbeAddinCS4.Connect"), Guid("C4586D47-5BCF-4800-94AA-92DFF99D3805")]
    public class Connect : COMAddin
    {
        public Connect()
        {
            OnStartupComplete += Connect_OnStartupComplete;
        }

        public CommandBarButtons Buttons { get; private set; }

        private void Connect_OnStartupComplete(ref Array custom)
        {
            Buttons = new CommandBarButtons(Application);

            Buttons.ExportRequested += delegate
            {
                var project = SelectProjectDialog.SelectProject(new ProjectCollector(Application));
                if (null != project)
                {
                    using (var repository = new VbaProjectRepository(Application))
                    {
                        repository.Export(project);
                    }
                }
            };

            Buttons.ImportRequested += delegate
            {
                using (var repository = new VbaProjectRepository(Application))
                {
                    var project = SelectProjectDialog.SelectProject(new ProjectCollector(repository));
                    if (null != project)
                        repository.Import(project);
                }              
            };
        }
    }
}