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
    public class CommandBarButtons
    {
        public CommandBarButtons(Vbe.VBE application)
        {
            Create(application);
        }

        public event Action ExportRequested;

        public event Action ImportRequested;

        private void Create(Vbe.VBE application)
        {
            var commandBars = application.CommandBars;

            Office.CommandBar commandBar = commandBars.Add("Share code from file system", MsoBarPosition.msoBarTop, null, true);
            commandBar.Visible = true;

            Office.CommandBarButton exportButton = (Office.CommandBarButton)commandBar.Controls.Add(MsoControlType.msoControlButton, null, null, null, true);
            exportButton.Style = MsoButtonStyle.msoButtonIconAndCaption;
            exportButton.FaceId = 9;
            exportButton.Caption = "Export Vba Code Modules";
            exportButton.Visible = true;
            exportButton.ClickEvent += ExportButton_ClickEvent;

            Office.CommandBarButton importButton = (Office.CommandBarButton)commandBar.Controls.Add(MsoControlType.msoControlButton, null, null, null, true);
            importButton.Style = MsoButtonStyle.msoButtonIconAndCaption;
            importButton.FaceId = 9;
            importButton.Caption = "Import Vba Code Modules";
            importButton.Visible = true;
            importButton.ClickEvent += ImportButton_ClickEvent;

            commandBars.Dispose(false);
        }

        private void ExportButton_ClickEvent(NetOffice.OfficeApi.CommandBarButton ctrl, ref bool cancelDefault)
        {
            ExportRequested?.Invoke();
        }

        private void ImportButton_ClickEvent(NetOffice.OfficeApi.CommandBarButton Ctrl, ref bool cancelDefault)
        {
            ImportRequested?.Invoke();
        }
    }
}