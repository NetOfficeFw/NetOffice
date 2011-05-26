using System;
using System.Collections.Generic;
using Microsoft.Win32;
using Extensibility;
using System.Runtime.InteropServices;
using System.Text;
using System.Reflection;

using LateBindingApi.Core;
using Office = NetOffice.OfficeApi;

using Excel = NetOffice.ExcelApi;
using Word = NetOffice.WordApi;
using Outlook = NetOffice.OutlookApi;
using PowerPoint = NetOffice.PowerPointApi;
using Access = NetOffice.AccessApi;
  
namespace SuperAddin.UIMapper
{
    /// <summary>
    /// handles classic ui
    /// </summary>
    public class ClassicUI
    {
        #region Fields
        
        AddinUI _parent;
        
        #endregion

        #region Construction

        internal ClassicUI(AddinUI parent)
        {
            _parent = parent;
        }
        
        #endregion

        #region Methods

        /// <summary>
        /// calls specific create method for office application type
        /// note: all applications has the same code except outlook (active inspector)
        /// </summary>
        public void CreateUI()
        {
            if (_parent.Application.Application is Excel.Application)
            {
                CreateExcelUI();
            }
            else if (_parent.Application.Application is Word.Application)
            {
                CreateWordUI();
            }
            else if (_parent.Application.Application is Outlook.Application)
            {
                CreateOutlookUI();
            }
            else if (_parent.Application.Application is PowerPoint.Application)
            {
                CreatePowerPointUI();
            }
            else if (_parent.Application.Application is Access.Application)
            {
                CreateAccessUI();
            }
        }
        
        #endregion

        #region Event Trigger
        
        public void commandBarBtn_ClickEvent(Office.CommandBarButton Ctrl, ref bool CancelDefault)
        {
            try
            {
                _parent.RaiseButtonClick(new ButtonClickArgs(Ctrl));
            }
            catch (Exception throwedException)
            {
                FormShowError.LogError("An error ocurred while perform commandBarBtn_ClickEvent.", throwedException);
            }
        }

        #endregion

        #region Specific CreateUI Methods

        private void CreateExcelUI()
        {
            Excel.Application excelApp = _parent.Application.Application as Excel.Application;
            Office.CommandBar commandBar = excelApp.CommandBars.Add("SuperAddinCommandbar", Office.Enums.MsoBarPosition.msoBarTop, false, true);
            commandBar.Visible = true;

            // add a button to the toolbar
            Office.CommandBarButton commandBarBtn = (Office.CommandBarButton)commandBar.Controls.Add(Office.Enums.MsoControlType.msoControlButton, Missing.Value, Missing.Value, Missing.Value, true);
            commandBarBtn.Style = Office.Enums.MsoButtonStyle.msoButtonIconAndCaption;
            commandBarBtn.Caption = "ExcelButton1";
            commandBarBtn.FaceId = 2;
            commandBarBtn.ClickEvent += new Office.CommandBarButton_ClickEventHandler(commandBarBtn_ClickEvent);

            excelApp.DisposeChildInstances(false);
        }

        private void CreateOutlookUI()
        {
            Outlook.Application outlookApp = _parent.Application.Application as Outlook.Application;
            Office.CommandBar commandBar = outlookApp.ActiveExplorer().CommandBars.Add("SuperAddinCommandBar", Office.Enums.MsoBarPosition.msoBarTop, false, true);
            commandBar.Visible = true;
            
            // add a button to the toolbar
            Office.CommandBarButton commandBarBtn = (Office.CommandBarButton)commandBar.Controls.Add(Office.Enums.MsoControlType.msoControlButton, Missing.Value, Missing.Value, Missing.Value, true);
            commandBarBtn.Style = Office.Enums.MsoButtonStyle.msoButtonIconAndCaption;
            commandBarBtn.Caption = "OutlookButton1";
            commandBarBtn.FaceId = 2;
            commandBarBtn.ClickEvent += new Office.CommandBarButton_ClickEventHandler(commandBarBtn_ClickEvent);

            outlookApp.DisposeChildInstances(false);
        }

        private void CreateWordUI()
        {
            Word.Application   wordApp = _parent.Application.Application as Word.Application;
            Office.CommandBar commandBar = wordApp.CommandBars.Add("SuperAddinCommandbar", Office.Enums.MsoBarPosition.msoBarTop, false, true);
            commandBar.Visible = true;

            // add a button to the toolbar
            Office.CommandBarButton commandBarBtn = (Office.CommandBarButton)commandBar.Controls.Add(Office.Enums.MsoControlType.msoControlButton, Missing.Value, Missing.Value, Missing.Value, true);
            commandBarBtn.Style = Office.Enums.MsoButtonStyle.msoButtonIconAndCaption;
            commandBarBtn.Caption = "WordButton1";
            commandBarBtn.FaceId = 2;
            commandBarBtn.ClickEvent += new Office.CommandBarButton_ClickEventHandler(commandBarBtn_ClickEvent);

            wordApp.DisposeChildInstances(false);
        }

        private void CreatePowerPointUI()
        {
            PowerPoint.Application powerApp = _parent.Application.Application as PowerPoint.Application;
            Office.CommandBar commandBar = powerApp.CommandBars.Add("SuperAddinCommandBar", Office.Enums.MsoBarPosition.msoBarTop, false, true);
            commandBar.Visible = true;

            // add a button to the toolbar
            Office.CommandBarButton commandBarBtn = (Office.CommandBarButton)commandBar.Controls.Add(Office.Enums.MsoControlType.msoControlButton, Missing.Value, Missing.Value, Missing.Value, true);
            commandBarBtn.Style = Office.Enums.MsoButtonStyle.msoButtonIconAndCaption;
            commandBarBtn.Caption = "PowerButton1";
            commandBarBtn.FaceId = 2;
            commandBarBtn.ClickEvent += new Office.CommandBarButton_ClickEventHandler(commandBarBtn_ClickEvent);

            powerApp.DisposeChildInstances(false);
        }

        private void CreateAccessUI()
        {
            Access.Application accessApp = _parent.Application.Application as Access.Application;
            Office.CommandBar commandBar = accessApp.CommandBars.Add("SuperAddinCommandBar", Office.Enums.MsoBarPosition.msoBarTop, false, true);
            commandBar.Visible = true;

            // add a button to the toolbar
            Office.CommandBarButton commandBarBtn = (Office.CommandBarButton)commandBar.Controls.Add(Office.Enums.MsoControlType.msoControlButton, Missing.Value, Missing.Value, Missing.Value, true);
            commandBarBtn.Style = Office.Enums.MsoButtonStyle.msoButtonIconAndCaption;
            commandBarBtn.Caption = "AccessButton1";
            commandBarBtn.FaceId = 2;
            commandBarBtn.ClickEvent += new Office.CommandBarButton_ClickEventHandler(commandBarBtn_ClickEvent);;

            accessApp.DisposeChildInstances(false);
        }

        #endregion
    }
}
