using System;
using System.Runtime.InteropServices;
using NetOffice.Tools;
using NetOffice.OfficeApi;
using NetOffice.OfficeApi.Enums;
using NetOffice.OfficeApi.Tools;
using NOTools.CodeCommander.Logic;
using NOTools.InMemoryCompiler;

namespace NOTools.CodeCommander
{
    /// <summary>
    /// addin connect class, inherites from NetOffice.OfficeApi.Tools.COMAddin to connect the addin in all office applications
    /// </summary>
    [COMAddin("NetOffice Code Commander", "A task pane which allows you to manipulate the automation model at runtime", 3)]
    [ProgId("NOToolsCodeCommander.Addin"), Guid("BA38FD48-47BD-43de-8177-0D067A01B566"), CustomUI("NOTools.CodeCommander.UI.RibbonUI.xml"), Tweak(true)]
    [MultiRegister(RegisterIn.Excel, RegisterIn.Word, RegisterIn.Outlook, RegisterIn.PowerPoint, RegisterIn.Access, RegisterIn.MSProject)]
    public class Addin : COMAddin
    {
        /// <summary>
        /// creates an instance of the class
        /// </summary>
        public Addin()
        {
            Factory.Settings.ExceptionMessage = "#Error";
            Factory.Console.Name = "CodeCommander";

            this.OnStartupComplete += new OnStartupCompleteEventHandler(Addin_OnStartupComplete);
            this.OnDisconnection += new OnDisconnectionEventHandler(Addin_OnDisconnection);

            TaskPanes.Add(typeof(UI.DeveloperPane), "Code Commander");
            TaskPanes[0].DockPosition = MsoCTPDockPosition.msoCTPDockPositionRight;
            TaskPanes[0].DockPositionRestrict = MsoCTPDockPositionRestrict.msoCTPDockPositionRestrictNone;
            TaskPanes[0].Width = 320;
            TaskPanes[0].Visible = true;
            TaskPanes[0].Arguments = new object[] { this };
            TaskPanes[0].VisibleStateChange += new CustomTaskPane_VisibleStateChangeEventHandler(TaskPane_VisibleStateChange);
        }

        #region Properties

        internal DynamicCommandDefinitionCollection Commands { get; private set; }

        internal IRibbonUI RibbonUI { get; private set; }

        #endregion

        #region Ribbon Trigger

        public void OnLoadRibbonUI(IRibbonUI ribbonUI)
        {
            RibbonUI = ribbonUI;
        }

        public void OnCheckActionToogleButton(IRibbonControl control, bool check)
        {
            TaskPanes[0].Pane.Visible = check;
        }

        public bool GetPressedToogleButton(IRibbonControl control)
        {
            return TaskPanes[0].Pane.Visible;
        }

        #endregion

        #region Addin Trigger

        private void TaskPane_VisibleStateChange(_CustomTaskPane CustomTaskPaneInst)
        {
        }

        private void Addin_OnDisconnection(ext_DisconnectMode RemoveMode, ref Array custom)
        {
            if (null != Commands)
                Commands.SaveToFile(Application);
        }

        private void Addin_OnStartupComplete(ref Array custom)
        {
            Commands = new DynamicCommandDefinitionCollection();
            Commands.LoadFromFile(Application);
        }

        #endregion

        /*
         *         DynamicAssembly assembly = new DynamicAssembly("MyDynamicAssembly", 
                new string[] { "NetOfficeDeveloperAddin.dll", "NetOffice.dll", "OfficeApi.dll", "ExcelApi.dll" });
         
            DynamicClass myClass = assembly.Classes.AddNew("MyClass",
                new string[] { "System.IO", "NetOfficeDeveloperAddin", "Office = NetOffice.OfficeApi", "NetOffice.OfficeApi.Enums", "Excel = NetOffice.ExcelApi", "NetOffice.ExcelApi.Enums" });

            myClass.Interfaces.AddNew("NetOfficeDeveloperAddin.Logic.DynamicCommand");

            myClass.Properties.Add("Excel.Application", "Application");

            myClass.Methods.AddNew("Execute", "MessageBox.Show(\"hello\");");

            CompileResult result = CSharpCompiler.CompileDynamicAssembly(assembly);
         */
    }
}
