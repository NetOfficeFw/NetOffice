using System;
using System.Runtime.InteropServices;
using NetOffice.Tools;
using NetOffice.OfficeApi;
using NetOffice.OfficeApi.Enums;
using NetOffice.OfficeApi.Tools;
using NOTools.DeveloperAddin.Logic;
using NOTools.InMemoryCompiler;

namespace NOTools.DeveloperAddin
{
    /// <summary>
    /// addin connect class, inherites from NetOffice.OfficeApi.Tools.COMAddin to connect the addin in all office applications
    /// </summary>
    [COMAddin("NetOffice Developer Addin", "A task pane which allows you to manipulate the automation model at runtime", 3)]
    [ProgId("NetOfficeDeveloperAddin.Addin"), Guid("BA38FD48-47BD-43de-8177-0D067A01B566"), CustomUI("NetOfficeDeveloperAddin.UI.RibbonUI.xml")]
    [MultiRegister(RegisterIn.Excel, RegisterIn.Word, RegisterIn.Outlook, RegisterIn.PowerPoint, RegisterIn.Access, RegisterIn.MSProject)]
    public class Addin : COMAddin
    {
        /// <summary>
        /// creates an instance of the class
        /// </summary>
        public Addin()
        {
            Factory.Settings.ExceptionMessage = "#Error";
            this.OnStartupComplete += new OnStartupCompleteEventHandler(Addin_OnStartupComplete);
            this.OnDisconnection += new OnDisconnectionEventHandler(Addin_OnDisconnection);

            TaskPanes.Add(typeof(UI.DeveloperPane), "NetOffice Developer Pane");
            TaskPanes[0].DockPosition = MsoCTPDockPosition.msoCTPDockPositionRight;
            TaskPanes[0].DockPositionRestrict = MsoCTPDockPositionRestrict.msoCTPDockPositionRestrictNoChange;
            TaskPanes[0].Width = 320;
            TaskPanes[0].Visible = true;
            TaskPanes[0].Arguments = new object[] { this };           
        }

        internal DynamicCommandDefinitionCollection Commands { get; private set; }
       
        internal IRibbonUI RibbonUI { get; private set; }

        public void OnLoadRibbonUI(IRibbonUI ribbonUI)
        {
            RibbonUI = ribbonUI;
            (TaskPaneInstances[0] as UI.DeveloperPane).ParentVisibleChanged += new EventHandler(Addin_ParentVisibleChanged);
        }
      
        public void OnCheckActionToogleButton(IRibbonControl control, bool check)
        {
            TaskPanes[0].Pane.Visible = check;
        }

        public bool GetPressedToogleButton(IRibbonControl control)
        {
            return TaskPanes[0].Pane.Visible;
        }

        private void Addin_ParentVisibleChanged(object sender, EventArgs e)
        {
            RibbonUI.InvalidateControl("toogleButtongroupNetOfficeDeveloperAddin");
        }

        void Addin_OnDisconnection(ext_DisconnectMode RemoveMode, ref Array custom)
        {
            if (null != Commands)
                Commands.SaveToFile(Application);
        }

        void Addin_OnStartupComplete(ref Array custom)
        {
            Commands = new DynamicCommandDefinitionCollection();
            Commands.LoadFromFile(Application);
        }

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
