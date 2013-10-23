using System;
using System.Collections.Generic;
using System.Windows.Forms;
using NetOffice;
using NetOffice.OfficeApi.Tools;
using NOTools.CodeCommander.Logic;

namespace NOTools.CodeCommander.UI
{
    /// <summary>
    /// the shown taskpane in office
    /// </summary>
    public partial class DeveloperPane : UserControl, ITaskPane
    {
        // Ctor
        public DeveloperPane()
        {
            InitializeComponent();
        }
        
        // events

        public event EventHandler ParentVisibleChanged;

        private void RaiseParentVisibleChanged()
        {
            if (null != ParentVisibleChanged)
                ParentVisibleChanged(this, new EventArgs());
        }
         
        // Properties

        private NetOffice.OfficeApi._CustomTaskPane ParentPane { get; set; }

        public bool ParentPaneVisible
        {
            get
            {
                if (null != ParentPane)
                    return ParentPane.Visible;
                else
                    return false;
            }
        }

        private bool LastParentPaneVisible { get; set; }

        // ITaskpane Member

        public void OnConnection(COMObject application, NetOffice.OfficeApi._CustomTaskPane parentPane, object[] customArguments)
        {
            ParentPane = parentPane;
            LastParentPaneVisible = parentPane.Visible;
            commandPane1.OnConnection(application, parentPane, customArguments);
            propertyPane1.OnConnection(application, parentPane, customArguments);
            infoPane1.OnConnection(application, parentPane, customArguments);
        }

        public void OnDisconnection()
        {
            commandPane1.OnDisconnection();
            propertyPane1.OnDisconnection();
            infoPane1.OnDisconnection();
        }   

        // Event Trigger
  
        private void DirtyLittleTimer_Tick(object sender, EventArgs e)
        {
            // not nice so far but no other way to fire a visible change event to update the toogle buttons
            if (null == ParentPane)
                return;
            if (ParentPane.Visible != LastParentPaneVisible)
            {
                RaiseParentVisibleChanged();
                LastParentPaneVisible = ParentPane.Visible;
            }
        }
    }
}
