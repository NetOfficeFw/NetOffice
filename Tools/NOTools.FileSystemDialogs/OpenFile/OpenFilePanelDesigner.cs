using System;
using System.ComponentModel;
using System.ComponentModel.Design;
using System.Collections.Generic;
using System.Windows.Forms;
using System.Windows.Forms.Design;
using System.Text;

namespace NOTools.FileSystemDialogs
{
    /// <summary>
    /// Designtime helper to the IDE
    /// </summary>
    internal class OpenFilePanelDesigner : ControlDesigner
    {
        #region Fields

        private OpenFilePanel _panel;
        
        #endregion

        #region Overrides

        public override void Initialize(IComponent component)
        {
            base.Initialize(component);
            _panel = component as OpenFilePanel;
        }

        public override System.ComponentModel.Design.DesignerVerbCollection Verbs
        {
            get
            {
                DesignerVerbCollection verbs = new DesignerVerbCollection();
                if (null != _panel)
                {
                    verbs.Add(new DesignerVerb("Restore En-Us(1033) Default Localization", new EventHandler(Set1033Default)));
                    verbs.Add(new DesignerVerb("Restore De(1031) Default Localization", new EventHandler(Set1031Default)));
                }
                return verbs;
            }
        }

        #endregion

        #region Methods

        private void Set1033Default(object sender, System.EventArgs e)
        {
            if (null == _panel)
                return;

            IDesignerHost host = GetService(typeof(IDesignerHost)) as IDesignerHost;
            IComponentChangeService service = (IComponentChangeService)GetService(typeof(IComponentChangeService));
            if (null != host && null != service)
            {
                DesignerTransaction transaction = host.CreateTransaction("Set1033Default");
                service.OnComponentChanging(_panel, null);
                _panel.Localization.Set1033Default(sender, e);
                service.OnComponentChanged(_panel, null, null, null);
                transaction.Commit();
            }
        }

        private void Set1031Default(object sender, System.EventArgs e)
        {
            if (null == _panel)
                return;

            IDesignerHost host = GetService(typeof(IDesignerHost)) as IDesignerHost;
            IComponentChangeService service = (IComponentChangeService)GetService(typeof(IComponentChangeService));
            if(null != host && null!= service)
            {
                DesignerTransaction transaction = host.CreateTransaction("Set1031Default");
                service.OnComponentChanging(_panel, null);
                _panel.Localization.Set1031Default(sender, e);
                service.OnComponentChanged(_panel, null, null, null);
                transaction.Commit();
           }
        }

        #endregion
    }
}
