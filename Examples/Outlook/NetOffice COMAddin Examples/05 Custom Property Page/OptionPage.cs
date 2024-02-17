using System;
using System.Runtime.InteropServices;
using System.Windows.Forms;
using NetOffice;
using Outlook = NetOffice.OutlookApi;

namespace Outlook05AddinCS4
{
    public partial class OptionPage : UserControl, Outlook.Native.PropertyPage
    {
        public OptionPage(Core core)
        {
            InitializeComponent();
            DataSource = core.Settings;
            EditSource = new Settings(core.Settings);
            SettingsGrid.SelectedObject = EditSource;
            EditSource.PropertyChanged += delegate
            {
                 PageContainer?.OnStatusChange();
            };
        }

        private Settings DataSource { get; set; }

        private Settings EditSource { get; set; }
        
        private Outlook.Native.PropertyPageSite PageContainer { get; set; }

        protected override void OnLoad(EventArgs e)
        {
            base.OnLoad(e);
            PageContainer = Outlook.Tools.Contribution.ApplicationUtils.TryGetPageContainer(this);
        }

        public bool Dirty
        {
            get
            {
                return !DataSource.IsEqualTo(EditSource);
            }
        }

        public void Apply()
        {
            if (Dirty)
            {
                DataSource.LoadFrom(EditSource);
            }
        }

        public void GetPageInfo(ref string HelpFile, ref int HelpContext)
        {
            
        }
    }
}
