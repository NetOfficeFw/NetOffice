using System;
using System.Drawing;
using System.Diagnostics;
using System.Windows.Forms;

namespace TutorialsBase
{
    public partial class AreaForm : Form
    {
        private ITutorial _tutorial;

        public AreaForm(ITutorial tutorial)
        {
            InitializeComponent();
            _tutorial = tutorial;
            Text = tutorial.Caption;
            if (OptionsForm.ConnectToDocumentation)
            {
                tabControl1.TabPages.Remove(OfflineTabPage);
                webBrowserTutorialContent.Url = new Uri(tutorial.Uri);
                WindowState = FormWindowState.Maximized;
            }
            else
            {
                tabControl1.TabPages.Remove(OnlineTabPage);
                linkLabelTutorialContent.Text = tutorial.Uri;
            }
            if (null != tutorial.Panel)
            {
                tabControl1.TabPages.Remove(SampleTabPage);
                AreaTabPage.Controls.Add(tutorial.Panel);
                tutorial.Panel.Dock = DockStyle.Fill;
            }
            else
            {
                tabControl1.TabPages.Remove(AreaTabPage);
            }
        }

        public static void ShowForm(IWin32Window owner, ITutorial tutorial)
        {
            AreaForm form = new AreaForm(tutorial);
            form.ShowDialog(owner);
            form.DoDispose();
            tutorial.Disconnect();
        }

        internal void DoDispose()
        {
            AreaTabPage.Controls.Clear();
            Dispose();
        }

        private void DoResize()
        {
            buttonRunTutorial.Location = new Point(
               (Width / 2) - (buttonRunTutorial.Width / 2),
               ((Height / 2) - (buttonRunTutorial.Height / 2))-20);
        }

        private void panelTutorials_Resize(object sender, System.EventArgs e)
        {
            DoResize();
        }

        private void tabControl1_SelectedIndexChanged(object sender, System.EventArgs e)
        {
            DoResize();
        }

        private void linkLabelTutorialContent_LinkClicked(object sender, LinkLabelLinkClickedEventArgs e)
        {
            try
            {
                LinkLabel label = sender as LinkLabel;
                Process.Start(label.Text);
            }
            catch (Exception)
            {
                ;
            }
        }

        private void buttonRunTutorial_Click(object sender, EventArgs e)
        {
            try
            {
                _tutorial.Run();
            }
            catch (Exception exception)
            {
                ErrorForm.Show(this, null, exception.Message, exception);
            }
        }
    }
}
