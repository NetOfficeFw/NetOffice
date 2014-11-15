using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Drawing;
using System.Data;
using System.Linq;
using System.Text;
using System.Media;
using System.Windows.Forms;

namespace NetOffice.DeveloperToolbox.ToolboxControls.About
{
    public partial class EasterEggControl : UserControl
    {
        private SoundPlayer _playerGunshot;
        private SoundPlayer _playerWait;
        private int _lcid = 1033;

        public EasterEggControl()
        {
            InitializeComponent();
        }

        public EasterEggControl(int lcid)
        {
            InitializeComponent();
            _lcid = lcid;
        }

        internal event EventHandler Done;

        private void RaiseDone()
        {
            if (null != Done)
                Done(this, EventArgs.Empty);
        }

        internal void ShowwGernot()
        {
            try
            {
                CreatePlayers();

                string txt = _lcid == 1031 ? "Nun aber an die Arbeit" : "Thanks for using Developer Toolbox";

                List<Control> list = new List<Control>();
                int i = 140;
                foreach (var item in txt.Split(new string[] { " " }, StringSplitOptions.RemoveEmptyEntries))
                {
                    Label text = new Label();
                    text.AutoSize = true;
                    text.Location = new Point(i + 10, 30);
                    text.Font = panelMessage.Font;
                    text.Text = item;
                    text.Visible = true;
                    list.Add(text);
                    i += TextRenderer.MeasureText(item, text.Font).Width;
                }

                Utils.Animation.Effects.EffectsAnimator.DoEffect(pictureBoxGernot, Utils.Animation.Effects.EffectsKind.Collapse, true, 250);

                Timer timerText = new Timer();
                timerText.Interval = 400;
                timerText.Enabled = true;
                int ctrlIndex = 0;
                bool playWait = false;
                DateTime playWaitStart = DateTime.Now;
                timerText.Tick += delegate(object sender, EventArgs e)
                {
                    if (ctrlIndex >= list.Count)
                    {
                        if (!playWait)
                        {
                            pictureBoxWait.Visible = true;
                            pictureBoxWait.BringToFront();
                            panelMessage.Visible = false;
                            pictureBoxGernot.Visible = false;
                            playWait = true;
                            playWaitStart = DateTime.Now;
                            PlayWait();
                        }
                        else
                        {
                            if (DateTime.Now.Subtract(playWaitStart).TotalSeconds >= 18.0)
                            {
                                timerText.Enabled = false;
                                RaiseDone();
                                return;
                            }
                        }
                    }
                    else
                    {
                        PlayGunshot();
                        Control ctrl = list[ctrlIndex];
                        panelMessage.Controls.Add(ctrl);
                        ctrlIndex++;
                    }
                };
            }
            catch
            {
                RaiseDone();
            }
        }

        private void PlayWait()
        {
            try
            {
                if (null != _playerWait)
                    _playerWait.Play();
            }
            catch
            {
                ;
            }
        }

        private void PlayGunshot()
        {
            try
            {
                if (null != _playerGunshot)
                    _playerGunshot.Play();
            }
            catch
            {
                ;
            }
        }

        private void CreatePlayers()
        {
            try
            {
                DisposePlayers();
                System.Reflection.Assembly a = System.Reflection.Assembly.GetExecutingAssembly();
                System.IO.Stream s1 = a.GetManifestResourceStream("NetOffice.DeveloperToolbox.ToolboxControls.About.Gunshot.wav");
                s1.Seek(0, System.IO.SeekOrigin.Begin);
                _playerGunshot = new SoundPlayer(s1);
                System.IO.Stream s2 = a.GetManifestResourceStream("NetOffice.DeveloperToolbox.ToolboxControls.About.Wait.wav");
                s2.Seek(0, System.IO.SeekOrigin.Begin);
                _playerWait = new SoundPlayer(s2);
            }
            catch
            {
                ;
            }
        }

        private void DisposePlayers()
        {
            try
            {
                if (null != _playerWait)
                {
                    _playerWait.Stop();
                    _playerWait.Dispose();
                    _playerWait = null;
                }

                if (null != _playerGunshot)
                {
                    _playerGunshot.Stop();
                    _playerGunshot.Dispose();
                    _playerGunshot = null;
                }
            }
            catch
            {
                ;                
            }          
        }

        private void EasterEggControl_Resize(object sender, EventArgs e)
        {
            try
            {
                pictureBoxGernot.Location = new Point(Width / 2 - pictureBoxGernot.Width / 2, Height / 2 - pictureBoxGernot.Height / 2);
                panelMessage.Location = new Point(pictureBoxGernot.Left, pictureBoxGernot.Top - panelMessage.Height);
            }
            catch (Exception exception)
            {
                Forms.ErrorForm.ShowError(this, exception, Forms.ErrorCategory.NonCritical);
            }
        }
    }
}
