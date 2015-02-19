using System;
using System.IO;
using System.Collections.Generic;
using System.ComponentModel;
using System.Drawing;
using System.Data;
using System.Linq;
using System.Text;
using System.Media;
using System.Windows.Forms;
using ICSharpCode.SharpZipLib;
using ICSharpCode.SharpZipLib.Zip;

namespace NetOffice.DeveloperToolbox.ToolboxControls.About
{
    /// <summary>
    /// A small easter egg to make people smile :)
    /// </summary>
    public partial class EasterEggControl : UserControl
    {
        #region Fields

        private SoundPlayer _playerGunshot;
        private SoundPlayer _playerWait;
        private int _lcid = 1033;
        private Stream _waitStream;
        private string _messageDefault = "Thanks for using Developer Toolbox";
        private string _messageGerman = "Java Script? Dieses beschränkte kleine...";

        #endregion

        #region Ctor

        /// <summary>
        /// Creates an instance of the class
        /// </summary>
        public EasterEggControl()
        {
            InitializeComponent();
        }

        /// <summary>
        /// Creates an instance of the class
        /// </summary>
        /// <param name="lcid">current user language id</param>
        public EasterEggControl(int lcid)
        {
            InitializeComponent();
            _lcid = lcid;
        }

        #endregion

        #region Events

        /// <summary>
        /// Egg is done and want close
        /// </summary>
        internal event EventHandler Done;

        private void RaiseDone()
        {
            if (null != Done)
                Done(this, EventArgs.Empty);
        }

        #endregion

        #region Methods

        /// <summary>
        /// Show the angry old men
        /// </summary>
        internal void ShowGernot()
        {
            try
            {
                CreatePlayers();

                string txt = _lcid == 1031 ? _messageGerman : _messageDefault;

                List<Control> list = new List<Control>();
                int i = 0;
                foreach (var item in txt.Split(new string[] { " " }, StringSplitOptions.RemoveEmptyEntries))
                {
                    Label text = new Label();
                    text.AutoSize = true;
                    text.Location = new Point(i + 4, 30);
                    text.Font = panelMessage.Font;
                    text.ForeColor = panelMessage.ForeColor;
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
                int delayTicks = 0;
                bool playWait = false;
                DateTime playWaitStart = DateTime.Now;
                timerText.Tick += delegate(object sender, EventArgs e)
                {
                    if (ctrlIndex >= list.Count)
                    {
                        delayTicks++;
                        if (delayTicks < 5)
                            return;
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
                                DisposePlayers();
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

                if (null == _waitStream)
                { 
                    System.IO.Stream waitZip = a.GetManifestResourceStream("NetOffice.DeveloperToolbox.ToolboxControls.About.Wait.zip");
                    waitZip.Seek(0, SeekOrigin.Begin);
                    ZipFile file = new ZipFile(waitZip);
                    var waitFile = file.GetInputStream(0);
                    _waitStream = waitFile;
                }
                _playerWait = new SoundPlayer(_waitStream);
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

        #endregion

        #region Trigger

        private void EasterEggControl_Resize(object sender, EventArgs e)
        {
            try
            {
                pictureBoxGernot.Location = new Point(Width / 2 - pictureBoxGernot.Width / 2, Height / 2 - pictureBoxGernot.Height / 2);
                panelMessage.Location = new Point(pictureBoxGernot.Left, pictureBoxGernot.Top - panelMessage.Height);
            }
            catch (Exception exception)
            {
                Forms.ErrorForm.ShowError(this, exception,ErrorCategory.NonCritical);
            }
        }

        #endregion
    }
}
