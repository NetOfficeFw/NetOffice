using System;
using System.ComponentModel;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Drawing;
using System.Drawing.Imaging;
using System.Drawing.Drawing2D;
using System.Windows.Forms;

namespace NetOffice.DeveloperToolbox.Utils.Animation.Panel
{
    public class AnimatedPanel : System.Windows.Forms.Panel
    {
        #region Fields

        private bool _stopDrawing;

        private bool _animation1Rotate = true;
        private int _imageCount1 = 10;
        private Image _animationImage1;
        private int _animationIntervall1 = 40;
        private bool _animationEnabled1 = false;
        private float _opacity1 = 1.0f;
        private Timer _animationTimer1;
        private List<Shape> _shapes1 = new List<Shape>();

        private bool _animation2Rotate = true;
        private int _imageCount2 = 10;
        private Image _animationImage2;
        private int _animationIntervall2 = 40;
        private bool _animationEnabled2 = false;
        private float _opacity2 = 1.0f;
        private Timer _animationTimer2;
        private List<Shape> _shapes2 = new List<Shape>();

        #endregion

        #region Ctor

        public AnimatedPanel()
        {
            SetStyle(ControlStyles.AllPaintingInWmPaint | ControlStyles.DoubleBuffer | ControlStyles.ResizeRedraw | ControlStyles.UserPaint, true);
           
            _shapes1 = new List<Shape>();
            Animation1ImageCount = 10;
            _animationTimer1 = new Timer();
            _animationTimer1.Interval = _animationIntervall1;
            _animationTimer1.Tick += new EventHandler(AnimationTimer1_Tick);

            _shapes2 = new List<Shape>();
            Animation2ImageCount = 10;
            _animationTimer2 = new Timer();
            _animationTimer2.Interval = _animationIntervall2;
            _animationTimer2.Tick += new EventHandler(AnimationTimer2_Tick);
        }
       
        #endregion

        #region Properties

        [DefaultValue(true), Category("Animation1")]
        public bool Animation1Rotate
        {
            get
            {
                return _animation1Rotate;
            }
            set 
            {
                _animation1Rotate = value;
                foreach (var item in _shapes1)
                    item.Rotate = value;
            }
        }

        [DefaultValue(null), Category("Animation1")]
        public Image Animation1Image
        {
            get
            {
                return _animationImage1;
            }
            set
            {
                _animationImage1 = value;
                foreach (var item in _shapes1)
                {
                    ImageShape img = item as ImageShape;
                    if (null != img)
                    {
                        if(null != value)
                            img.Size = new System.Drawing.Size(value.Width, value.Height);
                        img.Image = value;
                    }
                }             
            }
        }

        [DefaultValue(10), Category("Animation1")]
        public int Animation1ImageCount
        {
            get
            {
                return _imageCount1;
            }
            set
            {
                if (value < 1)
                    throw new ArgumentException();
                if (value > 256)
                    throw new ArgumentException();
                try
                {
                    _stopDrawing = true;

                    _imageCount1 = value;
                    _shapes1.Clear();
                    Random rnd = new Random();
                    for (int i = 0; i < _imageCount1; i++)
                    {
                        ImageShape shape = new ImageShape();
                        shape.Limits = GetLimits();
                        shape.Location = new Point(rnd.Next(this.ClientRectangle.Width), rnd.Next(this.ClientRectangle.Height));
                        if (null != Animation1Image)
                            shape.Size = new Size(Animation1Image.Width, Animation1Image.Height);
                        else
                            shape.Size = new Size(32, 32);
                        shape.RotationDelta = (float)rnd.Next(20);
                        shape.Vector = new Size(-10 + rnd.Next(20), -10 + rnd.Next(20));
                        shape.Image = Animation1Image;
                        shape.Rotate = _animation1Rotate;
                        _shapes1.Add(shape);
                    }
                }
                catch (Exception)
                {

                    throw;
                }

                finally
                {
                    _stopDrawing = false;
                }
            }
        }

        [DefaultValue(40), Category("Animation1")]
        public int Animation1Intervall
        {
            get
            {
                return _animationIntervall1;
            }
            set 
            {
                if (_animationIntervall1 < 1)
                    throw new ArgumentException();
                _animationIntervall1 = value;
                _animationTimer1.Interval = value;

            }
        }

        [DefaultValue(false), Category("Animation1")]
        public bool Animation1Enabled
        {
            get
            {
                return _animationEnabled1;
            }
            set
            {
                _animationEnabled1 = value;
                _animationTimer1.Enabled = value;
            }
        }

        [DefaultValue(1.0f), Category("Animation1")]
        public float Animation1Opacity
        {
            get
            {
                return _opacity1;
            }
            set
            {
                if (value < 0.0f || value > 1.0f)
                    throw new ArgumentException();
                _opacity1 = value;
                foreach (var item in _shapes1)
                {
                    ImageShape shp = item as ImageShape;
                    if (null != shp)
                        shp.Opacity = _opacity1;
                }
            }
        }

        [DefaultValue(true), Category("Animation2")]
        public bool Animation2Rotate
        {
            get
            {
                return _animation2Rotate;
            }
            set
            {
                _animation2Rotate = value;
                foreach (var item in _shapes2)
                    item.Rotate = value;
            }
        }

        [DefaultValue(null), Category("Animation2")]
        public Image Animation2Image
        {
            get
            {
                return _animationImage2;
            }
            set
            {
                _animationImage2 = value;
                foreach (var item in _shapes2)
                {
                    ImageShape img = item as ImageShape;
                    if (null != img)
                    {
                        if (null != value)
                            img.Size = new System.Drawing.Size(value.Width, value.Height);
                        img.Image = value;
                    }
                }
            }
        }

        [DefaultValue(10), Category("Animation2")]
        public int Animation2ImageCount
        {
            get
            {
                return _imageCount2;
            }
            set
            {
                if (value < 1)
                    throw new ArgumentException();
                if (value > 256)
                    throw new ArgumentException();
                try
                {
                    _stopDrawing = true;

                    _imageCount2 = value;
                    _shapes2.Clear();
                    Random rnd = new Random();
                    for (int i = 0; i < _imageCount2; i++)
                    {
                        ImageShape shape = new ImageShape();
                        shape.Limits = GetLimits();
                        shape.Location = new Point(rnd.Next(this.ClientRectangle.Width), rnd.Next(this.ClientRectangle.Height));
                        if (null != Animation1Image)
                            shape.Size = new Size(Animation1Image.Width, Animation1Image.Height);
                        else
                            shape.Size = new Size(32, 32);
                        shape.RotationDelta = (float)rnd.Next(20);
                        shape.Vector = new Size(-10 + rnd.Next(20), -10 + rnd.Next(20));
                        shape.Image = Animation2Image;
                        shape.Rotate = _animation2Rotate;
                        _shapes2.Add(shape);
                    }
                }
                catch (Exception)
                {

                    throw;
                }

                finally
                {
                    _stopDrawing = false;
                }
            }
        }

        [DefaultValue(40), Category("Animation2")]
        public int Animation2Intervall
        {
            get
            {
                return _animationIntervall2;
            }
            set
            {
                if (_animationIntervall2 < 1)
                    throw new ArgumentException();
                _animationIntervall2 = value;
                _animationTimer2.Interval = value;
            }
        }

        [DefaultValue(false), Category("Animation2")]
        public bool Animation2Enabled
        {
            get
            {
                return _animationEnabled2;
            }
            set
            {
                _animationEnabled2 = value;
                _animationTimer2.Enabled = value;
            }
        }

        [DefaultValue(1.0f), Category("Animation2")]
        public float Animation2Opacity
        {
            get
            {
                return _opacity2;
            }
            set
            {
                if (value < 0.0f || value > 1.0f)
                    throw new ArgumentException();
                _opacity2 = value;
                foreach (var item in _shapes2)
                {
                    ImageShape shp = item as ImageShape;
                    if (null != shp)
                        shp.Opacity = _opacity2;
                }
            }
        }

        #endregion

        #region Methods

        private Rectangle GetLimits()
        {
            return new Rectangle(ClientRectangle.X, ClientRectangle.Y, ClientRectangle.Width, ClientRectangle.Height);
        }

        #endregion

        #region Overrides

        protected override void OnPaint(PaintEventArgs e)
        {
            if (true == _animationTimer1.Enabled && false == _stopDrawing)
            {
                foreach (Shape s in this._shapes2)
                    s.Draw(e.Graphics);

                foreach (Shape s in this._shapes1)
                    s.Draw(e.Graphics);
            }
        }

        protected override void Dispose(bool disposing)
        {
            _animationTimer1.Enabled = false;
            _animationTimer1.Enabled = false;
            _animationTimer2.Enabled = false;
            _animationTimer2.Enabled = false;
            base.Dispose(disposing);
        }

        protected override void OnSizeChanged(EventArgs e)
        {
            if (false == _stopDrawing)
            {
                foreach (Shape s in this._shapes1)
                    s.Limits = GetLimits();

                foreach (Shape s in this._shapes2)
                    s.Limits = GetLimits();
            }
            base.OnSizeChanged(e);
        }

        #endregion

        #region Trigger

        private void AnimationTimer1_Tick(object sender, EventArgs e)
        {
            if (_stopDrawing)
                return;
            foreach (Shape s in _shapes1)
                s.Tick();
            Invalidate();
        }

        private void AnimationTimer2_Tick(object sender, EventArgs e)
        {
            if (_stopDrawing)
                return;
            foreach (Shape s in _shapes2)
                s.Tick();
            Invalidate();
        }

        #endregion
    }
}
