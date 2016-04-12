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
    internal class Shape
    {
        #region Fields

        private GraphicsState _state;
        private Point _location;
        private Size _size;
        private Color _backColor;
        private Color _foreColor;
        private int _lineThickness;
        private float _transparency = 0.9f;
        private float _rotation;
        private Size _vector;
        private float _rotationDelta;
        private Rectangle _limits;
        private bool _rotate = true;

        #endregion

        #region Properties

        public Point Location
        {
            get { return _location; }
            set { _location = value; }
        }

        public Size Size
        {
            get { return _size; }
            set { _size = value; }
        }

        public Color BackColor
        {
            get { return _backColor; }
            set { _backColor = value; }
        }

        public Color ForeColor
        {
            get { return _foreColor; }
            set { _foreColor = value; }
        }

        public int LineThickness
        {
            get { return _lineThickness; }
            set { _lineThickness = value; }
        }

        public float Transparency
        {
            get { return _transparency; }
            set
            {
                _transparency = (value >= 0 ? (value <= 1 ? value : 1) : 0);
            }
        }

        public float Rotation
        {
            get { return _rotation; }
            set { _rotation = value; }
        }

        public Size Vector
        {
            get { return _vector; }
            set { _vector = value; }
        }

        public float RotationDelta
        {
            get { return _rotationDelta; }
            set { _rotationDelta = value; }
        }

        public Rectangle Limits
        {
            get { return _limits; }
            set { _limits = value; }
        }

        public virtual bool Rotate
        {
            get { return _rotate; }
            set { _rotate = value; }
        }

        #endregion

        #region Methods

        protected void SetupTransform(Graphics g)
        {
            _state = g.Save();
            Matrix mx = new Matrix();
            if(_rotate)
                mx.Rotate(_rotation, MatrixOrder.Append);
            mx.Translate(this.Location.X, this.Location.Y, MatrixOrder.Append);
            g.Transform = mx;
        }

        protected void RestoreTransform(Graphics g)
        {
            g.Restore(_state);
        }

        public void Draw(Graphics g)
        {
            SetupTransform(g);
            RenderObject(g);
            RestoreTransform(g);
        }

        public virtual void Tick()
        {
            if (this.Location.X > this.Limits.Right)
                this.Location = new Point(this.Limits.Right - 1, this.Location.Y);
            if (this.Location.Y > this.Limits.Bottom)
                this.Location = new Point(this.Location.X, this.Limits.Bottom - 1);

            int newx = this.Location.X + this.Vector.Width;
            if (newx > this.Limits.Right || newx < this.Limits.Left)
                this.Vector = new Size(-1 * this.Vector.Width, this.Vector.Height);
            int newy = this.Location.Y + this.Vector.Height;
            if (newy > this.Limits.Bottom || newy < this.Limits.Top)
                this.Vector = new Size(this.Vector.Width, -1 * this.Vector.Height);
        
            Location = new Point(this.Location.X + this.Vector.Width, this.Location.Y + this.Vector.Height);

            this.Rotation += this.RotationDelta;
            Rotation = (Rotation < 360f ? (Rotation >= 0 ? Rotation : Rotation + 360f) : Rotation - 360f);
        }

        public virtual void RenderObject(Graphics g)
        {
        }

        #endregion
    }
}
