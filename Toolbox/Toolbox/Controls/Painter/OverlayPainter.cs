using System;
using System.ComponentModel;
using System.Drawing;
using System.Windows.Forms;

namespace NetOffice.DeveloperToolbox.Controls.Painter
{
    /// <summary>
    /// Support Painter to create a paint event which is possible to use as overlayer
    /// </summary>
    public partial class OverlayPainter : Component
    {
        #region Fields

        private Form _form;

        #endregion

        #region Ctor

        /// <summary>
        /// Creates an instance of the class
        /// </summary>
        public OverlayPainter()
        {
            InitializeComponent();
            this.Disposed += new EventHandler(OverlayPainter_Disposed);
        }

        /// <summary>
        /// Creates an instance of the class
        /// </summary>
        /// <param name="container">parent container</param>
        public OverlayPainter(IContainer container)
        {
            container.Add(this);
            InitializeComponent();
            this.Disposed += new EventHandler(OverlayPainter_Disposed);
        }

        #endregion

        #region Events

        /// <summary>
        /// Paint event to draw on top as overlayer
        /// </summary>
        public event EventHandler<PaintEventArgs> Paint;

        #endregion

        #region Properties

        /// <summary>
        /// Top level window
        /// </summary>
        [DesignerSerializationVisibility(DesignerSerializationVisibility.Hidden)]
        public Form Owner
        {
            get { return _form; }
            set
            {
                if (value == null)
                    throw new ArgumentNullException();
                if (_form != null)
                    throw new InvalidOperationException();

                _form = value;
                _form.Resize += new EventHandler(Form_Resize);
                ConnectPaintEventHandlers(_form);
            }
        }

        #endregion

        #region Methods

        private void ConnectPaintEventHandlers(Control control)
        {
            control.Paint -= new PaintEventHandler(Control_Paint);
            control.Paint += new PaintEventHandler(Control_Paint);

            control.ControlAdded -= new ControlEventHandler(Control_ControlAdded);
            control.ControlAdded += new ControlEventHandler(Control_ControlAdded);

            foreach (Control child in control.Controls)
                ConnectPaintEventHandlers(child);
        }

        private void DisconnectPaintEventHandlers(Control control)
        {
            control.Paint -= new PaintEventHandler(Control_Paint);
            control.ControlAdded -= new ControlEventHandler(Control_ControlAdded);
            foreach (Control child in control.Controls)
                DisconnectPaintEventHandlers(child);
        }

        private void OnPaint(object sender, PaintEventArgs e)
        {
            if (Paint != null)
                Paint(sender, e);
        }

        #endregion

        #region Trigger

        private void OverlayPainter_Disposed(object sender, EventArgs e)
        {
            if (null != _form)
                DisconnectPaintEventHandlers(_form);
        }

        private void Form_Resize(object sender, EventArgs e)
        {
            if(null != _form)
                _form.Invalidate(true);
        }

        private void Control_ControlAdded(object sender, ControlEventArgs e)
        {
            ConnectPaintEventHandlers(e.Control);
        }

        private void Control_Paint(object sender, PaintEventArgs e)
        {
            if (null == _form || _form.IsDisposed)
                return;

            Control control = sender as Control;
            Point location;

            if (control == _form)
                location = control.Location;
            else
            {
                location = _form.PointToClient(control.Parent.PointToScreen(control.Location));
                location += new Size((control.Width - control.ClientSize.Width) / 2, (control.Height - control.ClientSize.Height) / 2);
            }

            if (control != _form)
                e.Graphics.TranslateTransform(-location.X, -location.Y);

            OnPaint(sender, e);
        }

        #endregion
    }
}

namespace System.Windows.Forms
{
    using System.Drawing;

    public static class Extensions
    {
        /// <summary>
        /// Coordinates from control on toplevel control or desktop
        /// </summary>
        /// <param name="control">target control</param>
        /// <returns>coordinates</returns>
        public static Rectangle Coordinates(this Control control)
        {
            Rectangle coordinates;
            Form form = control.TopLevelControl as Form;

            if (control == form)
                coordinates = form.ClientRectangle;
            else
                coordinates = form.RectangleToClient(control.Parent.RectangleToScreen(control.Bounds));

            return coordinates;
        }
    }
}
