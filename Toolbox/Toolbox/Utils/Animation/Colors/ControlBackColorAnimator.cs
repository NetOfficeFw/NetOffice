using System;
using System.ComponentModel;
using System.Drawing;
using System.Windows.Forms;

namespace NetOffice.DeveloperToolbox.Utils.Animation.Colors
{
    public class ControlBackColorAnimator : AnimatorBase
    {
        #region Fields

        private Control _control;
        private Color _startColor;
        private Color _endColor;

        #endregion

        #region Constructors

        /// <summary>
        /// Creates a new instance.
        /// </summary>
        /// <param name="container">Container the new instance should be added to.</param>
        public ControlBackColorAnimator(IContainer container)
            : base(container)
        {
            Initialize();
        }

        /// <summary>
        /// Creates a new instance.
        /// </summary>
        public ControlBackColorAnimator()
        {
            Initialize();
        }

        private void Initialize()
        {
            _startColor = DefaultStartColor;
            _endColor = DefaultEndColor;
        }

        #endregion

        #region Public interface

        /// <summary>
        /// Gets or sets the starting color for the animation.
        /// </summary>
        [Browsable(true), Category("Appearance")]
        [Description("Gets or sets the starting color for the animation.")]
        public Color StartColor
        {
            get { return _startColor; }
            set
            {
                if (_startColor == value)
                    return;

                _startColor = value;

                OnStartValueChanged(EventArgs.Empty);
            }
        }

        /// <summary>
        /// Gets or sets the ending Color for the animation.
        /// </summary>
        [Browsable(true), Category("Appearance")]
        [Description("Gets or sets the ending Color for the animation.")]
        public Color EndColor
        {
            get { return _endColor; }
            set
            {
                if (_endColor == value)
                    return;

                _endColor = value;

                OnEndValueChanged(EventArgs.Empty);
            }
        }

        /// <summary>
        /// Gets or sets the <see cref="Control"/> which 
        /// <see cref="System.Windows.Forms.Control.BackColor"/> should be animated.
        /// </summary>
        [Browsable(true), Category("Behavior")]
        [DefaultValue(null), RefreshProperties(RefreshProperties.Repaint)]
        [Description("Gets or sets which Control should be animated.")]
        public virtual Control Control
        {
            get { return _control; }
            set
            {
                if (_control == value)
                    return;

                if (_control != null)
                    _control.BackColorChanged -= new EventHandler(OnCurrentValueChanged);

                _control = value;

                if (_control != null)
                    _control.BackColorChanged += new EventHandler(OnCurrentValueChanged);

                base.ResetValues();
            }
        }

        #endregion

        #region Overridden from AnimatorBase

        /// <summary>
        /// Gets or sets the currently shown value.
        /// </summary>
        protected override object CurrentValueInternal
        {
            get { return _control == null ? Color.Empty : _control.BackColor; }
            set
            {
                if (_control != null)
                    _control.BackColor = (Color)value;
            }
        }

        /// <summary>
        /// Gets or sets the starting value for the animation.
        /// </summary>
        public override object StartValue
        {
            get { return StartColor; }
            set { StartColor = (Color)value; }
        }

        /// <summary>
        /// Gets or sets the ending value for the animation.
        /// </summary>
        public override object EndValue
        {
            get { return EndColor; }
            set { EndColor = (Color)value; }
        }

        /// <summary>
        /// Calculates an interpolated value between <see cref="StartValue"/> and
        /// <see cref="EndValue"/> for a given step in %.
        /// Giving 0 will return the <see cref="StartValue"/>.
        /// Giving 100 will return the <see cref="EndValue"/>.
        /// </summary>
        /// <param name="step">Animation step in %</param>
        /// <returns>Interpolated value for the given step.</returns>
        protected override object GetValueForStep(double step)
        {
            if (_startColor == Color.Empty || _endColor == Color.Empty)
                return CurrentValue;

            return InterpolateColors(_startColor, _endColor, step);
        }

        #endregion

        #region Protected

        /// <summary>
        /// Gets the default value of the <see cref="StartColor"/> property.
        /// </summary>
        protected virtual Color DefaultStartColor
        {
            get { return Color.Empty; }
        }

        /// <summary>
        /// Gets the default value of the <see cref="EndColor"/> property.
        /// </summary>
        protected virtual Color DefaultEndColor
        {
            get { return Color.Empty; }
        }

        /// <summary>
        /// Indicates the designer whether <see cref="StartColor"/> needs
        /// to be serialized.
        /// </summary>
        protected virtual bool ShouldSerializeStartColor()
        {
            return _startColor != DefaultStartColor;
        }

        /// <summary>
        /// Indicates the designer whether <see cref="EndColor"/> needs
        /// to be serialized.
        /// </summary>
        protected virtual bool ShouldSerializeEndColor()
        {
            return _endColor != DefaultEndColor;
        }

        #endregion
    }
}
