using System;
using System.Reflection;
using System.Runtime.InteropServices;
using System.Drawing;
using System.ComponentModel;
using System.ComponentModel.Design;
using System.Collections;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace NetOffice.DeveloperToolbox.Utils.Animation
{
    #region Enum SynchronizationMode

    /// <summary>
    /// Enumeration holding modes for design time support of <see cref="AnimatorBase"/>
    /// and its inheriting classes.
    /// </summary>
    public enum SynchronizationMode
    {
        /// <summary>
        /// No synchronization.
        /// </summary>
        None,
        /// <summary>
        /// Synchronize designer value with starting values.
        /// </summary>
        Start,
        /// <summary>
        /// Synchronize designer value with ending values.
        /// </summary>
        End,
        /// <summary>
        /// Reset starting and ending value to current value.
        /// </summary>
        ResetToCurrent
    }

    #endregion

    #region Enum LoopMode

    /// <summary>
    /// Ernumeration holding the possible modes for looping an animation.
    /// </summary>
    public enum LoopMode
    {
        /// <summary>
        /// No looping (animation stops after reaching end value).
        /// </summary>
        None,
        /// <summary>
        /// Animation gets restarted everytime the end value is reached.
        /// </summary>
        Repeat,
        /// <summary>
        /// The animation loops continueously back and forth between start and endvalue.
        /// </summary>
        Bidirectional
    }

    #endregion

    /// <summary>
    /// Abstract base class for all components animating something.
    /// It holds a <see cref="System.Timers.Timer"/> to control the
    /// animation.
    /// Implementing classes must override the following:
    /// <see cref="StartValue"/>: Getter and setter for concrete starting value of the animation.
    /// <see cref="EndValue"/>: Getter and setter for concrete ending value of the animation.
    /// <see cref="CurrentValue"/>: Getter and setter for concrete value currently showing.
    /// <see cref="GetValueForStep"/>: Function calculating the value for a single animation step.
    /// All handled values must always be of the same type.
    /// Moreover every inheriting class must have two constructors having
    /// the same signature as the constructors provided here.
    /// </summary>
    public abstract class AnimatorBase : System.ComponentModel.Component, System.ComponentModel.ISupportInitialize
    {
        #region Fields

        /// <summary>
        /// The default value for the <see cref="StepSize"/> property.
        /// </summary>
        public const double DEFAULT_STEP_SIZE = 2;

        /// <summary>
        /// The default value for the <see cref="Intervall"/> property.
        /// </summary>
        public const int DEFAULT_INTERVALL = 10;

        /// <summary>
        /// The default value for the <see cref="LoopMode"/> property.
        /// </summary>
        public const LoopMode DEFAULT_LOOP_ANIMATION = LoopMode.None;

        private const bool DEFAULT_NEVER_ENDING_TIMER = false;
        private const SynchronizationMode DEFAULT_SYNCHRONIZATION_MODE = SynchronizationMode.None;
        private const string SET_PROP_WITH_PARENT_ANIMATOR_ERROR_MESSAGE = "Property cannot be set while ParentAnimator is set to anything other than null.";

        private System.Windows.Forms.Timer _timer;
        private System.ComponentModel.Container components = null;

        private double _stepSize = DEFAULT_STEP_SIZE;
        private double _currentStep;
        private LoopMode _loopMode = DEFAULT_LOOP_ANIMATION;
        private bool _neverEndingTimer = DEFAULT_NEVER_ENDING_TIMER;

        private SynchronizationMode _syncMode = DEFAULT_SYNCHRONIZATION_MODE;

        private AnimatorBase _parentAnimator;
        private AnimatorBase _triggerAnimator;

        private bool _isInitializing = false;
        private ArrayList _childAnimators = new ArrayList();
        private bool _settingCurrentValue = false;

        #endregion

        #region Events

        /// <summary>
        /// Event which gets fired when animation has been started with <see cref="Start()"/>.
        /// </summary>
        public event EventHandler AnimationStarted;

        /// <summary>
        /// Event which gets fired when animation has been started with <see cref="Stop()"/>.
        /// </summary>
        public event EventHandler AnimationStopped;

        /// <summary>
        /// Event which gets fired when animation has been started with <see cref="Continue()"/>.
        /// </summary>
        public event EventHandler AnimationContinued;

        /// <summary>
        /// Event which gets fired when animation has finished running.
        /// </summary>
        public event EventHandler AnimationFinished;

        /// <summary>
        /// Event which gets fired when <see cref="StepSize"/> has changed.
        /// </summary>
        public event EventHandler StepSizeChanged;

        /// <summary>
        /// Event which gets fired when <see cref="Intervall"/> has changed.
        /// </summary>
        public event EventHandler IntervallChanged;

        /// <summary>
        /// Event which gets fired when <see cref="CurrentStep"/> has changed.
        /// </summary>
        public event EventHandler CurrentStepChanged;

        /// <summary>
        /// Event which gets fired when <see cref="LoopMode"/> has changed.
        /// </summary>
        public event EventHandler LoopAnimationChanged;

        /// <summary>
        /// Event which gets fired when <see cref="StartValue"/> has changed.
        /// </summary>
        public event EventHandler StartValueChanged;

        /// <summary>
        /// Event which gets fired when <see cref="EndValue"/> has changed.
        /// </summary>
        public event EventHandler EndValueChanged;

        /// <summary>
        /// Event which gets fired when <see cref="SynchronizationMode"/> has changed.
        /// </summary>
        public event EventHandler SynchronizationModeChanged;

        #endregion

        #region Constructors

        /// <summary>
        /// Base constructor.
        /// </summary>
        /// <param name="container">Container the new instance should be added to.</param>
        public AnimatorBase(System.ComponentModel.IContainer container)
        {
            container.Add(this);
            InitializeComponent();
            Initialize();
        }

        /// <summary>
        /// Base constructor.
        /// </summary>
        public AnimatorBase()
        {
            InitializeComponent();
            Initialize();
        }

        private void Initialize()
        {
            _timer.Interval = DEFAULT_INTERVALL;
        }

        #endregion

        #region Designer generated code
        /// <summary>
        /// Erforderliche Methode für die Designerunterstützung. 
        /// Der Inhalt der Methode darf nicht mit dem Code-Editor geändert werden.
        /// </summary>
        private void InitializeComponent()
        {
            this._timer = new System.Windows.Forms.Timer();
            // 
            // _timer
            // 
            this._timer.Tick += new EventHandler(this.OnTimerElapsed);

        }
        #endregion

        #region Overridden from Component

        /// <summary>
        /// Frees used resources.
        /// </summary>
        /// <param name="disposing"></param>
        protected override void Dispose(bool disposing)
        {
            this.ParentAnimator = null;
            _childAnimators.Clear();
            this.TriggerAnimator = null;

            if (disposing)
            {
                if (components != null)
                {
                    components.Dispose();
                }
            }
            base.Dispose(disposing);
        }

        #endregion

        #region Public interface

        #region Properties

        #region Value getters and setters (abstract)

        /// <summary>
        /// Gets or sets the starting value for the animation.
        /// </summary>
        [Browsable(false)]
        [DesignerSerializationVisibility(DesignerSerializationVisibility.Hidden)]
        public abstract object StartValue { get; set; }

        /// <summary>
        /// Gets or sets the ending value for the animation.
        /// </summary>
        [Browsable(false)]
        [DesignerSerializationVisibility(DesignerSerializationVisibility.Hidden)]
        public abstract object EndValue { get; set; }

        /// <summary>
        /// Gets or sets the currently shown value.
        /// </summary>
        [Browsable(false)]
        [DesignerSerializationVisibility(DesignerSerializationVisibility.Hidden)]
        public object CurrentValue
        {
            get { return CurrentValueInternal; }
            set
            {
                if (_settingCurrentValue)
                    throw new InvalidOperationException();

                try
                {
                    _settingCurrentValue = true;
                    CurrentValueInternal = value;
                }
                finally
                {
                    _settingCurrentValue = false;
                }
            }
        }

        #endregion

        /// <summary>
        /// Gets or sets the <see cref="AnimatorBase"/> which should trigger the animation
        /// of this instance when it has finished animating.
        /// </summary>
        [Browsable(true), DefaultValue(null), Category("Behavior")]
        [Description("Gets or sets the AnimatorBase which should trigger the animation of this instance when it has finished animating.")]
        public AnimatorBase TriggerAnimator
        {
            get { return _triggerAnimator; }
            set
            {
                if (_triggerAnimator == value)
                    return;

                if (_triggerAnimator == this)
                    throw new InvalidOperationException("Cannot set itself as TriggerAnimator.");

                if (_triggerAnimator != null)
                    _triggerAnimator.AnimationFinished -= new EventHandler(OnTriggerAnimatorAnimationFinished);

                _triggerAnimator = value;

                if (_triggerAnimator != null)
                    _triggerAnimator.AnimationFinished += new EventHandler(OnTriggerAnimatorAnimationFinished);
            }
        }

        /// <summary>
        /// Sets the <see cref="AnimatorBase"/> which acts as a parent of this instance.
        /// Thus the settings <see cref="Intervall"/>, <see cref="StepSize"/>, <see cref="LoopMode"/>
        /// and <see cref="SynchronizationMode"/> of this instance will be set accordingly to and synchronized 
        /// with the settings of this parent.
        /// </summary>
        [Browsable(true), DefaultValue(null), Category("Behavior")]
        [RefreshProperties(RefreshProperties.Repaint)]
        public AnimatorBase ParentAnimator
        {
            get { return _parentAnimator; }
            set
            {
                if (_parentAnimator == value)
                    return;

                if (_parentAnimator == this)
                    throw new InvalidOperationException("Cannot set itself as ParentAnimator.");

                if (_parentAnimator != null)
                    _parentAnimator.RemoveChildAnimator(this);

                _parentAnimator = value;

                if (_parentAnimator != null)
                    _parentAnimator.AddChildAnimator(this);
            }
        }

        /// <summary>
        /// Gets or sets the mode of design time synchronization.
        /// </summary>
        [Description("Gets or sets the mode of design time synchronization.")]
        [Browsable(true), Category("Design"), RefreshProperties(RefreshProperties.Repaint)]
        public SynchronizationMode SynchronizationMode
        {
            get { return _syncMode; }
            set { SetSynchronizationMode(value, true); }
        }

        /// <summary>
        /// Gets or sets the intervall (in milliseconds) between updates to the animation.
        /// </summary>
        [Description("Gets or sets the intervall (in milliseconds) between updates to the animation.")]
        [Browsable(true), Category("Behavior"), DefaultValue(DEFAULT_INTERVALL)]
        public int Intervall
        {
            get { return _timer.Interval; }
            set { SetIntervall(value, true); }
        }

        /// <summary>
        /// Gets or sets the size of each step (in %) when updating the animation.
        /// </summary>
        [Description("Gets or sets the size of each step (in %) when updating the animation.")]
        [Browsable(true), Category("Behavior"), DefaultValue(DEFAULT_STEP_SIZE)]
        public double StepSize
        {
            get { return _stepSize; }
            set { SetStepSize(value, true); }
        }

        /// <summary>
        /// Gets or sets whether the animation should loop between <see cref="StartValue"/>
        /// and <see cref="EndValue"/> until <see cref="Stop()"/> is called.
        /// </summary>
        [Description("Gets or sets whether the animation should loop between StartValue and EndValue until Stop() is called.")]
        [Browsable(true), Category("Behavior"), DefaultValue(DEFAULT_LOOP_ANIMATION)]
        public LoopMode LoopMode
        {
            get { return _loopMode; }
            set { SetLoopMode(value, true); }
        }

        /// <summary>
        /// Gets or sets the current step (in %) of the animation.
        /// </summary>
        [Browsable(false)]
        [DesignerSerializationVisibility(DesignerSerializationVisibility.Hidden)]
        public double CurrentStep
        {
            get { return _currentStep; }
            set
            {
                if (_currentStep == value)
                    return;

                _currentStep = value;

                if (_currentStep > 100)
                    _currentStep = 100;
                else if (_currentStep < 0)
                    _currentStep = 0;

                CurrentValue = GetValueForStep(_currentStep);

                foreach (AnimatorBase childAnimator in _childAnimators)
                    childAnimator.CurrentStep = _currentStep;

                OnCurrentStepChanged(EventArgs.Empty);
            }
        }

        /// <summary>
        /// Gets whether an animation is currently running.
        /// </summary>
        [Browsable(false)]
        [DesignerSerializationVisibility(DesignerSerializationVisibility.Hidden)]
        public bool IsRunning
        {
            get { return _parentAnimator == null ? _timer.Enabled : _parentAnimator.IsRunning; }
        }

        /// <summary>
        /// Gets or sets whether the internal timer should always continue running
        /// even if the animation has reached its end. This can be useful when
        /// the animation is frequently continued after is has finished because
        /// starting and stopping the timer has a great influence on the performance.
        /// </summary>
        [Description("Gets or sets whether the internal timer should always continue running even if the animation has reached its end")]
        [Browsable(true), Category("Behavior"), DefaultValue(DEFAULT_NEVER_ENDING_TIMER)]
        public bool NeverEndingTimer
        {
            get { return _neverEndingTimer; }
            set { _neverEndingTimer = value; }
        }

        #endregion

        #region Animation controlling

        /// <summary>
        /// Continues the animation.
        /// If the animation finished normally (without call to <see cref="Stop()"/>)
        /// than a call to this function will not have any effect.
        /// </summary>
        public void Continue()
        {
            _timer.Start();

            OnAnimationContinued(EventArgs.Empty);
        }

        /// <summary>
        /// Sets the <see cref="AnimatorBase.StartValue"/> to the
        /// <see cref="AnimatorBase.CurrentValue"/>, sets the given
        /// value to <see cref="EndValue"/> and starts the animation.
        /// </summary>
        /// <param name="endValue">new end value for the animation.</param>
        public void Start(object endValue)
        {
            if (_childAnimators.Count > 0)
                throw new InvalidOperationException("Function cannot be called when ChildAnimators are set.");

            EndValue = endValue;
            Start(true);
        }

        /// <summary>
        /// Sets <see cref="AnimatorBase.StartValue"/> to
        /// <see cref="AnimatorBase.CurrentValue"/> and starts the animation.
        /// </summary>
        public void Start()
        {
            Start(true);
        }

        /// <summary>
        /// Optionally sets <see cref="AnimatorBase.StartValue"/> to
        /// <see cref="AnimatorBase.CurrentValue"/> and starts the animation.
        /// </summary>
        /// <param name="setStartValuesToCurrentValues">Indicates whether the start value
        /// should be changed prior to starting the animation.</param>
        public void Start(bool setStartValuesToCurrentValues)
        {
            if (setStartValuesToCurrentValues)
                SetStartValuesToCurrentValue();

            this.CurrentStep = 0;
            if (!_timer.Enabled)
                _timer.Start();

            OnAnimationStarted(EventArgs.Empty);
        }

        /// <summary>
        /// Sets <see cref="CurrentValue"/> to the value of <see cref="StartValue"/>
        /// for this instance and all registered childs.
        /// </summary>
        public void SetCurrentValuesToStartValues()
        {
            CurrentValue = StartValue;
            foreach (AnimatorBase childAnimator in _childAnimators)
                childAnimator.SetCurrentValuesToStartValues();
        }

        /// <summary>
        /// Sets <see cref="StartValue"/> to the value of <see cref="CurrentValue"/>
        /// for this instance and all registered childs.
        /// </summary>
        public void SetStartValuesToCurrentValue()
        {
            StartValue = CurrentValue;
            foreach (AnimatorBase childAnimator in _childAnimators)
                childAnimator.SetStartValuesToCurrentValue();
        }

        /// <summary>
        /// Sets <see cref="StartValue"/> to the value of <see cref="EndValue"/>
        /// and vice versa for this instance and all registered childs.
        /// </summary>
        public void SwitchStartEndValues()
        {
            object temp = StartValue;
            StartValue = EndValue;
            EndValue = temp;
            foreach (AnimatorBase childAnimator in _childAnimators)
                childAnimator.SwitchStartEndValues();
        }

        /// <summary>
        /// Stops the animation.
        /// If no animation is running nothing happens.
        /// </summary>
        public void Stop()
        {
            if (_timer.Enabled)
            {
                _timer.Stop();
                OnAnimationStopped(EventArgs.Empty);
            }
        }

        #endregion

        #endregion

        #region Protected

        #region Design time synchronization

        /// <summary>
        /// Changes <see cref="StartValue"/>, <see cref="EndValue"/>
        /// or <see cref="CurrentValue"/> accordingly to the
        /// current <see cref="SynchronizationMode"/>.
        /// </summary>
        protected void SynchronizeToSource()
        {
            if (!Program.IsDesign)
                return;

            switch (_syncMode)
            {
                case SynchronizationMode.Start:
                    CurrentValue = StartValue;
                    break;
                case SynchronizationMode.End:
                    CurrentValue = EndValue;
                    break;
            }
        }

        /// <summary>
        /// Sets <see cref="StartValue"/> or <see cref="EndValue"/>
        /// to <see cref="CurrentValue"/> depending on the currently
        /// set <see cref="SynchronizationMode"/>.
        /// </summary>
        protected void SynchronizeFromSource()
        {
            if (!Program.IsDesign)
                return;

            switch (_syncMode)
            {
                case SynchronizationMode.Start:
                    StartValue = CurrentValue;
                    break;
                case SynchronizationMode.End:
                    EndValue = CurrentValue;
                    break;
            }
        }

        /// <summary>
        /// Sets <see cref="StartValue"/> and <see cref="EndValue"/>
        /// to <see cref="CurrentValue"/>.
        /// </summary>
        protected void ResetValues()
        {
            if (_isInitializing)
                return;

            StartValue = CurrentValue;
            EndValue = CurrentValue;
        }

        #endregion

        #region Event raising

        /// <summary>
        /// Raises the <see cref="AnimationStarted"/> event.
        /// </summary>
        /// <param name="eventArgs">Event data.</param>
        protected virtual void OnAnimationStarted(EventArgs eventArgs)
        {
            if (AnimationStarted != null)
                AnimationStarted(this, eventArgs);
        }

        /// <summary>
        /// Raises the <see cref="AnimationContinued"/> event.
        /// </summary>
        /// <param name="eventArgs">Event data.</param>
        protected virtual void OnAnimationContinued(EventArgs eventArgs)
        {
            if (AnimationContinued != null)
                AnimationContinued(this, eventArgs);
        }

        /// <summary>
        /// Raises the <see cref="AnimationStopped"/> event.
        /// </summary>
        /// <param name="eventArgs">Event data.</param>
        protected virtual void OnAnimationStopped(EventArgs eventArgs)
        {
            if (AnimationStopped != null)
                AnimationStopped(this, eventArgs);
        }

        /// <summary>
        /// Raises the <see cref="AnimationFinished"/> event.
        /// </summary>
        /// <param name="eventArgs">Event data.</param>
        protected virtual void OnAnimationFinished(EventArgs eventArgs)
        {
            if (AnimationFinished != null)
                AnimationFinished(this, eventArgs);
        }

        /// <summary>
        /// Raises the <see cref="LoopAnimationChanged"/> event.
        /// </summary>
        /// <param name="eventArgs">Event data.</param>
        protected virtual void OnLoopAnimationChanged(EventArgs eventArgs)
        {
            if (LoopAnimationChanged != null)
                LoopAnimationChanged(this, eventArgs);
        }

        /// <summary>
        /// Raises the <see cref="StepSizeChanged"/> event.
        /// </summary>
        /// <param name="eventArgs">Event data.</param>
        protected virtual void OnStepSizeChanged(EventArgs eventArgs)
        {
            if (StepSizeChanged != null)
                StepSizeChanged(this, eventArgs);
        }

        /// <summary>
        /// Raises the <see cref="IntervallChanged"/> event.
        /// </summary>
        /// <param name="eventArgs">Event data.</param>
        protected virtual void OnIntervallChanged(EventArgs eventArgs)
        {
            if (IntervallChanged != null)
                IntervallChanged(this, eventArgs);
        }

        /// <summary>
        /// Raises the <see cref="SynchronizationModeChanged"/> event.
        /// </summary>
        /// <param name="eventArgs">Event data.</param>
        protected void OnSynchronizationModeChanged(EventArgs eventArgs)
        {
            if (SynchronizationModeChanged != null)
                SynchronizationModeChanged(this, eventArgs);
        }

        /// <summary>
        /// Raises the <see cref="CurrentStepChanged"/> event.
        /// </summary>
        /// <param name="eventArgs">Event data.</param>
        protected virtual void OnCurrentStepChanged(EventArgs eventArgs)
        {
            if (CurrentStepChanged != null)
                CurrentStepChanged(this, eventArgs);
        }

        /// <summary>
        /// Raises the <see cref="StartValueChanged"/> event.
        /// </summary>
        /// <param name="eventArgs">Event data.</param>
        protected virtual void OnStartValueChanged(EventArgs eventArgs)
        {
            if (_syncMode == SynchronizationMode.Start)
                CurrentValue = StartValue;

            if (StartValueChanged != null)
                StartValueChanged(this, eventArgs);
        }

        /// <summary>
        /// Raises the <see cref="EndValueChanged"/> event.
        /// </summary>
        /// <param name="eventArgs">Event data.</param>
        protected virtual void OnEndValueChanged(EventArgs eventArgs)
        {
            if (_syncMode == SynchronizationMode.End)
                CurrentValue = EndValue;

            if (EndValueChanged != null)
                EndValueChanged(this, eventArgs);
        }

        #endregion

        /// <summary>
        /// Gets or sets the currently shown value (for internal usage).
        /// </summary>
        protected abstract object CurrentValueInternal { get; set; }

        /// <summary>
        /// Gets whether the control is in the process of setting 
        /// the <see cref="CurrentValue"/> internally. 
        /// </summary>
        protected bool SettingCurrentValue
        {
            get { return _settingCurrentValue; }
        }

        /// <summary>
        /// Calculates an interpolated value between <see cref="StartValue"/> and
        /// <see cref="EndValue"/> for a given step in %.
        /// Giving 0 will return the <see cref="StartValue"/>.
        /// Giving 100 will return the <see cref="EndValue"/>.
        /// </summary>
        /// <param name="step">Animation step in %</param>
        /// <returns>Interpolated value for the given step.</returns>
        protected abstract object GetValueForStep(double step);

        /// <summary>
        /// Adds an <see cref="AnimatorBase"/> which acts as a child of this instance.
        /// Thus its <see cref="Intervall"/>, <see cref="StepSize"/>, <see cref="LoopMode"/>
        /// and <see cref="SynchronizationMode"/> will be set accordingly to and synchronized 
        /// with the settings of this instance.
        /// </summary>
        /// <param name="animator">Child to add.</param>
        protected void AddChildAnimator(AnimatorBase animator)
        {
            if (animator == null)
                throw new ArgumentNullException("animator");

            if (!_childAnimators.Contains(animator))
                _childAnimators.Add(animator);

            animator.SetIntervall(this.Intervall, false);
            animator.SetStepSize(this.StepSize, false);
            animator.SetLoopMode(this.LoopMode, false);
            animator.SetSynchronizationMode(this.SynchronizationMode, false);
        }

        /// <summary>
        /// Removes an <see cref="AnimatorBase"/> which acts as a child of this instance.
        /// </summary>
        /// <param name="animator">Child to be removed.</param>
        protected void RemoveChildAnimator(AnimatorBase animator)
        {
            if (animator == null)
                throw new ArgumentNullException("animator");

            if (_childAnimators.Contains(animator))
                _childAnimators.Remove(animator);
        }

        /// <summary>
        /// Gets whether the control is in its initialization process.
        /// </summary>
        protected bool IsInitializing
        {
            get { return _isInitializing; }
        }

        /// <summary>
        /// Indicates the designer whether <see cref="SynchronizationMode"/> needs
        /// to be serialized.
        /// </summary>
        protected virtual bool ShouldSerializeSynchronizationMode()
        {
            return false;
        }

        /// <summary>
        /// Should be called by inheriting classes whenever the current
        /// value of whatever they animated changes. This ensures optimal
        /// design time support.
        /// </summary>
        /// <param name="sender">Sender of the notification.</param>
        /// <param name="e">Event arguments.</param>
        protected virtual void OnCurrentValueChanged(object sender, EventArgs e)
        {
            if (SettingCurrentValue)
                return;

            SynchronizeFromSource();
        }

        #endregion

        #region Privates

        #region Internal setters

        private void SetSynchronizationMode(SynchronizationMode synchronizationMode, bool checkParentAnimator)
        {
            if (_syncMode == synchronizationMode)
                return;

            if (synchronizationMode == SynchronizationMode.ResetToCurrent)
            {
                if (!Program.IsDesign)
                    return;

                ResetValues();

                return;
            }

            if (_parentAnimator != null && checkParentAnimator && !_isInitializing)
                throw new InvalidOperationException(SET_PROP_WITH_PARENT_ANIMATOR_ERROR_MESSAGE);

            _syncMode = synchronizationMode;
            SynchronizeToSource();

            foreach (AnimatorBase childAnimator in _childAnimators)
                childAnimator.SetSynchronizationMode(synchronizationMode, false);

            OnSynchronizationModeChanged(EventArgs.Empty);
        }

        private void SetIntervall(int intervall, bool checkParentAnimator)
        {
            if (_timer.Interval == intervall)
                return;

            if (_parentAnimator != null && checkParentAnimator && !_isInitializing)
                throw new InvalidOperationException(SET_PROP_WITH_PARENT_ANIMATOR_ERROR_MESSAGE);

            _timer.Interval = intervall;

            foreach (AnimatorBase childAnimator in _childAnimators)
                childAnimator.SetIntervall(intervall, false);

            OnIntervallChanged(EventArgs.Empty);
        }

        private void SetStepSize(double stepSize, bool checkParentAnimator)
        {
            if (_stepSize == stepSize)
                return;

            if (_parentAnimator != null && checkParentAnimator && !_isInitializing)
                throw new InvalidOperationException(SET_PROP_WITH_PARENT_ANIMATOR_ERROR_MESSAGE);

            _stepSize = stepSize;

            foreach (AnimatorBase childAnimator in _childAnimators)
                childAnimator.SetStepSize(stepSize, false);

            OnStepSizeChanged(EventArgs.Empty);
        }

        private void SetLoopMode(LoopMode loopAnimation, bool checkParentAnimator)
        {
            if (_loopMode == loopAnimation)
                return;

            if (_parentAnimator != null && checkParentAnimator && !_isInitializing)
                throw new InvalidOperationException(SET_PROP_WITH_PARENT_ANIMATOR_ERROR_MESSAGE);

            _loopMode = loopAnimation;

            foreach (AnimatorBase childAnimator in _childAnimators)
                childAnimator.SetLoopMode(loopAnimation, false);

            OnLoopAnimationChanged(EventArgs.Empty);
        }

        #endregion

        #region Event handler

        private void OnTimerElapsed(object sender, EventArgs e)
        {
            this.CurrentStep += _stepSize;

            if (this.CurrentStep >= 100)
            {
                bool timerEnabled = _timer.Enabled;
                if (_timer.Enabled && !_neverEndingTimer && _loopMode == LoopMode.None)
                    _timer.Stop();

                OnAnimationFinished(EventArgs.Empty);

                if (timerEnabled)
                {
                    if (_loopMode == LoopMode.Repeat)
                    {
                        this.CurrentStep -= 100;
                    }
                    else if (_loopMode == LoopMode.Bidirectional)
                    {
                        SwitchStartEndValues();
                        this.Start();
                    }
                }
            }
        }

        private void OnTriggerAnimatorAnimationFinished(object sender, EventArgs e)
        {
            this.Start();
        }

        #endregion

        #endregion

        #region Static helpers for value interpolation

        /// <summary>
        /// Interpolates two <see cref="Color"/> instances.
        /// </summary>
        /// <param name="color1">First color.</param>
        /// <param name="color2">Second color.</param>
        /// <param name="percent">Value ranging from 0 to 100 indicating how big the part
        /// of the two colors is in the result.</param>
        /// <returns>If percent is 0 then color1. If percent is 100 then color2. For 
        /// everything else an interpolated color is returned.</returns>
        public static Color InterpolateColors(Color color1, Color color2, double percent)
        {
            return Color.FromArgb(
                InterpolateIntegerValues(color1.A, color2.A, percent),
                InterpolateIntegerValues(color1.R, color2.R, percent),
                InterpolateIntegerValues(color1.G, color2.G, percent),
                InterpolateIntegerValues(color1.B, color2.B, percent));
        }

        /// <summary>
        /// Interpolates two <see cref="Rectangle"/> instances.
        /// </summary>
        /// <param name="rectangle1">First rectangle.</param>
        /// <param name="rectangle2">Second rectangle.</param>
        /// <param name="percent">Value ranging from 0 to 100 indicating how big the part
        /// of the two rectangles is in the result.</param>
        /// <returns>If percent is 0 then rectangle1. If percent is 100 then rectangle2. For 
        /// everything else an interpolated rectangle is returned.</returns>
        public static Rectangle InterpolateRectangles(Rectangle rectangle1, Rectangle rectangle2, double percent)
        {
            return new Rectangle(InterpolatePoints(rectangle1.Location, rectangle2.Location, percent),
                InterpolateSizes(rectangle1.Size, rectangle2.Size, percent));
        }

        /// <summary>
        /// Interpolates two <see cref="Point"/> instances.
        /// </summary>
        /// <param name="point1">First point.</param>
        /// <param name="point2">Second point.</param>
        /// <param name="percent">Value ranging from 0 to 100 indicating how big the part
        /// of the two points is in the result.</param>
        /// <returns>If percent is 0 then point1. If percent is 100 then point2. For 
        /// everything else an interpolated point is returned.</returns>
        public static Point InterpolatePoints(Point point1, Point point2, double percent)
        {
            return new Point(InterpolateIntegerValues(point1.X, point2.X, percent),
                InterpolateIntegerValues(point1.Y, point2.Y, percent));
        }

        /// <summary>
        /// Interpolates two <see cref="Size"/> instances.
        /// </summary>
        /// <param name="size1">First size.</param>
        /// <param name="size2">Second size.</param>
        /// <param name="percent">Value ranging from 0 to 100 indicating how big the part
        /// of the two sizes is in the result.</param>
        /// <returns>If percent is 0 then size1. If percent is 100 then size2. For 
        /// everything else an interpolated size is returned.</returns>
        public static Size InterpolateSizes(Size size1, Size size2, double percent)
        {
            return new Size(InterpolateIntegerValues(size1.Width, size2.Width, percent),
                InterpolateIntegerValues(size1.Height, size2.Height, percent));
        }

        /// <summary>
        /// Interpolates two <see cref="double"/> values.
        /// </summary>
        /// <param name="value1">First value.</param>
        /// <param name="value2">Second value.</param>
        /// <param name="percent">Value ranging from 0 to 100 indicating how big the part
        /// of the two values is in the result.</param>
        /// <returns>If percent is 0 then value1. If percent is 100 then value2. For 
        /// everything else an interpolated value is returned.</returns>
        public static double InterpolateDoubleValues(double value1, double value2, double percent)
        {
            if (percent < 0 || percent > 100)
                throw new ArgumentException("Value must be between 0 and 100.", "percent");

            return percent * (value2 - value1) / 100 + value1;
        }

        /// <summary>
        /// Interpolates two <see cref="int"/> values.
        /// </summary>
        /// <param name="value1">First value.</param>
        /// <param name="value2">Second value.</param>
        /// <param name="percent">Value ranging from 0 to 100 indicating how big the part
        /// of the two values is in the result.</param>
        /// <returns>If percent is 0 then value1. If percent is 100 then value2. For 
        /// everything else an interpolated value is returned.</returns>
        public static int InterpolateIntegerValues(int value1, int value2, double percent)
        {
            if (percent < 0 || percent > 100)
                throw new ArgumentException("Value must be between 0 and 100.", "percent");

            return Convert.ToInt32(percent * (value2 - value1) / 100 + value1);
        }

        #endregion

        #region ISupportInitialize Member

        /// <summary>
        /// Signals the object that initialization is starting.
        /// </summary>
        public void BeginInit()
        {
            _isInitializing = true;
        }

        /// <summary>
        /// Signals the object that initialization is complete.
        /// </summary>
        public void EndInit()
        {
            _isInitializing = false;
        }

        #endregion
    }
}
