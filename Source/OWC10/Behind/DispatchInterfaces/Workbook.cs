using System;
using NetRuntimeSystem = System;
using System.ComponentModel;
using NetOffice.Attributes;

namespace NetOffice.OWC10Api.Behind
{
    /// <summary>
    /// Workbook
    /// </summary>
    [SyntaxBypass]
    public class Workbook_ : COMObject, NetOffice.OWC10Api.Workbook_
    {
        #region Ctor

        /// <summary>
        /// Stub Ctor, not intended to use
        /// </summary>
        [EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
        public Workbook_() : base()
        {
        }

        #endregion

        #region Properties
        /// <summary>
        /// SupportByVersion OWC10 1
        /// Get/Set
        /// </summary>
        /// <param name="index">optional object index</param>
        [SupportByVersion("OWC10", 1)]
        [EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
        public virtual object get_Colors(object index)
        {
            return InvokerService.InvokeInternal.ExecuteVariantPropertyGet(this, "Colors", index);
        }

        /// <summary>
        /// SupportByVersion OWC10 1
        /// Get/Set
        /// </summary>
        /// <param name="index">optional object index</param>
        /// <param name="value">optional object value</param>
        [SupportByVersion("OWC10", 1)]
        [EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
        public virtual void set_Colors(object index, object value)
        {
            InvokerService.InvokeInternal.ExecutePropertySet(this, "Colors", index, value);
        }

        /// <summary>
        /// SupportByVersion OWC10 1
        /// Alias for get_Colors
        /// </summary>
        /// <param name="index">optional object index</param>
        [SupportByVersion("OWC10", 1), Redirect("get_Colors")]
        public virtual object Colors(object index)
        {
            return get_Colors(index);
        }

        #endregion
    }

    /// <summary>
    /// DispatchInterface Workbook 
    /// SupportByVersion OWC10, 1
    /// </summary>
    [SupportByVersion("OWC10", 1)]
    [EntityType(EntityType.IsDispatchInterface)]
    public class Workbook : Workbook_, NetOffice.OWC10Api.Workbook
    {
        #pragma warning disable

        #region Type Information

        /// <summary>
        /// Contract Type
        /// </summary>
        [EditorBrowsable(EditorBrowsableState.Advanced), Browsable(false), Category("NetOffice"), CoreOverridden]
        public override Type ContractType
        {
            get
            {
                if(null == _contractType)
                    _contractType = typeof(NetOffice.OWC10Api.Workbook);
                return _contractType;
            }
        }
        private static Type _contractType;


        /// <summary>
        /// Instance Type
        /// </summary>
        [EditorBrowsable(EditorBrowsableState.Advanced), Browsable(false), Category("NetOffice"), CoreOverridden]
        public override Type InstanceType
        {
            get
            {
                return LateBindingApiWrapperType;
            }
        }

        private static Type _type;

        [EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
        public static Type LateBindingApiWrapperType
        {
            get
            {
                if (null == _type)
                    _type = typeof(Workbook);
                return _type;
            }
        }

        #endregion

        #region Ctor

        /// <summary>
        /// Stub Ctor, not intended to use
        /// </summary>
        public Workbook() : base()
        {
        }

        #endregion

        #region Properties

        /// <summary>
        /// SupportByVersion OWC10 1
        /// Get
        /// </summary>
        [SupportByVersion("OWC10", 1)]
        public virtual NetOffice.OWC10Api.Worksheet ActiveSheet
        {
            get
            {
                return InvokerService.InvokeInternal.ExecuteKnownReferencePropertyGet<NetOffice.OWC10Api.Worksheet>(this, "ActiveSheet", typeof(NetOffice.OWC10Api.Worksheet));
            }
        }

        /// <summary>
        /// SupportByVersion OWC10 1
        /// Get
        /// </summary>
        [SupportByVersion("OWC10", 1)]
        [BaseResult]
        public virtual NetOffice.OWC10Api.ISpreadsheet Application
        {
            get
            {
                return InvokerService.InvokeInternal.ExecuteBaseReferencePropertyGet<NetOffice.OWC10Api.ISpreadsheet>(this, "Application");
            }
        }

        /// <summary>
        /// SupportByVersion OWC10 1
        /// Get
        /// </summary>
        [SupportByVersion("OWC10", 1)]
        public virtual Int32 CalculationVersion
        {
            get
            {
                return InvokerService.InvokeInternal.ExecuteInt32PropertyGet(this, "CalculationVersion");
            }
        }

        /// <summary>
        /// SupportByVersion OWC10 1
        /// Get/Set
        /// </summary>
        [SupportByVersion("OWC10", 1)]
        public virtual object Colors
        {
            get
            {
                return InvokerService.InvokeInternal.ExecuteVariantPropertyGet(this, "Colors");
            }
            set
            {
                InvokerService.InvokeInternal.ExecuteVariantPropertySet(this, "Colors", value);
            }
        }

        /// <summary>
        /// SupportByVersion OWC10 1
        /// Get
        /// </summary>
        [SupportByVersion("OWC10", 1)]
        public virtual string Name
        {
            get
            {
                return InvokerService.InvokeInternal.ExecuteStringPropertyGet(this, "Name");
            }
        }

        /// <summary>
        /// SupportByVersion OWC10 1
        /// Get
        /// </summary>
        [SupportByVersion("OWC10", 1)]
        public virtual NetOffice.OWC10Api.Names Names
        {
            get
            {
                return InvokerService.InvokeInternal.ExecuteKnownReferencePropertyGet<NetOffice.OWC10Api.Names>(this, "Names", typeof(NetOffice.OWC10Api.Names));
            }
        }

        /// <summary>
        /// SupportByVersion OWC10 1
        /// Get
        /// </summary>
        [SupportByVersion("OWC10", 1)]
        [BaseResult]
        public virtual NetOffice.OWC10Api.ISpreadsheet Parent
        {
            get
            {
                return InvokerService.InvokeInternal.ExecuteBaseReferencePropertyGet<NetOffice.OWC10Api.ISpreadsheet>(this, "Parent");
            }
        }

        /// <summary>
        /// SupportByVersion OWC10 1
        /// Get
        /// </summary>
        [SupportByVersion("OWC10", 1)]
        public virtual bool ProtectStructure
        {
            get
            {
                return InvokerService.InvokeInternal.ExecuteBoolPropertyGet(this, "ProtectStructure");
            }
        }

        /// <summary>
        /// SupportByVersion OWC10 1
        /// Get
        /// </summary>
        [SupportByVersion("OWC10", 1)]
        public virtual NetOffice.OWC10Api.Sheets Sheets
        {
            get
            {
                return InvokerService.InvokeInternal.ExecuteKnownReferencePropertyGet<NetOffice.OWC10Api.Sheets>(this, "Sheets", typeof(NetOffice.OWC10Api.Sheets));
            }
        }

        /// <summary>
        /// SupportByVersion OWC10 1
        /// Get
        /// </summary>
        [SupportByVersion("OWC10", 1)]
        public virtual NetOffice.OWC10Api.Windows Windows
        {
            get
            {
                return InvokerService.InvokeInternal.ExecuteKnownReferencePropertyGet<NetOffice.OWC10Api.Windows>(this, "Windows", typeof(NetOffice.OWC10Api.Windows));
            }
        }

        /// <summary>
        /// SupportByVersion OWC10 1
        /// Get
        /// </summary>
        [SupportByVersion("OWC10", 1)]
        public virtual NetOffice.OWC10Api.Worksheets Worksheets
        {
            get
            {
                return InvokerService.InvokeInternal.ExecuteKnownReferencePropertyGet<NetOffice.OWC10Api.Worksheets>(this, "Worksheets", typeof(NetOffice.OWC10Api.Worksheets));
            }
        }

        #endregion

        #region Methods

        /// <summary>
        /// SupportByVersion OWC10 1
        /// </summary>
        /// <param name="password">optional object password</param>
        /// <param name="structure">optional object structure</param>
        /// <param name="windows">optional object windows</param>
        [SupportByVersion("OWC10", 1)]
        public virtual void Protect(object password, object structure, object windows)
        {
            InvokerService.InvokeInternal.ExecuteMethod(this, "Protect", password, structure, windows);
        }

        /// <summary>
        /// SupportByVersion OWC10 1
        /// </summary>
        [CustomMethod]
        [SupportByVersion("OWC10", 1)]
        public virtual void Protect()
        {
            InvokerService.InvokeInternal.ExecuteMethod(this, "Protect");
        }

        /// <summary>
        /// SupportByVersion OWC10 1
        /// </summary>
        /// <param name="password">optional object password</param>
        [CustomMethod]
        [SupportByVersion("OWC10", 1)]
        public virtual void Protect(object password)
        {
            InvokerService.InvokeInternal.ExecuteMethod(this, "Protect", password);
        }

        /// <summary>
        /// SupportByVersion OWC10 1
        /// </summary>
        /// <param name="password">optional object password</param>
        /// <param name="structure">optional object structure</param>
        [CustomMethod]
        [SupportByVersion("OWC10", 1)]
        public virtual void Protect(object password, object structure)
        {
            InvokerService.InvokeInternal.ExecuteMethod(this, "Protect", password, structure);
        }

        /// <summary>
        /// SupportByVersion OWC10 1
        /// </summary>
        [SupportByVersion("OWC10", 1)]
        public virtual void ResetColors()
        {
            InvokerService.InvokeInternal.ExecuteMethod(this, "ResetColors");
        }

        /// <summary>
        /// SupportByVersion OWC10 1
        /// </summary>
        /// <param name="password">optional object password</param>
        [SupportByVersion("OWC10", 1)]
        public virtual void Unprotect(object password)
        {
            InvokerService.InvokeInternal.ExecuteMethod(this, "Unprotect", password);
        }

        /// <summary>
        /// SupportByVersion OWC10 1
        /// </summary>
        [CustomMethod]
        [SupportByVersion("OWC10", 1)]
        public virtual void Unprotect()
        {
            InvokerService.InvokeInternal.ExecuteMethod(this, "Unprotect");
        }

        #endregion

        #pragma warning restore
    }
}

