using System.Reflection;
using System;
using NetRuntimeSystem = System;
using System.ComponentModel;
using NetOffice.Attributes;
using NetOffice.OfficeApi;

namespace NetOffice.OfficeApi.Behind
{
    /// <summary>
    /// IAccessible
    /// </summary>
    [SyntaxBypass]
    public class IAccessible_ : COMObject, NetOffice.OfficeApi.IAccessible_
    {
        #region Ctor

        /// <summary>
        /// Stub Ctor, not indented to use
        /// </summary>
        public IAccessible_() : base()
        {

        }

        #endregion

        #region Properties

        /// <summary>
        /// SupportByVersion Office 9, 10, 11, 12, 14, 15, 16
        /// Get/Set
        /// </summary>
        /// <param name="varChild">optional object varChild</param>
        [SupportByVersion("Office", 9, 10, 11, 12, 14, 15, 16)]
        [EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
        public virtual string get_accName(object varChild)
        {
            return InvokerService.InvokeInternal.ExecuteStringPropertyGet(this, "accName", varChild);
        }

        /// <summary>
        /// SupportByVersion Office 9, 10, 11, 12, 14, 15, 16
        /// Get/Set
        /// </summary>
        /// <param name="varChild">optional object varChild</param>
        /// <param name="value">optional string value</param>
        [SupportByVersion("Office", 9, 10, 11, 12, 14, 15, 16)]
        [EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
        public virtual void set_accName(object varChild, string value)
        {
            InvokerService.InvokeInternal.ExecutePropertySet(this, "accName", varChild, value);
        }

        /// <summary>
        /// SupportByVersion Office 9, 10, 11, 12, 14, 15, 16
        /// Alias for get_accName
        /// </summary>
        /// <param name="varChild">optional object varChild</param>
        [SupportByVersion("Office", 9, 10, 11, 12, 14, 15, 16), Redirect("get_accName")]
        public virtual string accName(object varChild)
        {
            return get_accName(varChild);
        }

        /// <summary>
        /// SupportByVersion Office 9, 10, 11, 12, 14, 15, 16
        /// Get/Set
        /// </summary>
        /// <param name="varChild">optional object varChild</param>
        [SupportByVersion("Office", 9, 10, 11, 12, 14, 15, 16)]
        [EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
        public virtual string get_accValue(object varChild)
        {
            return InvokerService.InvokeInternal.ExecuteStringPropertyGet(this, "accValue", varChild);
        }

        /// <summary>
        /// SupportByVersion Office 9, 10, 11, 12, 14, 15, 16
        /// Get/Set
        /// </summary>
        /// <param name="varChild">optional object varChild</param>
        /// <param name="value">optional string value</param>
        [SupportByVersion("Office", 9, 10, 11, 12, 14, 15, 16)]
        [EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
        public virtual void set_accValue(object varChild, string value)
        {
            InvokerService.InvokeInternal.ExecutePropertySet(this, "accValue", varChild, value);
        }

        /// <summary>
        /// SupportByVersion Office 9, 10, 11, 12, 14, 15, 16
        /// Alias for get_accValue
        /// </summary>
        /// <param name="varChild">optional object varChild</param>
        [SupportByVersion("Office", 9, 10, 11, 12, 14, 15, 16), Redirect("get_accValue")]
        public virtual string accValue(object varChild)
        {
            return get_accValue(varChild);
        }

        /// <summary>
        /// SupportByVersion Office 9, 10, 11, 12, 14, 15, 16
        /// Get
        /// </summary>
        /// <param name="varChild">optional object varChild</param>
        [SupportByVersion("Office", 9, 10, 11, 12, 14, 15, 16)]
        [EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
        public virtual string get_accDescription(object varChild)
        {
            return InvokerService.InvokeInternal.ExecuteStringPropertyGet(this, "accDescription", varChild);
        }

        /// <summary>
        /// SupportByVersion Office 9, 10, 11, 12, 14, 15, 16
        /// Alias for get_accDescription
        /// </summary>
        /// <param name="varChild">optional object varChild</param>
        [SupportByVersion("Office", 9, 10, 11, 12, 14, 15, 16), Redirect("get_accDescription")]
        public virtual string accDescription(object varChild)
        {
            return get_accDescription(varChild);
        }

        /// <summary>
        /// SupportByVersion Office 9, 10, 11, 12, 14, 15, 16
        /// Get
        /// </summary>
        /// <param name="varChild">optional object varChild</param>
        [SupportByVersion("Office", 9, 10, 11, 12, 14, 15, 16)]
        [EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
        public virtual object get_accRole(object varChild)
        {
            return InvokerService.InvokeInternal.ExecuteVariantPropertyGet(this, "accRole", varChild);
        }

        /// <summary>
        /// SupportByVersion Office 9, 10, 11, 12, 14, 15, 16
        /// Alias for get_accRole
        /// </summary>
        /// <param name="varChild">optional object varChild</param>
        [SupportByVersion("Office", 9, 10, 11, 12, 14, 15, 16), Redirect("get_accRole")]
        public virtual object accRole(object varChild)
        {
            return get_accRole(varChild);
        }

        /// <summary>
        /// SupportByVersion Office 9, 10, 11, 12, 14, 15, 16
        /// Get
        /// </summary>
        /// <param name="varChild">optional object varChild</param>
        [SupportByVersion("Office", 9, 10, 11, 12, 14, 15, 16)]
        [EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
        public virtual object get_accState(object varChild)
        {
            return InvokerService.InvokeInternal.ExecuteVariantPropertyGet(this, "accState", varChild);
        }

        /// <summary>
        /// SupportByVersion Office 9, 10, 11, 12, 14, 15, 16
        /// Alias for get_accState
        /// </summary>
        /// <param name="varChild">optional object varChild</param>
        [SupportByVersion("Office", 9, 10, 11, 12, 14, 15, 16), Redirect("get_accState")]
        public virtual object accState(object varChild)
        {
            return get_accState(varChild);
        }

        /// <summary>
        /// SupportByVersion Office 9, 10, 11, 12, 14, 15, 16
        /// Get
        /// </summary>
        /// <param name="varChild">optional object varChild</param>
        [SupportByVersion("Office", 9, 10, 11, 12, 14, 15, 16)]
        [EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
        public virtual string get_accHelp(object varChild)
        {
            return InvokerService.InvokeInternal.ExecuteStringPropertyGet(this, "accHelp", varChild);
        }

        /// <summary>
        /// SupportByVersion Office 9, 10, 11, 12, 14, 15, 16
        /// Alias for get_accHelp
        /// </summary>
        /// <param name="varChild">optional object varChild</param>
        [SupportByVersion("Office", 9, 10, 11, 12, 14, 15, 16), Redirect("get_accHelp")]
        public virtual string accHelp(object varChild)
        {
            return get_accHelp(varChild);
        }

        /// <summary>
        /// SupportByVersion Office 9, 10, 11, 12, 14, 15, 16
        /// Get
        /// </summary>
        /// <param name="varChild">optional object varChild</param>
        [SupportByVersion("Office", 9, 10, 11, 12, 14, 15, 16)]
        [EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
        public virtual string get_accKeyboardShortcut(object varChild)
        {
            return InvokerService.InvokeInternal.ExecuteStringPropertyGet(this, "accKeyboardShortcut", varChild);
        }

        /// <summary>
        /// SupportByVersion Office 9, 10, 11, 12, 14, 15, 16
        /// Alias for get_accKeyboardShortcut
        /// </summary>
        /// <param name="varChild">optional object varChild</param>
        [SupportByVersion("Office", 9, 10, 11, 12, 14, 15, 16), Redirect("get_accKeyboardShortcut")]
        public virtual string accKeyboardShortcut(object varChild)
        {
            return get_accKeyboardShortcut(varChild);
        }

        /// <summary>
        /// SupportByVersion Office 9, 10, 11, 12, 14, 15, 16
        /// Get
        /// </summary>
        /// <param name="varChild">optional object varChild</param>
        [SupportByVersion("Office", 9, 10, 11, 12, 14, 15, 16)]
        [EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
        public virtual string get_accDefaultAction(object varChild)
        {
            return InvokerService.InvokeInternal.ExecuteStringPropertyGet(this, "accDefaultAction", varChild);
        }

        /// <summary>
        /// SupportByVersion Office 9, 10, 11, 12, 14, 15, 16
        /// Alias for get_accDefaultAction
        /// </summary>
        /// <param name="varChild">optional object varChild</param>
        [SupportByVersion("Office", 9, 10, 11, 12, 14, 15, 16), Redirect("get_accDefaultAction")]
        public virtual string accDefaultAction(object varChild)
        {
            return get_accDefaultAction(varChild);
        }

        #endregion

        #region Methods

        #endregion
    }

    /// <summary>
    /// DispatchInterface IAccessible
    /// SupportByVersion Office, 9,10,11,12,14,15,16
    /// </summary>
    /// <remarks> MSDN Online: https://msdn.microsoft.com/en-us/library/microsoft.office.core.iaccessible.aspx </remarks>
    [SupportByVersion("Office", 9, 10, 11, 12, 14, 15, 16)]
    [EntityType(EntityType.IsDispatchInterface), BaseType]
    public class IAccessible : NetOffice.OfficeApi.Behind.IAccessible_, NetOffice.OfficeApi.IAccessible
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
                    _contractType = typeof(NetOffice.OfficeApi.IAccessible);
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
                    _type = typeof(IAccessible);
                return _type;
            }
        }

        #endregion

        #region Ctor

        /// <summary>
        /// Stub Ctor, not indented to use
        /// </summary>
        public IAccessible() : base()
        {

        }

        #endregion

        #region Properties

        /// <summary>
        /// SupportByVersion Office 9, 10, 11, 12, 14, 15, 16
        /// Get
        /// Unknown COM Proxy
        /// </summary>
        [SupportByVersion("Office", 9, 10, 11, 12, 14, 15, 16), ProxyResult]
        [EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
        public virtual object accParent
        {
            get
            {
                return InvokerService.InvokeInternal.ExecuteReferencePropertyGet(this, "accParent");
            }
        }

        /// <summary>
        /// SupportByVersion Office 9, 10, 11, 12, 14, 15, 16
        /// Get
        /// </summary>
        [SupportByVersion("Office", 9, 10, 11, 12, 14, 15, 16)]
        [EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
        public virtual Int32 accChildCount
        {
            get
            {
                return InvokerService.InvokeInternal.ExecuteInt32PropertyGet(this, "accChildCount");
            }
        }

        /// <summary>
        /// SupportByVersion Office 9, 10, 11, 12, 14, 15, 16
        /// Get
        /// Unknown COM Proxy
        /// </summary>
        /// <param name="varChild">object varChild</param>
        [SupportByVersion("Office", 9, 10, 11, 12, 14, 15, 16), ProxyResult]
        [EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
        public virtual object get_accChild(object varChild)
        {
            return InvokerService.InvokeInternal.ExecuteReferencePropertyGet(this, "accChild", varChild);
        }

        /// <summary>
        /// SupportByVersion Office 9, 10, 11, 12, 14, 15, 16
        /// Alias for get_accChild
        /// Unknown COM Proxy
        /// </summary>
        /// <param name="varChild">object varChild</param>
        [SupportByVersion("Office", 9, 10, 11, 12, 14, 15, 16), ProxyResult, Redirect("get_accChild")]
        public virtual object accChild(object varChild)
        {
            return get_accChild(varChild);
        }

        /// <summary>
        /// SupportByVersion Office 9, 10, 11, 12, 14, 15, 16
        /// Get/Set
        /// </summary>
        [SupportByVersion("Office", 9, 10, 11, 12, 14, 15, 16)]
        [EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
        public virtual string accName
        {
            get
            {
                return InvokerService.InvokeInternal.ExecuteStringPropertyGet(this, "accName");
            }
            set
            {
                InvokerService.InvokeInternal.ExecuteValuePropertySet(this, "accName", value);
            }
        }

        /// <summary>
        /// SupportByVersion Office 9, 10, 11, 12, 14, 15, 16
        /// Get/Set
        /// </summary>
        [SupportByVersion("Office", 9, 10, 11, 12, 14, 15, 16)]
        [EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
        public virtual string accValue
        {
            get
            {
                return InvokerService.InvokeInternal.ExecuteStringPropertyGet(this, "accValue");
            }
            set
            {
                InvokerService.InvokeInternal.ExecuteValuePropertySet(this, "accValue", value);
            }
        }

        /// <summary>
        /// SupportByVersion Office 9, 10, 11, 12, 14, 15, 16
        /// Get
        /// </summary>
        [SupportByVersion("Office", 9, 10, 11, 12, 14, 15, 16)]
        [EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
        public virtual string accDescription
        {
            get
            {
                return InvokerService.InvokeInternal.ExecuteStringPropertyGet(this, "accDescription");
            }
        }

        /// <summary>
        /// SupportByVersion Office 9, 10, 11, 12, 14, 15, 16
        /// Get
        /// </summary>
        [SupportByVersion("Office", 9, 10, 11, 12, 14, 15, 16)]
        [EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
        public virtual object accRole
        {
            get
            {
                return InvokerService.InvokeInternal.ExecuteVariantPropertyGet(this, "accRole");
            }
        }

        /// <summary>
        /// SupportByVersion Office 9, 10, 11, 12, 14, 15, 16
        /// Get
        /// </summary>
        [SupportByVersion("Office", 9, 10, 11, 12, 14, 15, 16)]
        [EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
        public virtual object accState
        {
            get
            {
                return InvokerService.InvokeInternal.ExecuteVariantPropertyGet(this, "accState");
            }
        }

        /// <summary>
        /// SupportByVersion Office 9, 10, 11, 12, 14, 15, 16
        /// Get
        /// </summary>
        [SupportByVersion("Office", 9, 10, 11, 12, 14, 15, 16)]
        [EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
        public virtual string accHelp
        {
            get
            {
                return InvokerService.InvokeInternal.ExecuteStringPropertyGet(this, "accHelp");
            }
        }

        /// <summary>
        /// SupportByVersion Office 9, 10, 11, 12, 14, 15, 16
        /// Get
        /// </summary>
        /// <param name="pszHelpFile">string pszHelpFile</param>
        /// <param name="varChild">optional object varChild</param>
        [SupportByVersion("Office", 9, 10, 11, 12, 14, 15, 16)]
        [EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
        public virtual Int32 get_accHelpTopic(out string pszHelpFile, object varChild)
        {
            ParameterModifier[] modifiers = Invoker.CreateParamModifiers(true, false);
            pszHelpFile = string.Empty;
            object[] paramsArray = Invoker.ValidateParamsArray(pszHelpFile, varChild);
            object returnItem = Invoker.PropertyGet(this, "accHelpTopic", paramsArray, modifiers);
            pszHelpFile = paramsArray[0] as string;
            return NetRuntimeSystem.Convert.ToInt32(returnItem);
        }

        /// <summary>
        /// SupportByVersion Office 9, 10, 11, 12, 14, 15, 16
        /// Alias for get_accHelpTopic
        /// </summary>
        /// <param name="pszHelpFile">string pszHelpFile</param>
        /// <param name="varChild">optional object varChild</param>
        [SupportByVersion("Office", 9, 10, 11, 12, 14, 15, 16), Redirect("get_accHelpTopic")]
        public virtual Int32 accHelpTopic(out string pszHelpFile, object varChild)
        {
            return get_accHelpTopic(out pszHelpFile, varChild);
        }

        /// <summary>
        /// SupportByVersion Office 9, 10, 11, 12, 14, 15, 16
        /// Get
        /// </summary>
        /// <param name="pszHelpFile">string pszHelpFile</param>
        [SupportByVersion("Office", 9, 10, 11, 12, 14, 15, 16)]
        [EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
        public virtual Int32 get_accHelpTopic(out string pszHelpFile)
        {
            ParameterModifier[] modifiers = Invoker.CreateParamModifiers(true);
            pszHelpFile = string.Empty;
            object[] paramsArray = new object[] { pszHelpFile };

            Int32 returnItem = InvokerService.InvokeInternal.ExecuteInt32PropertyGetExtended(this, "accHelpTopic", paramsArray, modifiers);

            pszHelpFile = paramsArray[0] as string;
            return returnItem;
        }

        /// <summary>
        /// SupportByVersion Office 9, 10, 11, 12, 14, 15, 16
        /// Alias for get_accHelpTopic
        /// </summary>
        /// <param name="pszHelpFile">string pszHelpFile</param>
        [SupportByVersion("Office", 9, 10, 11, 12, 14, 15, 16), Redirect("get_accHelpTopic")]
        public virtual Int32 accHelpTopic(out string pszHelpFile)
        {
            return get_accHelpTopic(out pszHelpFile);
        }

        /// <summary>
        /// SupportByVersion Office 9, 10, 11, 12, 14, 15, 16
        /// Get
        /// </summary>
        [SupportByVersion("Office", 9, 10, 11, 12, 14, 15, 16)]
        [EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
        public virtual string accKeyboardShortcut
        {
            get
            {
                return InvokerService.InvokeInternal.ExecuteStringPropertyGet(this, "accKeyboardShortcut");
            }
        }

        /// <summary>
        /// SupportByVersion Office 9, 10, 11, 12, 14, 15, 16
        /// Get
        /// </summary>
        [SupportByVersion("Office", 9, 10, 11, 12, 14, 15, 16)]
        [EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
        public virtual object accFocus
        {
            get
            {
                return InvokerService.InvokeInternal.ExecuteVariantPropertyGet(this, "accFocus");
            }
        }

        /// <summary>
        /// SupportByVersion Office 9, 10, 11, 12, 14, 15, 16
        /// Get
        /// </summary>
        [SupportByVersion("Office", 9, 10, 11, 12, 14, 15, 16)]
        [EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
        public virtual object accSelection
        {
            get
            {
                return InvokerService.InvokeInternal.ExecuteVariantPropertyGet(this, "accSelection");
            }
        }

        /// <summary>
        /// SupportByVersion Office 9, 10, 11, 12, 14, 15, 16
        /// Get
        /// </summary>
        [SupportByVersion("Office", 9, 10, 11, 12, 14, 15, 16)]
        [EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
        public virtual string accDefaultAction
        {
            get
            {
                return InvokerService.InvokeInternal.ExecuteStringPropertyGet(this, "accDefaultAction");
            }
        }

        #endregion

        #region Methods

        /// <summary>
        /// SupportByVersion Office 9, 10, 11, 12, 14, 15, 16
        /// </summary>
        /// <param name="flagsSelect">Int32 flagsSelect</param>
        /// <param name="varChild">optional object varChild</param>
        [EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
        [SupportByVersion("Office", 9, 10, 11, 12, 14, 15, 16)]
        public virtual void accSelect(Int32 flagsSelect, object varChild)
        {
            InvokerService.InvokeInternal.ExecuteMethod(this, "accSelect", flagsSelect, varChild);
        }

        /// <summary>
        /// SupportByVersion Office 9, 10, 11, 12, 14, 15, 16
        /// </summary>
        /// <param name="flagsSelect">Int32 flagsSelect</param>
        [EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
        [CustomMethod]
        [SupportByVersion("Office", 9, 10, 11, 12, 14, 15, 16)]
        public virtual void accSelect(Int32 flagsSelect)
        {
            InvokerService.InvokeInternal.ExecuteMethod(this, "accSelect", flagsSelect);
        }

        /// <summary>
        /// SupportByVersion Office 9, 10, 11, 12, 14, 15, 16
        /// </summary>
        /// <param name="pxLeft">Int32 pxLeft</param>
        /// <param name="pyTop">Int32 pyTop</param>
        /// <param name="pcxWidth">Int32 pcxWidth</param>
        /// <param name="pcyHeight">Int32 pcyHeight</param>
        /// <param name="varChild">optional object varChild</param>
        [EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
        [SupportByVersion("Office", 9, 10, 11, 12, 14, 15, 16)]
        public virtual void accLocation(out Int32 pxLeft, out Int32 pyTop, out Int32 pcxWidth, out Int32 pcyHeight, object varChild)
        {
            ParameterModifier[] modifiers = Invoker.CreateParamModifiers(true, true, true, true, false);
            pxLeft = 0;
            pyTop = 0;
            pcxWidth = 0;
            pcyHeight = 0;
            object[] paramsArray = new object[] { pxLeft, pyTop, pcxWidth, pcyHeight, varChild };

            InvokerService.InvokeInternal.ExecuteMethodExtended(this, "accLocation", paramsArray, modifiers);

            pxLeft = (Int32)paramsArray[0];
            pyTop = (Int32)paramsArray[1];
            pcxWidth = (Int32)paramsArray[2];
            pcyHeight = (Int32)paramsArray[3];
        }

        /// <summary>
        /// SupportByVersion Office 9, 10, 11, 12, 14, 15, 16
        /// </summary>
        /// <param name="pxLeft">Int32 pxLeft</param>
        /// <param name="pyTop">Int32 pyTop</param>
        /// <param name="pcxWidth">Int32 pcxWidth</param>
        /// <param name="pcyHeight">Int32 pcyHeight</param>
        [EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
        [CustomMethod]
        [SupportByVersion("Office", 9, 10, 11, 12, 14, 15, 16)]
        public virtual void accLocation(out Int32 pxLeft, out Int32 pyTop, out Int32 pcxWidth, out Int32 pcyHeight)
        {
            ParameterModifier[] modifiers = Invoker.CreateParamModifiers(true, true, true, true);
            pxLeft = 0;
            pyTop = 0;
            pcxWidth = 0;
            pcyHeight = 0;
            object[] paramsArray = new object[] { pxLeft, pyTop, pcxWidth, pcyHeight };

            InvokerService.InvokeInternal.ExecuteMethodExtended(this, "accLocation", paramsArray, modifiers);

            pxLeft = (Int32)paramsArray[0];
            pyTop = (Int32)paramsArray[1];
            pcxWidth = (Int32)paramsArray[2];
            pcyHeight = (Int32)paramsArray[3];
        }

        /// <summary>
        /// SupportByVersion Office 9, 10, 11, 12, 14, 15, 16
        /// </summary>
        /// <param name="navDir">Int32 navDir</param>
        /// <param name="varStart">optional object varStart</param>
        [EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
        [SupportByVersion("Office", 9, 10, 11, 12, 14, 15, 16)]
        public virtual object accNavigate(Int32 navDir, object varStart)
        {
            return InvokerService.InvokeInternal.ExecuteVariantMethodGet(this, "accNavigate", navDir, varStart);
        }

        /// <summary>
        /// SupportByVersion Office 9, 10, 11, 12, 14, 15, 16
        /// </summary>
        /// <param name="navDir">Int32 navDir</param>
        [EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
        [CustomMethod]
        [SupportByVersion("Office", 9, 10, 11, 12, 14, 15, 16)]
        public virtual object accNavigate(Int32 navDir)
        {
            return InvokerService.InvokeInternal.ExecuteVariantMethodGet(this, "accNavigate", navDir);
        }

        /// <summary>
        /// SupportByVersion Office 9, 10, 11, 12, 14, 15, 16
        /// </summary>
        /// <param name="xLeft">Int32 xLeft</param>
        /// <param name="yTop">Int32 yTop</param>
        [EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
        [SupportByVersion("Office", 9, 10, 11, 12, 14, 15, 16)]
        public virtual object accHitTest(Int32 xLeft, Int32 yTop)
        {
            return InvokerService.InvokeInternal.ExecuteVariantMethodGet(this, "accHitTest", xLeft, yTop);
        }

        /// <summary>
        /// SupportByVersion Office 9, 10, 11, 12, 14, 15, 16
        /// </summary>
        /// <param name="varChild">optional object varChild</param>
        [EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
        [SupportByVersion("Office", 9, 10, 11, 12, 14, 15, 16)]
        public virtual void accDoDefaultAction(object varChild)
        {
            InvokerService.InvokeInternal.ExecuteMethod(this, "accDoDefaultAction", varChild);
        }

        /// <summary>
        /// SupportByVersion Office 9, 10, 11, 12, 14, 15, 16
        /// </summary>
        [EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
        [CustomMethod]
        [SupportByVersion("Office", 9, 10, 11, 12, 14, 15, 16)]
        public virtual void accDoDefaultAction()
        {
            InvokerService.InvokeInternal.ExecuteMethod(this, "accDoDefaultAction");
        }

        #endregion

        #pragma warning restore
    }
}
