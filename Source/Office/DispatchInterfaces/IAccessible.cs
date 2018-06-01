using System.Reflection;
using System;
using NetRuntimeSystem = System;
using System.ComponentModel;
using NetOffice.Attributes;

namespace NetOffice.OfficeApi
{
    /// <summary>
    /// IAccessible
    /// </summary>
    [SyntaxBypass]
    public interface IAccessible_ : ICOMObject
    {
        #region Properties

        /// <summary>
        /// SupportByVersion Office 9, 10, 11, 12, 14, 15, 16
        /// Get/Set
        /// </summary>
        /// <param name="varChild">optional object varChild</param>
        [SupportByVersion("Office", 9, 10, 11, 12, 14, 15, 16)]
        [EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
        string get_accName(object varChild);

        /// <summary>
        /// SupportByVersion Office 9, 10, 11, 12, 14, 15, 16
        /// Get/Set
        /// </summary>
        /// <param name="varChild">optional object varChild</param>
        /// <param name="value">optional string value</param>
        [SupportByVersion("Office", 9, 10, 11, 12, 14, 15, 16)]
        [EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
        void set_accName(object varChild, string value);

        /// <summary>
        /// SupportByVersion Office 9, 10, 11, 12, 14, 15, 16
        /// Alias for get_accName
        /// </summary>
        /// <param name="varChild">optional object varChild</param>
        [SupportByVersion("Office", 9, 10, 11, 12, 14, 15, 16), Redirect("get_accName")]
        string accName(object varChild);

        /// <summary>
        /// SupportByVersion Office 9, 10, 11, 12, 14, 15, 16
        /// Get/Set
        /// </summary>
        /// <param name="varChild">optional object varChild</param>
        [SupportByVersion("Office", 9, 10, 11, 12, 14, 15, 16)]
        [EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
        string get_accValue(object varChild);

        /// <summary>
        /// SupportByVersion Office 9, 10, 11, 12, 14, 15, 16
        /// Get/Set
        /// </summary>
        /// <param name="varChild">optional object varChild</param>
        /// <param name="value">optional string value</param>
        [SupportByVersion("Office", 9, 10, 11, 12, 14, 15, 16)]
        [EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
        void set_accValue(object varChild, string value);

        /// <summary>
        /// SupportByVersion Office 9, 10, 11, 12, 14, 15, 16
        /// Alias for get_accValue
        /// </summary>
        /// <param name="varChild">optional object varChild</param>
        [SupportByVersion("Office", 9, 10, 11, 12, 14, 15, 16), Redirect("get_accValue")]
        string accValue(object varChild);

        /// <summary>
        /// SupportByVersion Office 9, 10, 11, 12, 14, 15, 16
        /// Get
        /// </summary>
        /// <param name="varChild">optional object varChild</param>
        [SupportByVersion("Office", 9, 10, 11, 12, 14, 15, 16)]
        [EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
        string get_accDescription(object varChild);

        /// <summary>
        /// SupportByVersion Office 9, 10, 11, 12, 14, 15, 16
        /// Alias for get_accDescription
        /// </summary>
        /// <param name="varChild">optional object varChild</param>
        [SupportByVersion("Office", 9, 10, 11, 12, 14, 15, 16), Redirect("get_accDescription")]
        string accDescription(object varChild);

        /// <summary>
        /// SupportByVersion Office 9, 10, 11, 12, 14, 15, 16
        /// Get
        /// </summary>
        /// <param name="varChild">optional object varChild</param>
        [SupportByVersion("Office", 9, 10, 11, 12, 14, 15, 16)]
        [EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
        object get_accRole(object varChild);

        /// <summary>
        /// SupportByVersion Office 9, 10, 11, 12, 14, 15, 16
        /// Alias for get_accRole
        /// </summary>
        /// <param name="varChild">optional object varChild</param>
        [SupportByVersion("Office", 9, 10, 11, 12, 14, 15, 16), Redirect("get_accRole")]
        object accRole(object varChild);

        /// <summary>
        /// SupportByVersion Office 9, 10, 11, 12, 14, 15, 16
        /// Get
        /// </summary>
        /// <param name="varChild">optional object varChild</param>
        [SupportByVersion("Office", 9, 10, 11, 12, 14, 15, 16)]
        [EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
        object get_accState(object varChild);

        /// <summary>
        /// SupportByVersion Office 9, 10, 11, 12, 14, 15, 16
        /// Alias for get_accState
        /// </summary>
        /// <param name="varChild">optional object varChild</param>
        [SupportByVersion("Office", 9, 10, 11, 12, 14, 15, 16), Redirect("get_accState")]
        object accState(object varChild);

        /// <summary>
        /// SupportByVersion Office 9, 10, 11, 12, 14, 15, 16
        /// Get
        /// </summary>
        /// <param name="varChild">optional object varChild</param>
        [SupportByVersion("Office", 9, 10, 11, 12, 14, 15, 16)]
        [EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
        string get_accHelp(object varChild);

        /// <summary>
        /// SupportByVersion Office 9, 10, 11, 12, 14, 15, 16
        /// Alias for get_accHelp
        /// </summary>
        /// <param name="varChild">optional object varChild</param>
        [SupportByVersion("Office", 9, 10, 11, 12, 14, 15, 16), Redirect("get_accHelp")]
        string accHelp(object varChild);

        /// <summary>
        /// SupportByVersion Office 9, 10, 11, 12, 14, 15, 16
        /// Get
        /// </summary>
        /// <param name="varChild">optional object varChild</param>
        [SupportByVersion("Office", 9, 10, 11, 12, 14, 15, 16)]
        [EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
        string get_accKeyboardShortcut(object varChild);

        /// <summary>
        /// SupportByVersion Office 9, 10, 11, 12, 14, 15, 16
        /// Alias for get_accKeyboardShortcut
        /// </summary>
        /// <param name="varChild">optional object varChild</param>
        [SupportByVersion("Office", 9, 10, 11, 12, 14, 15, 16), Redirect("get_accKeyboardShortcut")]
        string accKeyboardShortcut(object varChild);

        /// <summary>
        /// SupportByVersion Office 9, 10, 11, 12, 14, 15, 16
        /// Get
        /// </summary>
        /// <param name="varChild">optional object varChild</param>
        [SupportByVersion("Office", 9, 10, 11, 12, 14, 15, 16)]
        [EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
        string get_accDefaultAction(object varChild);

        /// <summary>
        /// SupportByVersion Office 9, 10, 11, 12, 14, 15, 16
        /// Alias for get_accDefaultAction
        /// </summary>
        /// <param name="varChild">optional object varChild</param>
        [SupportByVersion("Office", 9, 10, 11, 12, 14, 15, 16), Redirect("get_accDefaultAction")]
        string accDefaultAction(object varChild);

        #endregion
    }

    /// <summary>
    /// DispatchInterface IAccessible
    /// SupportByVersion Office, 9,10,11,12,14,15,16
    /// </summary>
    /// <remarks> MSDN Online: https://msdn.microsoft.com/en-us/library/microsoft.office.core.iaccessible.aspx </remarks>
    [SupportByVersion("Office", 9, 10, 11, 12, 14, 15, 16)]
    [EntityType(EntityType.IsDispatchInterface), BaseType]
    public interface IAccessible : IAccessible_
    {
        #region Properties

        /// <summary>
        /// SupportByVersion Office 9, 10, 11, 12, 14, 15, 16
        /// Get
        /// Unknown COM Proxy
        /// </summary>
        [SupportByVersion("Office", 9, 10, 11, 12, 14, 15, 16), ProxyResult]
        [EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
        object accParent { get; }

        /// <summary>
        /// SupportByVersion Office 9, 10, 11, 12, 14, 15, 16
        /// Get
        /// </summary>
        [SupportByVersion("Office", 9, 10, 11, 12, 14, 15, 16)]
        [EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
        Int32 accChildCount { get; }

        /// <summary>
        /// SupportByVersion Office 9, 10, 11, 12, 14, 15, 16
        /// Get
        /// Unknown COM Proxy
        /// </summary>
        /// <param name="varChild">object varChild</param>
        [SupportByVersion("Office", 9, 10, 11, 12, 14, 15, 16), ProxyResult]
        [EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
        object get_accChild(object varChild);

        /// <summary>
        /// SupportByVersion Office 9, 10, 11, 12, 14, 15, 16
        /// Alias for get_accChild
        /// Unknown COM Proxy
        /// </summary>
        /// <param name="varChild">object varChild</param>
        [SupportByVersion("Office", 9, 10, 11, 12, 14, 15, 16), ProxyResult, Redirect("get_accChild")]
        object accChild(object varChild);

        /// <summary>
        /// SupportByVersion Office 9, 10, 11, 12, 14, 15, 16
        /// Get/Set
        /// </summary>
        [SupportByVersion("Office", 9, 10, 11, 12, 14, 15, 16)]
        [EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
        new string accName { get; set; }

        /// <summary>
        /// SupportByVersion Office 9, 10, 11, 12, 14, 15, 16
        /// Get/Set
        /// </summary>
        [SupportByVersion("Office", 9, 10, 11, 12, 14, 15, 16)]
        [EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
        new string accValue { get; set; }

        /// <summary>
        /// SupportByVersion Office 9, 10, 11, 12, 14, 15, 16
        /// Get
        /// </summary>
        [SupportByVersion("Office", 9, 10, 11, 12, 14, 15, 16)]
        [EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
        new string accDescription { get; }

        /// <summary>
        /// SupportByVersion Office 9, 10, 11, 12, 14, 15, 16
        /// Get
        /// </summary>
        [SupportByVersion("Office", 9, 10, 11, 12, 14, 15, 16)]
        [EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
        new object accRole { get; }

        /// <summary>
        /// SupportByVersion Office 9, 10, 11, 12, 14, 15, 16
        /// Get
        /// </summary>
        [SupportByVersion("Office", 9, 10, 11, 12, 14, 15, 16)]
        [EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
        new object accState { get; }

        /// <summary>
        /// SupportByVersion Office 9, 10, 11, 12, 14, 15, 16
        /// Get
        /// </summary>
        [SupportByVersion("Office", 9, 10, 11, 12, 14, 15, 16)]
        [EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
        new string accHelp { get; }

        /// <summary>
        /// SupportByVersion Office 9, 10, 11, 12, 14, 15, 16
        /// Get
        /// </summary>
        /// <param name="pszHelpFile">string pszHelpFile</param>
        /// <param name="varChild">optional object varChild</param>
        [SupportByVersion("Office", 9, 10, 11, 12, 14, 15, 16)]
        [EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
        Int32 get_accHelpTopic(out string pszHelpFile, object varChild);

        /// <summary>
        /// SupportByVersion Office 9, 10, 11, 12, 14, 15, 16
        /// Alias for get_accHelpTopic
        /// </summary>
        /// <param name="pszHelpFile">string pszHelpFile</param>
        /// <param name="varChild">optional object varChild</param>
        [SupportByVersion("Office", 9, 10, 11, 12, 14, 15, 16), Redirect("get_accHelpTopic")]
        Int32 accHelpTopic(out string pszHelpFile, object varChild);

        /// <summary>
        /// SupportByVersion Office 9, 10, 11, 12, 14, 15, 16
        /// Get
        /// </summary>
        /// <param name="pszHelpFile">string pszHelpFile</param>
        [SupportByVersion("Office", 9, 10, 11, 12, 14, 15, 16)]
        [EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
        Int32 get_accHelpTopic(out string pszHelpFile);

        /// <summary>
        /// SupportByVersion Office 9, 10, 11, 12, 14, 15, 16
        /// Alias for get_accHelpTopic
        /// </summary>
        /// <param name="pszHelpFile">string pszHelpFile</param>
        [SupportByVersion("Office", 9, 10, 11, 12, 14, 15, 16), Redirect("get_accHelpTopic")]
        Int32 accHelpTopic(out string pszHelpFile);

        /// <summary>
        /// SupportByVersion Office 9, 10, 11, 12, 14, 15, 16
        /// Get
        /// </summary>
        [SupportByVersion("Office", 9, 10, 11, 12, 14, 15, 16)]
        [EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
        new string accKeyboardShortcut { get; }

        /// <summary>
        /// SupportByVersion Office 9, 10, 11, 12, 14, 15, 16
        /// Get
        /// </summary>
        [SupportByVersion("Office", 9, 10, 11, 12, 14, 15, 16)]
        [EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
        object accFocus { get; }

        /// <summary>
        /// SupportByVersion Office 9, 10, 11, 12, 14, 15, 16
        /// Get
        /// </summary>
        [SupportByVersion("Office", 9, 10, 11, 12, 14, 15, 16)]
        [EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
        object accSelection { get; }

        /// <summary>
        /// SupportByVersion Office 9, 10, 11, 12, 14, 15, 16
        /// Get
        /// </summary>
        [SupportByVersion("Office", 9, 10, 11, 12, 14, 15, 16)]
        [EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
        new string accDefaultAction { get; }

        #endregion

        #region Methods

        /// <summary>
        /// SupportByVersion Office 9, 10, 11, 12, 14, 15, 16
        /// </summary>
        /// <param name="flagsSelect">Int32 flagsSelect</param>
        /// <param name="varChild">optional object varChild</param>
        [EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
        [SupportByVersion("Office", 9, 10, 11, 12, 14, 15, 16)]
        void accSelect(Int32 flagsSelect, object varChild);

        /// <summary>
        /// SupportByVersion Office 9, 10, 11, 12, 14, 15, 16
        /// </summary>
        /// <param name="flagsSelect">Int32 flagsSelect</param>
        [EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
        [CustomMethod]
        [SupportByVersion("Office", 9, 10, 11, 12, 14, 15, 16)]
        void accSelect(Int32 flagsSelect);

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
        void accLocation(out Int32 pxLeft, out Int32 pyTop, out Int32 pcxWidth, out Int32 pcyHeight, object varChild);

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
        void accLocation(out Int32 pxLeft, out Int32 pyTop, out Int32 pcxWidth, out Int32 pcyHeight);

        /// <summary>
        /// SupportByVersion Office 9, 10, 11, 12, 14, 15, 16
        /// </summary>
        /// <param name="navDir">Int32 navDir</param>
        /// <param name="varStart">optional object varStart</param>
        [EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
        [SupportByVersion("Office", 9, 10, 11, 12, 14, 15, 16)]
        object accNavigate(Int32 navDir, object varStart);

        /// <summary>
        /// SupportByVersion Office 9, 10, 11, 12, 14, 15, 16
        /// </summary>
        /// <param name="navDir">Int32 navDir</param>
        [EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
        [CustomMethod]
        [SupportByVersion("Office", 9, 10, 11, 12, 14, 15, 16)]
        object accNavigate(Int32 navDir);

        /// <summary>
        /// SupportByVersion Office 9, 10, 11, 12, 14, 15, 16
        /// </summary>
        /// <param name="xLeft">Int32 xLeft</param>
        /// <param name="yTop">Int32 yTop</param>
        [EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
        [SupportByVersion("Office", 9, 10, 11, 12, 14, 15, 16)]
        object accHitTest(Int32 xLeft, Int32 yTop);

        /// <summary>
        /// SupportByVersion Office 9, 10, 11, 12, 14, 15, 16
        /// </summary>
        /// <param name="varChild">optional object varChild</param>
        [EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
        [SupportByVersion("Office", 9, 10, 11, 12, 14, 15, 16)]
        void accDoDefaultAction(object varChild);

        /// <summary>
        /// SupportByVersion Office 9, 10, 11, 12, 14, 15, 16
        /// </summary>
        [EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
        [CustomMethod]
        [SupportByVersion("Office", 9, 10, 11, 12, 14, 15, 16)]
        void accDoDefaultAction();

        #endregion
    }
}
