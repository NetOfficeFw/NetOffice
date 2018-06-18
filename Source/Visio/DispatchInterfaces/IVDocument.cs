using System.Reflection;
using System;
using NetRuntimeSystem = System;
using System.ComponentModel;
using NetOffice.Attributes;

namespace NetOffice.VisioApi
{
    /// <summary>
    /// IVDocument
    /// </summary>
    [SyntaxBypass]
    public interface IVDocument_ : ICOMObject
    {
        #region Properties

        /// <summary>
        /// SupportByVersion Visio 11, 12, 14, 15, 16
        /// Get/Set
        /// </summary>
        /// <param name="unitsNameOrCode">optional object unitsNameOrCode</param>
        [SupportByVersion("Visio", 11, 12, 14, 15, 16)]
        [EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
        Double get_LeftMargin(object unitsNameOrCode);

        /// <summary>
        /// SupportByVersion Visio 11, 12, 14, 15, 16
        /// Get/Set
        /// </summary>
        /// <param name="unitsNameOrCode">optional object unitsNameOrCode</param>
        /// <param name="value">optional double value</param>
        [SupportByVersion("Visio", 11, 12, 14, 15, 16)]
        [EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
        void set_LeftMargin(object unitsNameOrCode, Double value);

        /// <summary>
        /// SupportByVersion Visio 11, 12, 14, 15, 16
        /// Alias for get_LeftMargin
        /// </summary>
        /// <param name="unitsNameOrCode">optional object unitsNameOrCode</param>
        [SupportByVersion("Visio", 11, 12, 14, 15, 16), Redirect("get_LeftMargin")]
        Double LeftMargin(object unitsNameOrCode);

        /// <summary>
        /// SupportByVersion Visio 11, 12, 14, 15, 16
        /// Get/Set
        /// </summary>
        /// <param name="unitsNameOrCode">optional object unitsNameOrCode</param>
        [SupportByVersion("Visio", 11, 12, 14, 15, 16)]
        [EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
        Double get_RightMargin(object unitsNameOrCode);

        /// <summary>
        /// SupportByVersion Visio 11, 12, 14, 15, 16
        /// Get/Set
        /// </summary>
        /// <param name="unitsNameOrCode">optional object unitsNameOrCode</param>
        /// <param name="value">optional double value</param>
        [SupportByVersion("Visio", 11, 12, 14, 15, 16)]
        [EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
        void set_RightMargin(object unitsNameOrCode, Double value);

        /// <summary>
        /// SupportByVersion Visio 11, 12, 14, 15, 16
        /// Alias for get_RightMargin
        /// </summary>
        /// <param name="unitsNameOrCode">optional object unitsNameOrCode</param>
        [SupportByVersion("Visio", 11, 12, 14, 15, 16), Redirect("get_RightMargin")]
        Double RightMargin(object unitsNameOrCode);

        /// <summary>
        /// SupportByVersion Visio 11, 12, 14, 15, 16
        /// Get/Set
        /// </summary>
        /// <param name="unitsNameOrCode">optional object unitsNameOrCode</param>
        [SupportByVersion("Visio", 11, 12, 14, 15, 16)]
        [EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
        Double get_TopMargin(object unitsNameOrCode);

        /// <summary>
        /// SupportByVersion Visio 11, 12, 14, 15, 16
        /// Get/Set
        /// </summary>
        /// <param name="unitsNameOrCode">optional object unitsNameOrCode</param>
        /// <param name="value">optional double value</param>
        [SupportByVersion("Visio", 11, 12, 14, 15, 16)]
        [EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
        void set_TopMargin(object unitsNameOrCode, Double value);

        /// <summary>
        /// SupportByVersion Visio 11, 12, 14, 15, 16
        /// Alias for get_TopMargin
        /// </summary>
        /// <param name="unitsNameOrCode">optional object unitsNameOrCode</param>
        [SupportByVersion("Visio", 11, 12, 14, 15, 16), Redirect("get_TopMargin")]
        Double TopMargin(object unitsNameOrCode);

        /// <summary>
        /// SupportByVersion Visio 11, 12, 14, 15, 16
        /// Get/Set
        /// </summary>
        /// <param name="unitsNameOrCode">optional object unitsNameOrCode</param>
        [SupportByVersion("Visio", 11, 12, 14, 15, 16)]
        [EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
        Double get_BottomMargin(object unitsNameOrCode);

        /// <summary>
        /// SupportByVersion Visio 11, 12, 14, 15, 16
        /// Get/Set
        /// </summary>
        /// <param name="unitsNameOrCode">optional object unitsNameOrCode</param>
        /// <param name="value">optional double value</param>
        [SupportByVersion("Visio", 11, 12, 14, 15, 16)]
        [EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
        void set_BottomMargin(object unitsNameOrCode, Double value);

        /// <summary>
        /// SupportByVersion Visio 11, 12, 14, 15, 16
        /// Alias for get_BottomMargin
        /// </summary>
        /// <param name="unitsNameOrCode">optional object unitsNameOrCode</param>
        [SupportByVersion("Visio", 11, 12, 14, 15, 16), Redirect("get_BottomMargin")]
        Double BottomMargin(object unitsNameOrCode);

        /// <summary>
        /// SupportByVersion Visio 11, 12, 14, 15, 16
        /// Get/Set
        /// </summary>
        /// <param name="bstrPassword">optional object bstrPassword</param>
        [SupportByVersion("Visio", 11, 12, 14, 15, 16)]
        [EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
        NetOffice.VisioApi.Enums.VisProtection get_Protection(object bstrPassword);

        /// <summary>
        /// SupportByVersion Visio 11, 12, 14, 15, 16
        /// Get/Set
        /// </summary>
        /// <param name="bstrPassword">optional object bstrPassword</param>
        /// <param name="value">optional VisProtection value</param>
        [SupportByVersion("Visio", 11, 12, 14, 15, 16)]
        [EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
        void set_Protection(object bstrPassword, NetOffice.VisioApi.Enums.VisProtection value);

        /// <summary>
        /// SupportByVersion Visio 11, 12, 14, 15, 16
        /// Alias for get_Protection
        /// </summary>
        /// <param name="bstrPassword">optional object bstrPassword</param>
        [SupportByVersion("Visio", 11, 12, 14, 15, 16), Redirect("get_Protection")]
        NetOffice.VisioApi.Enums.VisProtection Protection(object bstrPassword);

        /// <summary>
        /// SupportByVersion Visio 11, 12, 14, 15, 16
        /// Get/Set
        /// </summary>
        /// <param name="unitsNameOrCode">optional object unitsNameOrCode</param>
        [SupportByVersion("Visio", 11, 12, 14, 15, 16)]
        [EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
        Double get_HeaderMargin(object unitsNameOrCode);

        /// <summary>
        /// SupportByVersion Visio 11, 12, 14, 15, 16
        /// Get/Set
        /// </summary>
        /// <param name="unitsNameOrCode">optional object unitsNameOrCode</param>
        /// <param name="value">optional double value</param>
        [SupportByVersion("Visio", 11, 12, 14, 15, 16)]
        [EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
        void set_HeaderMargin(object unitsNameOrCode, Double value);

        /// <summary>
        /// SupportByVersion Visio 11, 12, 14, 15, 16
        /// Alias for get_HeaderMargin
        /// </summary>
        /// <param name="unitsNameOrCode">optional object unitsNameOrCode</param>
        [SupportByVersion("Visio", 11, 12, 14, 15, 16), Redirect("get_HeaderMargin")]
        Double HeaderMargin(object unitsNameOrCode);

        /// <summary>
        /// SupportByVersion Visio 11, 12, 14, 15, 16
        /// Get/Set
        /// </summary>
        /// <param name="unitsNameOrCode">optional object unitsNameOrCode</param>
        [SupportByVersion("Visio", 11, 12, 14, 15, 16)]
        [EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
        Double get_FooterMargin(object unitsNameOrCode);

        /// <summary>
        /// SupportByVersion Visio 11, 12, 14, 15, 16
        /// Get/Set
        /// </summary>
        /// <param name="unitsNameOrCode">optional object unitsNameOrCode</param>
        /// <param name="value">optional double value</param>
        [SupportByVersion("Visio", 11, 12, 14, 15, 16)]
        [EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
        void set_FooterMargin(object unitsNameOrCode, Double value);

        /// <summary>
        /// SupportByVersion Visio 11, 12, 14, 15, 16
        /// Alias for get_FooterMargin
        /// </summary>
        /// <param name="unitsNameOrCode">optional object unitsNameOrCode</param>
        [SupportByVersion("Visio", 11, 12, 14, 15, 16), Redirect("get_FooterMargin")]
        Double FooterMargin(object unitsNameOrCode);

        /// <summary>
        /// SupportByVersion Visio 11, 12, 14, 15, 16
        /// Get/Set
        /// </summary>
        /// <param name="bstrExistingPassword">optional object bstrExistingPassword</param>
        [SupportByVersion("Visio", 11, 12, 14, 15, 16)]
        [EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
        string get_Password(object bstrExistingPassword);

        /// <summary>
        /// SupportByVersion Visio 11, 12, 14, 15, 16
        /// Get/Set
        /// </summary>
        /// <param name="bstrExistingPassword">optional object bstrExistingPassword</param>
        /// <param name="value">optional string value</param>
        [SupportByVersion("Visio", 11, 12, 14, 15, 16)]
        [EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
        void set_Password(object bstrExistingPassword, string value);

        /// <summary>
        /// SupportByVersion Visio 11, 12, 14, 15, 16
        /// Alias for get_Password
        /// </summary>
        /// <param name="bstrExistingPassword">optional object bstrExistingPassword</param>
        [SupportByVersion("Visio", 11, 12, 14, 15, 16), Redirect("get_Password")]
        string Password(object bstrExistingPassword);

        #endregion
    }
   
    /// <summary>
    /// DispatchInterface IVDocument 
    /// SupportByVersion Visio, 11,12,14,15,16
    /// </summary>
    [SupportByVersion("Visio", 11, 12, 14, 15, 16)]
    [EntityType(EntityType.IsDispatchInterface), BaseType]
	[TypeId("000D0705-0000-0000-C000-000000000046")]
    [CoClassSource(typeof(NetOffice.VisioApi.Document))]
    public interface IVDocument : IVDocument_
    {
        #region Properties

        /// <summary>
        /// SupportByVersion Visio 11, 12, 14, 15, 16
        /// Get
        /// </summary>
        [SupportByVersion("Visio", 11, 12, 14, 15, 16)]
        [BaseResult]
        NetOffice.VisioApi.IVApplication Application { get; }

        /// <summary>
        /// SupportByVersion Visio 11, 12, 14, 15, 16
        /// Get
        /// </summary>
        [SupportByVersion("Visio", 11, 12, 14, 15, 16)]
        Int16 Stat { get; }

        /// <summary>
        /// SupportByVersion Visio 11, 12, 14, 15, 16
        /// Get
        /// </summary>
        [SupportByVersion("Visio", 11, 12, 14, 15, 16)]
        Int16 ObjectType { get; }

        /// <summary>
        /// SupportByVersion Visio 11, 12, 14, 15, 16
        /// Get
        /// </summary>
        [SupportByVersion("Visio", 11, 12, 14, 15, 16)]
        Int16 InPlace { get; }

        /// <summary>
        /// SupportByVersion Visio 11, 12, 14, 15, 16
        /// Get
        /// </summary>
        [SupportByVersion("Visio", 11, 12, 14, 15, 16)]
        [BaseResult]
        NetOffice.VisioApi.IVMasters Masters { get; }

        /// <summary>
        /// SupportByVersion Visio 11, 12, 14, 15, 16
        /// Get
        /// </summary>
        [SupportByVersion("Visio", 11, 12, 14, 15, 16)]
        [BaseResult]
        NetOffice.VisioApi.IVPages Pages { get; }

        /// <summary>
        /// SupportByVersion Visio 11, 12, 14, 15, 16
        /// Get
        /// </summary>
        [SupportByVersion("Visio", 11, 12, 14, 15, 16)]
        [BaseResult]
        NetOffice.VisioApi.IVStyles Styles { get; }

        /// <summary>
        /// SupportByVersion Visio 11, 12, 14, 15, 16
        /// Get
        /// </summary>
        [SupportByVersion("Visio", 11, 12, 14, 15, 16)]
        string Name { get; }

        /// <summary>
        /// SupportByVersion Visio 11, 12, 14, 15, 16
        /// Get
        /// </summary>
        [SupportByVersion("Visio", 11, 12, 14, 15, 16)]
        string Path { get; }

        /// <summary>
        /// SupportByVersion Visio 11, 12, 14, 15, 16
        /// Get
        /// </summary>
        [SupportByVersion("Visio", 11, 12, 14, 15, 16)]
        string FullName { get; }

        /// <summary>
        /// SupportByVersion Visio 11, 12, 14, 15, 16
        /// Get
        /// </summary>
        [SupportByVersion("Visio", 11, 12, 14, 15, 16)]
        Int16 Index { get; }

        /// <summary>
        /// SupportByVersion Visio 11, 12, 14, 15, 16
        /// Get/Set
        /// </summary>
        [SupportByVersion("Visio", 11, 12, 14, 15, 16)]
        [EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
        Int16 old_Saved { get; set; }

        /// <summary>
        /// SupportByVersion Visio 11, 12, 14, 15, 16
        /// Get
        /// </summary>
        [SupportByVersion("Visio", 11, 12, 14, 15, 16)]
        Int16 ReadOnly { get; }

        /// <summary>
        /// SupportByVersion Visio 11, 12, 14, 15, 16
        /// Get/Set
        /// </summary>
        [SupportByVersion("Visio", 11, 12, 14, 15, 16)]
        [EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
        Int32 old_Version { get; set; }

        /// <summary>
        /// SupportByVersion Visio 11, 12, 14, 15, 16
        /// Get/Set
        /// </summary>
        [SupportByVersion("Visio", 11, 12, 14, 15, 16)]
        string Title { get; set; }

        /// <summary>
        /// SupportByVersion Visio 11, 12, 14, 15, 16
        /// Get/Set
        /// </summary>
        [SupportByVersion("Visio", 11, 12, 14, 15, 16)]
        string Subject { get; set; }

        /// <summary>
        /// SupportByVersion Visio 11, 12, 14, 15, 16
        /// Get/Set
        /// </summary>
        [SupportByVersion("Visio", 11, 12, 14, 15, 16)]
        string Creator { get; set; }

        /// <summary>
        /// SupportByVersion Visio 11, 12, 14, 15, 16
        /// Get/Set
        /// </summary>
        [SupportByVersion("Visio", 11, 12, 14, 15, 16)]
        string Keywords { get; set; }

        /// <summary>
        /// SupportByVersion Visio 11, 12, 14, 15, 16
        /// Get/Set
        /// </summary>
        [SupportByVersion("Visio", 11, 12, 14, 15, 16)]
        string Description { get; set; }

        /// <summary>
        /// SupportByVersion Visio 11, 12, 14, 15, 16
        /// Get
        /// </summary>
        [SupportByVersion("Visio", 11, 12, 14, 15, 16)]
        [BaseResult]
        NetOffice.VisioApi.IVUIObject CustomMenus { get; }

        /// <summary>
        /// SupportByVersion Visio 11, 12, 14, 15, 16
        /// Get/Set
        /// </summary>
        [SupportByVersion("Visio", 11, 12, 14, 15, 16)]
        string CustomMenusFile { get; set; }

        /// <summary>
        /// SupportByVersion Visio 11, 12, 14, 15, 16
        /// Get
        /// </summary>
        [SupportByVersion("Visio", 11, 12, 14, 15, 16)]
        [BaseResult]
        NetOffice.VisioApi.IVUIObject CustomToolbars { get; }

        /// <summary>
        /// SupportByVersion Visio 11, 12, 14, 15, 16
        /// Get/Set
        /// </summary>
        [SupportByVersion("Visio", 11, 12, 14, 15, 16)]
        string CustomToolbarsFile { get; set; }

        /// <summary>
        /// SupportByVersion Visio 11, 12, 14, 15, 16
        /// Get
        /// </summary>
        [SupportByVersion("Visio", 11, 12, 14, 15, 16)]
        [BaseResult]
        NetOffice.VisioApi.IVFonts Fonts { get; }

        /// <summary>
        /// SupportByVersion Visio 11, 12, 14, 15, 16
        /// Get
        /// </summary>
        [SupportByVersion("Visio", 11, 12, 14, 15, 16)]
        [BaseResult]
        NetOffice.VisioApi.IVColors Colors { get; }

        /// <summary>
        /// SupportByVersion Visio 11, 12, 14, 15, 16
        /// Get
        /// </summary>
        [SupportByVersion("Visio", 11, 12, 14, 15, 16)]
        [BaseResult]
        NetOffice.VisioApi.IVEventList EventList { get; }

        /// <summary>
        /// SupportByVersion Visio 11, 12, 14, 15, 16
        /// Get
        /// </summary>
        [SupportByVersion("Visio", 11, 12, 14, 15, 16)]
        string Template { get; }

        /// <summary>
        /// SupportByVersion Visio 11, 12, 14, 15, 16
        /// Get/Set
        /// </summary>
        [SupportByVersion("Visio", 11, 12, 14, 15, 16)]
        [EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
        Int16 old_SavePreviewMode { get; set; }

        /// <summary>
        /// SupportByVersion Visio 11, 12, 14, 15, 16
        /// Get/Set
        /// </summary>
        [SupportByVersion("Visio", 11, 12, 14, 15, 16)]
        new Double LeftMargin { get; set; }

        /// <summary>
        /// SupportByVersion Visio 11, 12, 14, 15, 16
        /// Get/Set
        /// </summary>
        [SupportByVersion("Visio", 11, 12, 14, 15, 16)]
        new Double RightMargin { get; set; }

        /// <summary>
        /// SupportByVersion Visio 11, 12, 14, 15, 16
        /// Get/Set
        /// </summary>
        [SupportByVersion("Visio", 11, 12, 14, 15, 16)]
        new Double TopMargin { get; set; }

        /// <summary>
        /// SupportByVersion Visio 11, 12, 14, 15, 16
        /// Get/Set
        /// </summary>
        [SupportByVersion("Visio", 11, 12, 14, 15, 16)]
        new Double BottomMargin { get; set; }

        /// <summary>
        /// SupportByVersion Visio 11, 12, 14, 15, 16
        /// Get/Set
        /// </summary>
        [SupportByVersion("Visio", 11, 12, 14, 15, 16)]
        [EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
        Int16 old_PrintLandscape { get; set; }

        /// <summary>
        /// SupportByVersion Visio 11, 12, 14, 15, 16
        /// Get/Set
        /// </summary>
        [SupportByVersion("Visio", 11, 12, 14, 15, 16)]
        [EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
        Int16 old_PrintCenteredH { get; set; }

        /// <summary>
        /// SupportByVersion Visio 11, 12, 14, 15, 16
        /// Get/Set
        /// </summary>
        [SupportByVersion("Visio", 11, 12, 14, 15, 16)]
        [EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
        Int16 old_PrintCenteredV { get; set; }

        /// <summary>
        /// SupportByVersion Visio 11, 12, 14, 15, 16
        /// Get/Set
        /// </summary>
        [SupportByVersion("Visio", 11, 12, 14, 15, 16)]
        Double PrintScale { get; set; }

        /// <summary>
        /// SupportByVersion Visio 11, 12, 14, 15, 16
        /// Get/Set
        /// </summary>
        [SupportByVersion("Visio", 11, 12, 14, 15, 16)]
        [EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
        Int16 old_PrintFitOnPages { get; set; }

        /// <summary>
        /// SupportByVersion Visio 11, 12, 14, 15, 16
        /// Get/Set
        /// </summary>
        [SupportByVersion("Visio", 11, 12, 14, 15, 16)]
        Int16 PrintPagesAcross { get; set; }

        /// <summary>
        /// SupportByVersion Visio 11, 12, 14, 15, 16
        /// Get/Set
        /// </summary>
        [SupportByVersion("Visio", 11, 12, 14, 15, 16)]
        Int16 PrintPagesDown { get; set; }

        /// <summary>
        /// SupportByVersion Visio 11, 12, 14, 15, 16
        /// Get/Set
        /// </summary>
        [SupportByVersion("Visio", 11, 12, 14, 15, 16)]
        string DefaultStyle { get; set; }

        /// <summary>
        /// SupportByVersion Visio 11, 12, 14, 15, 16
        /// Get/Set
        /// </summary>
        [SupportByVersion("Visio", 11, 12, 14, 15, 16)]
        string DefaultLineStyle { get; set; }

        /// <summary>
        /// SupportByVersion Visio 11, 12, 14, 15, 16
        /// Get/Set
        /// </summary>
        [SupportByVersion("Visio", 11, 12, 14, 15, 16)]
        string DefaultFillStyle { get; set; }

        /// <summary>
        /// SupportByVersion Visio 11, 12, 14, 15, 16
        /// Get/Set
        /// </summary>
        [SupportByVersion("Visio", 11, 12, 14, 15, 16)]
        string DefaultTextStyle { get; set; }

        /// <summary>
        /// SupportByVersion Visio 11, 12, 14, 15, 16
        /// Get
        /// </summary>
        [SupportByVersion("Visio", 11, 12, 14, 15, 16)]
        Int16 PersistsEvents { get; }

        /// <summary>
        /// SupportByVersion Visio 11, 12, 14, 15, 16
        /// Get
        /// Unknown COM Proxy
        /// </summary>
        [SupportByVersion("Visio", 11, 12, 14, 15, 16), ProxyResult]
        object VBProject { get; }

        /// <summary>
        /// SupportByVersion Visio 11, 12, 14, 15, 16
        /// Get
        /// </summary>
        /// <param name="unitsNameOrCode">object unitsNameOrCode</param>
        [SupportByVersion("Visio", 11, 12, 14, 15, 16)]
        [EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
        Double get_PaperWidth(object unitsNameOrCode);

        /// <summary>
        /// SupportByVersion Visio 11, 12, 14, 15, 16
        /// Alias for get_PaperWidth
        /// </summary>
        /// <param name="unitsNameOrCode">object unitsNameOrCode</param>
        [SupportByVersion("Visio", 11, 12, 14, 15, 16), Redirect("get_PaperWidth")]
        Double PaperWidth(object unitsNameOrCode);

        /// <summary>
        /// SupportByVersion Visio 11, 12, 14, 15, 16
        /// Get
        /// </summary>
        /// <param name="unitsNameOrCode">object unitsNameOrCode</param>
        [SupportByVersion("Visio", 11, 12, 14, 15, 16)]
        [EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
        Double get_PaperHeight(object unitsNameOrCode);

        /// <summary>
        /// SupportByVersion Visio 11, 12, 14, 15, 16
        /// Alias for get_PaperHeight
        /// </summary>
        /// <param name="unitsNameOrCode">object unitsNameOrCode</param>
        [SupportByVersion("Visio", 11, 12, 14, 15, 16), Redirect("get_PaperHeight")]
        Double PaperHeight(object unitsNameOrCode);

        /// <summary>
        /// SupportByVersion Visio 11, 12, 14, 15, 16
        /// Get/Set
        /// </summary>
        [SupportByVersion("Visio", 11, 12, 14, 15, 16)]
        [EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
        Int16 old_PaperSize { get; set; }

        /// <summary>
        /// SupportByVersion Visio 11, 12, 14, 15, 16
        /// Get
        /// </summary>
        [SupportByVersion("Visio", 11, 12, 14, 15, 16)]
        [EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
        string CodeName { get; }

        /// <summary>
        /// SupportByVersion Visio 11, 12, 14, 15, 16
        /// Get/Set
        /// </summary>
        [SupportByVersion("Visio", 11, 12, 14, 15, 16)]
        [EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
        Int16 old_Mode { get; set; }

        /// <summary>
        /// SupportByVersion Visio 11, 12, 14, 15, 16
        /// Get
        /// </summary>
        [SupportByVersion("Visio", 11, 12, 14, 15, 16)]
        [BaseResult]
        NetOffice.VisioApi.IVOLEObjects OLEObjects { get; }

        /// <summary>
        /// SupportByVersion Visio 11, 12, 14, 15, 16
        /// Get/Set
        /// </summary>
        [SupportByVersion("Visio", 11, 12, 14, 15, 16)]
        string Manager { get; set; }

        /// <summary>
        /// SupportByVersion Visio 11, 12, 14, 15, 16
        /// Get/Set
        /// </summary>
        [SupportByVersion("Visio", 11, 12, 14, 15, 16)]
        string Company { get; set; }

        /// <summary>
        /// SupportByVersion Visio 11, 12, 14, 15, 16
        /// Get/Set
        /// </summary>
        [SupportByVersion("Visio", 11, 12, 14, 15, 16)]
        string Category { get; set; }

        /// <summary>
        /// SupportByVersion Visio 11, 12, 14, 15, 16
        /// Get/Set
        /// </summary>
        [SupportByVersion("Visio", 11, 12, 14, 15, 16)]
        string HyperlinkBase { get; set; }

        /// <summary>
        /// SupportByVersion Visio 11, 12, 14, 15, 16
        /// Get
        /// </summary>
        [SupportByVersion("Visio", 11, 12, 14, 15, 16)]
        [BaseResult]
        NetOffice.VisioApi.IVShape DocumentSheet { get; }

        /// <summary>
        /// SupportByVersion Visio 11, 12, 14, 15, 16
        /// Get
        /// Unknown COM Proxy
        /// </summary>
        [SupportByVersion("Visio", 11, 12, 14, 15, 16), ProxyResult]
        object Container { get; }

        /// <summary>
        /// SupportByVersion Visio 11, 12, 14, 15, 16
        /// Get
        /// </summary>
        [SupportByVersion("Visio", 11, 12, 14, 15, 16)]
        string ClassID { get; }

        /// <summary>
        /// SupportByVersion Visio 11, 12, 14, 15, 16
        /// Get
        /// </summary>
        [SupportByVersion("Visio", 11, 12, 14, 15, 16)]
        string ProgID { get; }

        /// <summary>
        /// SupportByVersion Visio 11, 12, 14, 15, 16
        /// Get
        /// </summary>
        [SupportByVersion("Visio", 11, 12, 14, 15, 16)]
        [BaseResult]
        NetOffice.VisioApi.IVMasterShortcuts MasterShortcuts { get; }

        /// <summary>
        /// SupportByVersion Visio 11, 12, 14, 15, 16
        /// Get/Set
        /// </summary>
        [SupportByVersion("Visio", 11, 12, 14, 15, 16)]
        string AlternateNames { get; set; }

        /// <summary>
        /// SupportByVersion Visio 11, 12, 14, 15, 16
        /// Get
        /// </summary>
        [SupportByVersion("Visio", 11, 12, 14, 15, 16)]
        [BaseResult]
        NetOffice.VisioApi.IVShape GestureFormatSheet { get; }

        /// <summary>
        /// SupportByVersion Visio 11, 12, 14, 15, 16
        /// Get/Set
        /// </summary>
        [SupportByVersion("Visio", 11, 12, 14, 15, 16)]
        bool AutoRecover { get; set; }

        /// <summary>
        /// SupportByVersion Visio 11, 12, 14, 15, 16
        /// Get/Set
        /// </summary>
        [SupportByVersion("Visio", 11, 12, 14, 15, 16)]
        bool Saved { get; set; }

        /// <summary>
        /// SupportByVersion Visio 11, 12, 14, 15, 16
        /// Get/Set
        /// </summary>
        [SupportByVersion("Visio", 11, 12, 14, 15, 16)]
        NetOffice.VisioApi.Enums.VisDocVersions Version { get; set; }

        /// <summary>
        /// SupportByVersion Visio 11, 12, 14, 15, 16
        /// Get/Set
        /// </summary>
        [SupportByVersion("Visio", 11, 12, 14, 15, 16)]
        NetOffice.VisioApi.Enums.VisSavePreviewMode SavePreviewMode { get; set; }

        /// <summary>
        /// SupportByVersion Visio 11, 12, 14, 15, 16
        /// Get/Set
        /// </summary>
        [SupportByVersion("Visio", 11, 12, 14, 15, 16)]
        bool PrintLandscape { get; set; }

        /// <summary>
        /// SupportByVersion Visio 11, 12, 14, 15, 16
        /// Get/Set
        /// </summary>
        [SupportByVersion("Visio", 11, 12, 14, 15, 16)]
        bool PrintCenteredH { get; set; }

        /// <summary>
        /// SupportByVersion Visio 11, 12, 14, 15, 16
        /// Get/Set
        /// </summary>
        [SupportByVersion("Visio", 11, 12, 14, 15, 16)]
        bool PrintCenteredV { get; set; }

        /// <summary>
        /// SupportByVersion Visio 11, 12, 14, 15, 16
        /// Get/Set
        /// </summary>
        [SupportByVersion("Visio", 11, 12, 14, 15, 16)]
        bool PrintFitOnPages { get; set; }

        /// <summary>
        /// SupportByVersion Visio 11, 12, 14, 15, 16
        /// Get/Set
        /// </summary>
        [SupportByVersion("Visio", 11, 12, 14, 15, 16)]
        NetOffice.VisioApi.Enums.VisPaperSizes PaperSize { get; set; }

        /// <summary>
        /// SupportByVersion Visio 11, 12, 14, 15, 16
        /// Get/Set
        /// </summary>
        [SupportByVersion("Visio", 11, 12, 14, 15, 16)]
        NetOffice.VisioApi.Enums.VisDocModeArgs Mode { get; set; }

        /// <summary>
        /// SupportByVersion Visio 11, 12, 14, 15, 16
        /// Get/Set
        /// </summary>
        [SupportByVersion("Visio", 11, 12, 14, 15, 16)]
        bool SnapEnabled { get; set; }

        /// <summary>
        /// SupportByVersion Visio 11, 12, 14, 15, 16
        /// Get/Set
        /// </summary>
        [SupportByVersion("Visio", 11, 12, 14, 15, 16)]
        NetOffice.VisioApi.Enums.VisSnapSettings SnapSettings { get; set; }

        /// <summary>
        /// SupportByVersion Visio 11, 12, 14, 15, 16
        /// Get/Set
        /// </summary>
        [SupportByVersion("Visio", 11, 12, 14, 15, 16)]
        NetOffice.VisioApi.Enums.VisSnapExtensions SnapExtensions { get; set; }

        /// <summary>
        /// SupportByVersion Visio 11, 12, 14, 15, 16
        /// Get/Set
        /// </summary>
        [SupportByVersion("Visio", 11, 12, 14, 15, 16)]
        Double[] SnapAngles { get; set; }

        /// <summary>
        /// SupportByVersion Visio 11, 12, 14, 15, 16
        /// Get/Set
        /// </summary>
        [SupportByVersion("Visio", 11, 12, 14, 15, 16)]
        bool GlueEnabled { get; set; }

        /// <summary>
        /// SupportByVersion Visio 11, 12, 14, 15, 16
        /// Get/Set
        /// </summary>
        [SupportByVersion("Visio", 11, 12, 14, 15, 16)]
        NetOffice.VisioApi.Enums.VisGlueSettings GlueSettings { get; set; }

        /// <summary>
        /// SupportByVersion Visio 11, 12, 14, 15, 16
        /// Get/Set
        /// </summary>
        [SupportByVersion("Visio", 11, 12, 14, 15, 16)]
        bool DynamicGridEnabled { get; set; }

        /// <summary>
        /// SupportByVersion Visio 11, 12, 14, 15, 16
        /// Get/Set
        /// </summary>
        [SupportByVersion("Visio", 11, 12, 14, 15, 16)]
        string DefaultGuideStyle { get; set; }

        /// <summary>
        /// SupportByVersion Visio 11, 12, 14, 15, 16
        /// Get/Set
        /// </summary>
        [SupportByVersion("Visio", 11, 12, 14, 15, 16)]
        new NetOffice.VisioApi.Enums.VisProtection Protection { get; set; }

        /// <summary>
        /// SupportByVersion Visio 11, 12, 14, 15, 16
        /// Get/Set
        /// </summary>
        [SupportByVersion("Visio", 11, 12, 14, 15, 16)]
        string Printer { get; set; }

        /// <summary>
        /// SupportByVersion Visio 11, 12, 14, 15, 16
        /// Get/Set
        /// </summary>
        [SupportByVersion("Visio", 11, 12, 14, 15, 16)]
        [EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
        Int32 PrintCopies { get; set; }

        /// <summary>
        /// SupportByVersion Visio 11, 12, 14, 15, 16
        /// Get/Set
        /// </summary>
        [SupportByVersion("Visio", 11, 12, 14, 15, 16)]
        string HeaderLeft { get; set; }

        /// <summary>
        /// SupportByVersion Visio 11, 12, 14, 15, 16
        /// Get/Set
        /// </summary>
        [SupportByVersion("Visio", 11, 12, 14, 15, 16)]
        string HeaderCenter { get; set; }

        /// <summary>
        /// SupportByVersion Visio 11, 12, 14, 15, 16
        /// Get/Set
        /// </summary>
        [SupportByVersion("Visio", 11, 12, 14, 15, 16)]
        string HeaderRight { get; set; }

        /// <summary>
        /// SupportByVersion Visio 11, 12, 14, 15, 16
        /// Get/Set
        /// </summary>
        [SupportByVersion("Visio", 11, 12, 14, 15, 16)]
        new Double HeaderMargin { get; set; }

        /// <summary>
        /// SupportByVersion Visio 11, 12, 14, 15, 16
        /// Get/Set
        /// </summary>
        [SupportByVersion("Visio", 11, 12, 14, 15, 16)]
        string FooterLeft { get; set; }

        /// <summary>
        /// SupportByVersion Visio 11, 12, 14, 15, 16
        /// Get/Set
        /// </summary>
        [SupportByVersion("Visio", 11, 12, 14, 15, 16)]
        string FooterCenter { get; set; }

        /// <summary>
        /// SupportByVersion Visio 11, 12, 14, 15, 16
        /// Get/Set
        /// </summary>
        [SupportByVersion("Visio", 11, 12, 14, 15, 16)]
        string FooterRight { get; set; }

        /// <summary>
        /// SupportByVersion Visio 11, 12, 14, 15, 16
        /// Get/Set
        /// </summary>
        [SupportByVersion("Visio", 11, 12, 14, 15, 16)]
        new Double FooterMargin { get; set; }

        /// <summary>
        /// SupportByVersion Visio 11, 12, 14, 15, 16
        /// Get/Set
        /// </summary>
        [SupportByVersion("Visio", 11, 12, 14, 15, 16), NativeResult]
        stdole.Font HeaderFooterFont { get; set; }

        /// <summary>
        /// SupportByVersion Visio 11, 12, 14, 15, 16
        /// Get/Set
        /// </summary>
        [SupportByVersion("Visio", 11, 12, 14, 15, 16)]
        Int32 HeaderFooterColor { get; set; }

        /// <summary>
        /// SupportByVersion Visio 11, 12, 14, 15, 16
        /// Get/Set
        /// </summary>
        [SupportByVersion("Visio", 11, 12, 14, 15, 16)]
        [EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
        new string Password { get; set; }

        /// <summary>
        /// SupportByVersion Visio 11, 12, 14, 15, 16
        /// Get/Set
        /// </summary>
        [SupportByVersion("Visio", 11, 12, 14, 15, 16), NativeResult]
        stdole.Picture PreviewPicture { get; set; }

        /// <summary>
        /// SupportByVersion Visio 11, 12, 14, 15, 16
        /// Get
        /// </summary>
        [SupportByVersion("Visio", 11, 12, 14, 15, 16)]
        Int32 BuildNumberCreated { get; }

        /// <summary>
        /// SupportByVersion Visio 11, 12, 14, 15, 16
        /// Get
        /// </summary>
        [SupportByVersion("Visio", 11, 12, 14, 15, 16)]
        Int32 BuildNumberEdited { get; }

        /// <summary>
        /// SupportByVersion Visio 11, 12, 14, 15, 16
        /// Get
        /// </summary>
        [SupportByVersion("Visio", 11, 12, 14, 15, 16)]
        DateTime TimeCreated { get; }

        /// <summary>
        /// SupportByVersion Visio 11, 12, 14, 15, 16
        /// Get
        /// </summary>
        [SupportByVersion("Visio", 11, 12, 14, 15, 16)]
        DateTime Time { get; }

        /// <summary>
        /// SupportByVersion Visio 11, 12, 14, 15, 16
        /// Get
        /// </summary>
        [SupportByVersion("Visio", 11, 12, 14, 15, 16)]
        DateTime TimeEdited { get; }

        /// <summary>
        /// SupportByVersion Visio 11, 12, 14, 15, 16
        /// Get
        /// </summary>
        [SupportByVersion("Visio", 11, 12, 14, 15, 16)]
        DateTime TimePrinted { get; }

        /// <summary>
        /// SupportByVersion Visio 11, 12, 14, 15, 16
        /// Get
        /// </summary>
        [SupportByVersion("Visio", 11, 12, 14, 15, 16)]
        DateTime TimeSaved { get; }

        /// <summary>
        /// SupportByVersion Visio 11, 12, 14, 15, 16
        /// Get
        /// </summary>
        [SupportByVersion("Visio", 11, 12, 14, 15, 16)]
        bool ContainsWorkspace { get; }

        /// <summary>
        /// SupportByVersion Visio 11, 12, 14, 15, 16
        /// Get
        /// </summary>
        [SupportByVersion("Visio", 11, 12, 14, 15, 16)]
        object[] EmailRoutingData { get; }

		/// <summary>
		/// SupportByVersion Visio 11, 12, 14, 15, 16
		/// Get
		/// </summary>
		[SupportByVersion("Visio", 11, 12, 14, 15, 16)]
        byte[] VBProjectData { get; }

        /// <summary>
        /// SupportByVersion Visio 11, 12, 14, 15, 16
        /// Get
        /// </summary>
        [SupportByVersion("Visio", 11, 12, 14, 15, 16)]
        Int32 SolutionXMLElementCount { get; }

        /// <summary>
        /// SupportByVersion Visio 11, 12, 14, 15, 16
        /// Get
        /// </summary>
        /// <param name="index">Int32 index</param>
        [SupportByVersion("Visio", 11, 12, 14, 15, 16)]
        [EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
        string get_SolutionXMLElementName(Int32 index);

        /// <summary>
        /// SupportByVersion Visio 11, 12, 14, 15, 16
        /// Alias for get_SolutionXMLElementName
        /// </summary>
        /// <param name="index">Int32 index</param>
        [SupportByVersion("Visio", 11, 12, 14, 15, 16), Redirect("get_SolutionXMLElementName")]
        string SolutionXMLElementName(Int32 index);

        /// <summary>
        /// SupportByVersion Visio 11, 12, 14, 15, 16
        /// Get
        /// </summary>
        /// <param name="elementName">string elementName</param>
        [SupportByVersion("Visio", 11, 12, 14, 15, 16)]
        [EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
        bool get_SolutionXMLElementExists(string elementName);

        /// <summary>
        /// SupportByVersion Visio 11, 12, 14, 15, 16
        /// Alias for get_SolutionXMLElementExists
        /// </summary>
        /// <param name="elementName">string elementName</param>
        [SupportByVersion("Visio", 11, 12, 14, 15, 16), Redirect("get_SolutionXMLElementExists")]
        bool SolutionXMLElementExists(string elementName);

        /// <summary>
        /// SupportByVersion Visio 11, 12, 14, 15, 16
        /// Get/Set
        /// </summary>
        /// <param name="elementName">string elementName</param>
        [SupportByVersion("Visio", 11, 12, 14, 15, 16)]
        [EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
        string get_SolutionXMLElement(string elementName);

        /// <summary>
        /// SupportByVersion Visio 11, 12, 14, 15, 16
        /// Get/Set
        /// </summary>
        /// <param name="elementName">string elementName</param>
        /// <param name="value">string value</param>
        [SupportByVersion("Visio", 11, 12, 14, 15, 16)]
        [EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
        void set_SolutionXMLElement(string elementName, string value);

        /// <summary>
        /// SupportByVersion Visio 11, 12, 14, 15, 16
        /// Alias for get_SolutionXMLElement
        /// </summary>
        /// <param name="elementName">string elementName</param>
        [SupportByVersion("Visio", 11, 12, 14, 15, 16), Redirect("get_SolutionXMLElement")]
        string SolutionXMLElement(string elementName);

        /// <summary>
        /// SupportByVersion Visio 11, 12, 14, 15, 16
        /// Get
        /// </summary>
        [SupportByVersion("Visio", 11, 12, 14, 15, 16)]
        Int32 FullBuildNumberCreated { get; }

        /// <summary>
        /// SupportByVersion Visio 11, 12, 14, 15, 16
        /// Get
        /// </summary>
        [SupportByVersion("Visio", 11, 12, 14, 15, 16)]
        Int32 FullBuildNumberEdited { get; }

        /// <summary>
        /// SupportByVersion Visio 11, 12, 14, 15, 16
        /// Get
        /// </summary>
        [SupportByVersion("Visio", 11, 12, 14, 15, 16)]
        Int32 ID { get; }

        /// <summary>
        /// SupportByVersion Visio 11, 12, 14, 15, 16
        /// Get
        /// </summary>
        [SupportByVersion("Visio", 11, 12, 14, 15, 16)]
        bool MacrosEnabled { get; }

        /// <summary>
        /// SupportByVersion Visio 11, 12, 14, 15, 16
        /// Get/Set
        /// </summary>
        [SupportByVersion("Visio", 11, 12, 14, 15, 16)]
        NetOffice.VisioApi.Enums.VisZoomBehavior ZoomBehavior { get; set; }

        /// <summary>
        /// SupportByVersion Visio 11, 12, 14, 15, 16
        /// Get
        /// </summary>
        [SupportByVersion("Visio", 11, 12, 14, 15, 16)]
        NetOffice.VisioApi.Enums.VisDocumentTypes Type { get; }

        /// <summary>
        /// SupportByVersion Visio 11, 12, 14, 15, 16
        /// Get/Set
        /// </summary>
        [SupportByVersion("Visio", 11, 12, 14, 15, 16)]
        Int32 Language { get; set; }

        /// <summary>
        /// SupportByVersion Visio 11, 12, 14, 15, 16
        /// Get/Set
        /// </summary>
        [SupportByVersion("Visio", 11, 12, 14, 15, 16)]
        bool RemovePersonalInformation { get; set; }

        /// <summary>
        /// SupportByVersion Visio 11, 12, 14, 15, 16
        /// Get/Set
        /// </summary>
        [SupportByVersion("Visio", 11, 12, 14, 15, 16)]
        bool UndoEnabled { get; set; }

        /// <summary>
        /// SupportByVersion Visio 11, 12, 14, 15, 16
        /// Get
        /// Unknown COM Proxy
        /// </summary>
        [SupportByVersion("Visio", 11, 12, 14, 15, 16), ProxyResult]
        object SharedWorkspace { get; }

        /// <summary>
        /// SupportByVersion Visio 11, 12, 14, 15, 16
        /// Get
        /// Unknown COM Proxy
        /// </summary>
        [SupportByVersion("Visio", 11, 12, 14, 15, 16), ProxyResult]
        object Sync { get; }

        /// <summary>
        /// SupportByVersion Visio 12, 14, 15, 16
        /// Get
        /// </summary>
        [SupportByVersion("Visio", 12, 14, 15, 16)]
        [BaseResult]
        NetOffice.VisioApi.IVDataRecordsets DataRecordsets { get; }

        /// <summary>
        /// SupportByVersion Visio 12, 14, 15, 16
        /// Get/Set
        /// </summary>
        [SupportByVersion("Visio", 12, 14, 15, 16)]
        bool ContainsWorkspaceEx { get; set; }

        /// <summary>
        /// SupportByVersion Visio 12, 14, 15, 16
        /// Get/Set
        /// </summary>
        [SupportByVersion("Visio", 12, 14, 15, 16)]
        string DefaultSavePath { get; set; }

        /// <summary>
        /// SupportByVersion Visio 14, 15, 16
        /// Get/Set
        /// </summary>
        [SupportByVersion("Visio", 14, 15, 16)]
        string CustomUI { get; set; }

        /// <summary>
        /// SupportByVersion Visio 14, 15, 16
        /// Get/Set
        /// </summary>
        [SupportByVersion("Visio", 14, 15, 16)]
        string UserCustomUI { get; set; }

        /// <summary>
        /// SupportByVersion Visio 14, 15, 16
        /// Get
        /// </summary>
        [SupportByVersion("Visio", 14, 15, 16)]
        [BaseResult]
        NetOffice.VisioApi.IVServerPublishOptions ServerPublishOptions { get; }

        /// <summary>
        /// SupportByVersion Visio 14, 15, 16
        /// Get
        /// </summary>
        [SupportByVersion("Visio", 14, 15, 16)]
        [BaseResult]
        NetOffice.VisioApi.IVValidation Validation { get; }

        /// <summary>
        /// SupportByVersion Visio 14, 15, 16
        /// Get/Set
        /// </summary>
        [SupportByVersion("Visio", 14, 15, 16)]
        Int32 DiagramServicesEnabled { get; set; }

        /// <summary>
        /// SupportByVersion Visio 15,16
        /// Get
        /// </summary>
        [SupportByVersion("Visio", 15, 16)]
        bool CompatibilityMode { get; }

        /// <summary>
        /// SupportByVersion Visio 15,16
        /// Get
        /// </summary>
        [SupportByVersion("Visio", 15, 16)]
        [BaseResult]
        NetOffice.VisioApi.IVComments Comments { get; }

        #endregion

        #region Methods

        /// <summary>
        /// SupportByVersion Visio 11, 12, 14, 15, 16
        /// </summary>
        /// <param name="objectToDrop">object objectToDrop</param>
        /// <param name="xPos">Int16 xPos</param>
        /// <param name="yPos">Int16 yPos</param>
        [SupportByVersion("Visio", 11, 12, 14, 15, 16)]
        [BaseResult]
        NetOffice.VisioApi.IVMaster Drop(object objectToDrop, Int16 xPos, Int16 yPos);

        /// <summary>
        /// SupportByVersion Visio 11, 12, 14, 15, 16
        /// </summary>
        [SupportByVersion("Visio", 11, 12, 14, 15, 16)]
        Int16 Save();

        /// <summary>
        /// SupportByVersion Visio 11, 12, 14, 15, 16
        /// </summary>
        /// <param name="fileName">string fileName</param>
        [SupportByVersion("Visio", 11, 12, 14, 15, 16)]
        Int16 SaveAs(string fileName);

        /// <summary>
        /// SupportByVersion Visio 11, 12, 14, 15, 16
        /// </summary>
        [SupportByVersion("Visio", 11, 12, 14, 15, 16)]
        void Print();

        /// <summary>
        /// SupportByVersion Visio 11, 12, 14, 15, 16
        /// </summary>
        [SupportByVersion("Visio", 11, 12, 14, 15, 16)]
        void Close();

        /// <summary>
        /// SupportByVersion Visio 11, 12, 14, 15, 16
        /// </summary>
        /// <param name="menusObject">NetOffice.VisioApi.IVUIObject menusObject</param>
        [SupportByVersion("Visio", 11, 12, 14, 15, 16)]
        void SetCustomMenus(NetOffice.VisioApi.IVUIObject menusObject);

        /// <summary>
        /// SupportByVersion Visio 11, 12, 14, 15, 16
        /// </summary>
        [SupportByVersion("Visio", 11, 12, 14, 15, 16)]
        void ClearCustomMenus();

        /// <summary>
        /// SupportByVersion Visio 11, 12, 14, 15, 16
        /// </summary>
        /// <param name="toolbarsObject">NetOffice.VisioApi.IVUIObject toolbarsObject</param>
        [SupportByVersion("Visio", 11, 12, 14, 15, 16)]
        void SetCustomToolbars(NetOffice.VisioApi.IVUIObject toolbarsObject);

        /// <summary>
        /// SupportByVersion Visio 11, 12, 14, 15, 16
        /// </summary>
        [SupportByVersion("Visio", 11, 12, 14, 15, 16)]
        void ClearCustomToolbars();

        /// <summary>
        /// SupportByVersion Visio 11, 12, 14, 15, 16
        /// </summary>
        /// <param name="fileName">string fileName</param>
        /// <param name="saveFlags">Int16 saveFlags</param>
        [SupportByVersion("Visio", 11, 12, 14, 15, 16)]
        void SaveAsEx(string fileName, Int16 saveFlags);

        /// <summary>
        /// SupportByVersion Visio 11, 12, 14, 15, 16
        /// </summary>
        /// <param name="iD">Int16 iD</param>
        /// <param name="fileName">string fileName</param>
        [EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
        [SupportByVersion("Visio", 11, 12, 14, 15, 16)]
        void GetIcon(Int16 iD, string fileName);

        /// <summary>
        /// SupportByVersion Visio 11, 12, 14, 15, 16
        /// </summary>
        /// <param name="iD">Int16 iD</param>
        /// <param name="index">Int16 index</param>
        /// <param name="fileName">string fileName</param>
        [EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
        [SupportByVersion("Visio", 11, 12, 14, 15, 16)]
        void SetIcon(Int16 iD, Int16 index, string fileName);

        /// <summary>
        /// SupportByVersion Visio 11, 12, 14, 15, 16
        /// </summary>
        [SupportByVersion("Visio", 11, 12, 14, 15, 16)]
        [BaseResult]
        NetOffice.VisioApi.IVWindow OpenStencilWindow();

        /// <summary>
        /// SupportByVersion Visio 11, 12, 14, 15, 16
        /// </summary>
        /// <param name="line">string line</param>
        [SupportByVersion("Visio", 11, 12, 14, 15, 16)]
        void ParseLine(string line);

        /// <summary>
        /// SupportByVersion Visio 11, 12, 14, 15, 16
        /// </summary>
        /// <param name="line">string line</param>
        [SupportByVersion("Visio", 11, 12, 14, 15, 16)]
        void ExecuteLine(string line);

        /// <summary>
        /// SupportByVersion Visio 11, 12, 14, 15, 16
        /// </summary>
        /// <param name="target">string target</param>
        /// <param name="location">string location</param>
        [EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
        [SupportByVersion("Visio", 11, 12, 14, 15, 16)]
        void FollowHyperlink45(string target, string location);

        /// <summary>
        /// SupportByVersion Visio 11, 12, 14, 15, 16
        /// </summary>
        /// <param name="address">string address</param>
        /// <param name="subAddress">string subAddress</param>
        /// <param name="extraInfo">optional object extraInfo</param>
        /// <param name="frame">optional object frame</param>
        /// <param name="newWindow">optional object newWindow</param>
        /// <param name="res1">optional object res1</param>
        /// <param name="res2">optional object res2</param>
        /// <param name="res3">optional object res3</param>
        [SupportByVersion("Visio", 11, 12, 14, 15, 16)]
        void FollowHyperlink(string address, string subAddress, object extraInfo, object frame, object newWindow, object res1, object res2, object res3);

        /// <summary>
        /// SupportByVersion Visio 11, 12, 14, 15, 16
        /// </summary>
        /// <param name="address">string address</param>
        /// <param name="subAddress">string subAddress</param>
        [CustomMethod]
        [SupportByVersion("Visio", 11, 12, 14, 15, 16)]
        void FollowHyperlink(string address, string subAddress);

        /// <summary>
        /// SupportByVersion Visio 11, 12, 14, 15, 16
        /// </summary>
        /// <param name="address">string address</param>
        /// <param name="subAddress">string subAddress</param>
        /// <param name="extraInfo">optional object extraInfo</param>
        [CustomMethod]
        [SupportByVersion("Visio", 11, 12, 14, 15, 16)]
        void FollowHyperlink(string address, string subAddress, object extraInfo);

        /// <summary>
        /// SupportByVersion Visio 11, 12, 14, 15, 16
        /// </summary>
        /// <param name="address">string address</param>
        /// <param name="subAddress">string subAddress</param>
        /// <param name="extraInfo">optional object extraInfo</param>
        /// <param name="frame">optional object frame</param>
        [CustomMethod]
        [SupportByVersion("Visio", 11, 12, 14, 15, 16)]
        void FollowHyperlink(string address, string subAddress, object extraInfo, object frame);

        /// <summary>
        /// SupportByVersion Visio 11, 12, 14, 15, 16
        /// </summary>
        /// <param name="address">string address</param>
        /// <param name="subAddress">string subAddress</param>
        /// <param name="extraInfo">optional object extraInfo</param>
        /// <param name="frame">optional object frame</param>
        /// <param name="newWindow">optional object newWindow</param>
        [CustomMethod]
        [SupportByVersion("Visio", 11, 12, 14, 15, 16)]
        void FollowHyperlink(string address, string subAddress, object extraInfo, object frame, object newWindow);

        /// <summary>
        /// SupportByVersion Visio 11, 12, 14, 15, 16
        /// </summary>
        /// <param name="address">string address</param>
        /// <param name="subAddress">string subAddress</param>
        /// <param name="extraInfo">optional object extraInfo</param>
        /// <param name="frame">optional object frame</param>
        /// <param name="newWindow">optional object newWindow</param>
        /// <param name="res1">optional object res1</param>
        [CustomMethod]
        [SupportByVersion("Visio", 11, 12, 14, 15, 16)]
        void FollowHyperlink(string address, string subAddress, object extraInfo, object frame, object newWindow, object res1);

        /// <summary>
        /// SupportByVersion Visio 11, 12, 14, 15, 16
        /// </summary>
        /// <param name="address">string address</param>
        /// <param name="subAddress">string subAddress</param>
        /// <param name="extraInfo">optional object extraInfo</param>
        /// <param name="frame">optional object frame</param>
        /// <param name="newWindow">optional object newWindow</param>
        /// <param name="res1">optional object res1</param>
        /// <param name="res2">optional object res2</param>
        [CustomMethod]
        [SupportByVersion("Visio", 11, 12, 14, 15, 16)]
        void FollowHyperlink(string address, string subAddress, object extraInfo, object frame, object newWindow, object res1, object res2);

        /// <summary>
        /// SupportByVersion Visio 11, 12, 14, 15, 16
        /// </summary>
        [SupportByVersion("Visio", 11, 12, 14, 15, 16)]
        void ClearGestureFormatSheet();

        /// <summary>
        /// SupportByVersion Visio 11, 12, 14, 15, 16
        /// </summary>
        /// <param name="nTargets">optional object nTargets</param>
        /// <param name="nActions">optional object nActions</param>
        /// <param name="nAlerts">optional object nAlerts</param>
        /// <param name="nFixes">optional object nFixes</param>
        /// <param name="bStopOnError">optional object bStopOnError</param>
        /// <param name="bLogFileName">optional object bLogFileName</param>
        /// <param name="nReserved">optional object nReserved</param>
        [SupportByVersion("Visio", 11, 12, 14, 15, 16)]
        void Clean(object nTargets, object nActions, object nAlerts, object nFixes, object bStopOnError, object bLogFileName, object nReserved);

        /// <summary>
        /// SupportByVersion Visio 11, 12, 14, 15, 16
        /// </summary>
        [CustomMethod]
        [SupportByVersion("Visio", 11, 12, 14, 15, 16)]
        void Clean();

        /// <summary>
        /// SupportByVersion Visio 11, 12, 14, 15, 16
        /// </summary>
        /// <param name="nTargets">optional object nTargets</param>
        [CustomMethod]
        [SupportByVersion("Visio", 11, 12, 14, 15, 16)]
        void Clean(object nTargets);

        /// <summary>
        /// SupportByVersion Visio 11, 12, 14, 15, 16
        /// </summary>
        /// <param name="nTargets">optional object nTargets</param>
        /// <param name="nActions">optional object nActions</param>
        [CustomMethod]
        [SupportByVersion("Visio", 11, 12, 14, 15, 16)]
        void Clean(object nTargets, object nActions);

        /// <summary>
        /// SupportByVersion Visio 11, 12, 14, 15, 16
        /// </summary>
        /// <param name="nTargets">optional object nTargets</param>
        /// <param name="nActions">optional object nActions</param>
        /// <param name="nAlerts">optional object nAlerts</param>
        [CustomMethod]
        [SupportByVersion("Visio", 11, 12, 14, 15, 16)]
        void Clean(object nTargets, object nActions, object nAlerts);

        /// <summary>
        /// SupportByVersion Visio 11, 12, 14, 15, 16
        /// </summary>
        /// <param name="nTargets">optional object nTargets</param>
        /// <param name="nActions">optional object nActions</param>
        /// <param name="nAlerts">optional object nAlerts</param>
        /// <param name="nFixes">optional object nFixes</param>
        [CustomMethod]
        [SupportByVersion("Visio", 11, 12, 14, 15, 16)]
        void Clean(object nTargets, object nActions, object nAlerts, object nFixes);

        /// <summary>
        /// SupportByVersion Visio 11, 12, 14, 15, 16
        /// </summary>
        /// <param name="nTargets">optional object nTargets</param>
        /// <param name="nActions">optional object nActions</param>
        /// <param name="nAlerts">optional object nAlerts</param>
        /// <param name="nFixes">optional object nFixes</param>
        /// <param name="bStopOnError">optional object bStopOnError</param>
        [CustomMethod]
        [SupportByVersion("Visio", 11, 12, 14, 15, 16)]
        void Clean(object nTargets, object nActions, object nAlerts, object nFixes, object bStopOnError);

        /// <summary>
        /// SupportByVersion Visio 11, 12, 14, 15, 16
        /// </summary>
        /// <param name="nTargets">optional object nTargets</param>
        /// <param name="nActions">optional object nActions</param>
        /// <param name="nAlerts">optional object nAlerts</param>
        /// <param name="nFixes">optional object nFixes</param>
        /// <param name="bStopOnError">optional object bStopOnError</param>
        /// <param name="bLogFileName">optional object bLogFileName</param>
        [CustomMethod]
        [SupportByVersion("Visio", 11, 12, 14, 15, 16)]
        void Clean(object nTargets, object nActions, object nAlerts, object nFixes, object bStopOnError, object bLogFileName);

        /// <summary>
        /// SupportByVersion Visio 11, 12, 14, 15, 16
        /// </summary>
        /// <param name="pSourceDoc">NetOffice.VisioApi.IVDocument pSourceDoc</param>
        [SupportByVersion("Visio", 11, 12, 14, 15, 16)]
        void CopyPreviewPicture(NetOffice.VisioApi.IVDocument pSourceDoc);

        /// <summary>
        /// SupportByVersion Visio 11, 12, 14, 15, 16
        /// </summary>
        /// <param name="elementName">string elementName</param>
        [SupportByVersion("Visio", 11, 12, 14, 15, 16)]
        void DeleteSolutionXMLElement(string elementName);

        /// <summary>
        /// SupportByVersion Visio 11, 12, 14, 15, 16
        /// </summary>
        [SupportByVersion("Visio", 11, 12, 14, 15, 16)]
        bool CanCheckIn();

        /// <summary>
        /// SupportByVersion Visio 11, 12, 14, 15, 16
        /// </summary>
        /// <param name="saveChanges">optional bool SaveChanges = true</param>
        /// <param name="comments">optional object comments</param>
        /// <param name="makePublic">optional bool MakePublic = false</param>
        [SupportByVersion("Visio", 11, 12, 14, 15, 16)]
        void CheckIn(object saveChanges, object comments, object makePublic);

        /// <summary>
        /// SupportByVersion Visio 11, 12, 14, 15, 16
        /// </summary>
        [CustomMethod]
        [SupportByVersion("Visio", 11, 12, 14, 15, 16)]
        void CheckIn();

        /// <summary>
        /// SupportByVersion Visio 11, 12, 14, 15, 16
        /// </summary>
        /// <param name="saveChanges">optional bool SaveChanges = true</param>
        [CustomMethod]
        [SupportByVersion("Visio", 11, 12, 14, 15, 16)]
        void CheckIn(object saveChanges);

        /// <summary>
        /// SupportByVersion Visio 11, 12, 14, 15, 16
        /// </summary>
        /// <param name="saveChanges">optional bool SaveChanges = true</param>
        /// <param name="comments">optional object comments</param>
        [CustomMethod]
        [SupportByVersion("Visio", 11, 12, 14, 15, 16)]
        void CheckIn(object saveChanges, object comments);

        /// <summary>
        /// SupportByVersion Visio 11, 12, 14, 15, 16
        /// </summary>
        /// <param name="printRange">NetOffice.VisioApi.Enums.VisPrintOutRange printRange</param>
        /// <param name="fromPage">optional Int32 FromPage = 1</param>
        /// <param name="toPage">optional Int32 ToPage = -1</param>
        /// <param name="scaleCurrentViewToPaper">optional bool ScaleCurrentViewToPaper = false</param>
        /// <param name="printerName">optional string PrinterName = </param>
        /// <param name="printToFile">optional bool PrintToFile = false</param>
        /// <param name="outputFileName">optional string OutputFileName = </param>
        /// <param name="copies">optional Int32 Copies = 1</param>
        /// <param name="collate">optional bool Collate = false</param>
        /// <param name="colorAsBlack">optional bool ColorAsBlack = false</param>
        [SupportByVersion("Visio", 11, 12, 14, 15, 16)]
        void PrintOut(NetOffice.VisioApi.Enums.VisPrintOutRange printRange, object fromPage, object toPage, object scaleCurrentViewToPaper, object printerName, object printToFile, object outputFileName, object copies, object collate, object colorAsBlack);

        /// <summary>
        /// SupportByVersion Visio 11, 12, 14, 15, 16
        /// </summary>
        /// <param name="printRange">NetOffice.VisioApi.Enums.VisPrintOutRange printRange</param>
        [CustomMethod]
        [SupportByVersion("Visio", 11, 12, 14, 15, 16)]
        void PrintOut(NetOffice.VisioApi.Enums.VisPrintOutRange printRange);

        /// <summary>
        /// SupportByVersion Visio 11, 12, 14, 15, 16
        /// </summary>
        /// <param name="printRange">NetOffice.VisioApi.Enums.VisPrintOutRange printRange</param>
        /// <param name="fromPage">optional Int32 FromPage = 1</param>
        [CustomMethod]
        [SupportByVersion("Visio", 11, 12, 14, 15, 16)]
        void PrintOut(NetOffice.VisioApi.Enums.VisPrintOutRange printRange, object fromPage);

        /// <summary>
        /// SupportByVersion Visio 11, 12, 14, 15, 16
        /// </summary>
        /// <param name="printRange">NetOffice.VisioApi.Enums.VisPrintOutRange printRange</param>
        /// <param name="fromPage">optional Int32 FromPage = 1</param>
        /// <param name="toPage">optional Int32 ToPage = -1</param>
        [CustomMethod]
        [SupportByVersion("Visio", 11, 12, 14, 15, 16)]
        void PrintOut(NetOffice.VisioApi.Enums.VisPrintOutRange printRange, object fromPage, object toPage);

        /// <summary>
        /// SupportByVersion Visio 11, 12, 14, 15, 16
        /// </summary>
        /// <param name="printRange">NetOffice.VisioApi.Enums.VisPrintOutRange printRange</param>
        /// <param name="fromPage">optional Int32 FromPage = 1</param>
        /// <param name="toPage">optional Int32 ToPage = -1</param>
        /// <param name="scaleCurrentViewToPaper">optional bool ScaleCurrentViewToPaper = false</param>
        [CustomMethod]
        [SupportByVersion("Visio", 11, 12, 14, 15, 16)]
        void PrintOut(NetOffice.VisioApi.Enums.VisPrintOutRange printRange, object fromPage, object toPage, object scaleCurrentViewToPaper);

        /// <summary>
        /// SupportByVersion Visio 11, 12, 14, 15, 16
        /// </summary>
        /// <param name="printRange">NetOffice.VisioApi.Enums.VisPrintOutRange printRange</param>
        /// <param name="fromPage">optional Int32 FromPage = 1</param>
        /// <param name="toPage">optional Int32 ToPage = -1</param>
        /// <param name="scaleCurrentViewToPaper">optional bool ScaleCurrentViewToPaper = false</param>
        /// <param name="printerName">optional string PrinterName = </param>
        [CustomMethod]
        [SupportByVersion("Visio", 11, 12, 14, 15, 16)]
        void PrintOut(NetOffice.VisioApi.Enums.VisPrintOutRange printRange, object fromPage, object toPage, object scaleCurrentViewToPaper, object printerName);

        /// <summary>
        /// SupportByVersion Visio 11, 12, 14, 15, 16
        /// </summary>
        /// <param name="printRange">NetOffice.VisioApi.Enums.VisPrintOutRange printRange</param>
        /// <param name="fromPage">optional Int32 FromPage = 1</param>
        /// <param name="toPage">optional Int32 ToPage = -1</param>
        /// <param name="scaleCurrentViewToPaper">optional bool ScaleCurrentViewToPaper = false</param>
        /// <param name="printerName">optional string PrinterName = </param>
        /// <param name="printToFile">optional bool PrintToFile = false</param>
        [CustomMethod]
        [SupportByVersion("Visio", 11, 12, 14, 15, 16)]
        void PrintOut(NetOffice.VisioApi.Enums.VisPrintOutRange printRange, object fromPage, object toPage, object scaleCurrentViewToPaper, object printerName, object printToFile);

        /// <summary>
        /// SupportByVersion Visio 11, 12, 14, 15, 16
        /// </summary>
        /// <param name="printRange">NetOffice.VisioApi.Enums.VisPrintOutRange printRange</param>
        /// <param name="fromPage">optional Int32 FromPage = 1</param>
        /// <param name="toPage">optional Int32 ToPage = -1</param>
        /// <param name="scaleCurrentViewToPaper">optional bool ScaleCurrentViewToPaper = false</param>
        /// <param name="printerName">optional string PrinterName = </param>
        /// <param name="printToFile">optional bool PrintToFile = false</param>
        /// <param name="outputFileName">optional string OutputFileName = </param>
        [CustomMethod]
        [SupportByVersion("Visio", 11, 12, 14, 15, 16)]
        void PrintOut(NetOffice.VisioApi.Enums.VisPrintOutRange printRange, object fromPage, object toPage, object scaleCurrentViewToPaper, object printerName, object printToFile, object outputFileName);

        /// <summary>
        /// SupportByVersion Visio 11, 12, 14, 15, 16
        /// </summary>
        /// <param name="printRange">NetOffice.VisioApi.Enums.VisPrintOutRange printRange</param>
        /// <param name="fromPage">optional Int32 FromPage = 1</param>
        /// <param name="toPage">optional Int32 ToPage = -1</param>
        /// <param name="scaleCurrentViewToPaper">optional bool ScaleCurrentViewToPaper = false</param>
        /// <param name="printerName">optional string PrinterName = </param>
        /// <param name="printToFile">optional bool PrintToFile = false</param>
        /// <param name="outputFileName">optional string OutputFileName = </param>
        /// <param name="copies">optional Int32 Copies = 1</param>
        [CustomMethod]
        [SupportByVersion("Visio", 11, 12, 14, 15, 16)]
        void PrintOut(NetOffice.VisioApi.Enums.VisPrintOutRange printRange, object fromPage, object toPage, object scaleCurrentViewToPaper, object printerName, object printToFile, object outputFileName, object copies);

        /// <summary>
        /// SupportByVersion Visio 11, 12, 14, 15, 16
        /// </summary>
        /// <param name="printRange">NetOffice.VisioApi.Enums.VisPrintOutRange printRange</param>
        /// <param name="fromPage">optional Int32 FromPage = 1</param>
        /// <param name="toPage">optional Int32 ToPage = -1</param>
        /// <param name="scaleCurrentViewToPaper">optional bool ScaleCurrentViewToPaper = false</param>
        /// <param name="printerName">optional string PrinterName = </param>
        /// <param name="printToFile">optional bool PrintToFile = false</param>
        /// <param name="outputFileName">optional string OutputFileName = </param>
        /// <param name="copies">optional Int32 Copies = 1</param>
        /// <param name="collate">optional bool Collate = false</param>
        [CustomMethod]
        [SupportByVersion("Visio", 11, 12, 14, 15, 16)]
        void PrintOut(NetOffice.VisioApi.Enums.VisPrintOutRange printRange, object fromPage, object toPage, object scaleCurrentViewToPaper, object printerName, object printToFile, object outputFileName, object copies, object collate);

        /// <summary>
        /// SupportByVersion Visio 11, 12, 14, 15, 16
        /// </summary>
        /// <param name="bstrUndoScopeName">string bstrUndoScopeName</param>
        [SupportByVersion("Visio", 11, 12, 14, 15, 16)]
        Int32 BeginUndoScope(string bstrUndoScopeName);

        /// <summary>
        /// SupportByVersion Visio 11, 12, 14, 15, 16
        /// </summary>
        /// <param name="nScopeID">Int32 nScopeID</param>
        /// <param name="bCommit">bool bCommit</param>
        [SupportByVersion("Visio", 11, 12, 14, 15, 16)]
        void EndUndoScope(Int32 nScopeID, bool bCommit);

        /// <summary>
        /// SupportByVersion Visio 11, 12, 14, 15, 16
        /// </summary>
        /// <param name="pUndoUnit">object pUndoUnit</param>
        [SupportByVersion("Visio", 11, 12, 14, 15, 16)]
        void AddUndoUnit(object pUndoUnit);

        /// <summary>
        /// SupportByVersion Visio 11, 12, 14, 15, 16
        /// </summary>
        [SupportByVersion("Visio", 11, 12, 14, 15, 16)]
        void PurgeUndo();

        /// <summary>
        /// SupportByVersion Visio 11, 12, 14, 15, 16
        /// </summary>
        /// <param name="bstrScopeName">string bstrScopeName</param>
        [SupportByVersion("Visio", 11, 12, 14, 15, 16)]
        void RenameCurrentScope(string bstrScopeName);

        /// <summary>
        /// SupportByVersion Visio 12, 14, 15, 16
        /// </summary>
        /// <param name="removeHiddenInfoItems">Int32 removeHiddenInfoItems</param>
        [SupportByVersion("Visio", 12, 14, 15, 16)]
        void RemoveHiddenInformation(Int32 removeHiddenInfoItems);

        /// <summary>
        /// SupportByVersion Visio 12, 14, 15, 16
        /// </summary>
        /// <param name="eType">NetOffice.VisioApi.Enums.VisThemeTypes eType</param>
        /// <param name="nameArray">String[] nameArray</param>
        [SupportByVersion("Visio", 12, 14, 15, 16)]
        void GetThemeNames(NetOffice.VisioApi.Enums.VisThemeTypes eType, out String[] nameArray);

        /// <summary>
        /// SupportByVersion Visio 12, 14, 15, 16
        /// </summary>
        /// <param name="eType">NetOffice.VisioApi.Enums.VisThemeTypes eType</param>
        /// <param name="nameArray">String[] nameArray</param>
        [SupportByVersion("Visio", 12, 14, 15, 16)]
        void GetThemeNamesU(NetOffice.VisioApi.Enums.VisThemeTypes eType, out String[] nameArray);

        /// <summary>
        /// SupportByVersion Visio 12, 14, 15, 16
        /// </summary>
        [SupportByVersion("Visio", 12, 14, 15, 16)]
        bool CanUndoCheckOut();

        /// <summary>
        /// SupportByVersion Visio 12, 14, 15, 16
        /// </summary>
        [SupportByVersion("Visio", 12, 14, 15, 16)]
        void UndoCheckOut();

        /// <summary>
        /// SupportByVersion Visio 12, 14, 15, 16
        /// </summary>
        /// <param name="fixedFormat">NetOffice.VisioApi.Enums.VisFixedFormatTypes fixedFormat</param>
        /// <param name="outputFileName">string outputFileName</param>
        /// <param name="intent">NetOffice.VisioApi.Enums.VisDocExIntent intent</param>
        /// <param name="printRange">NetOffice.VisioApi.Enums.VisPrintOutRange printRange</param>
        /// <param name="fromPage">optional Int32 FromPage = 1</param>
        /// <param name="toPage">optional Int32 ToPage = -1</param>
        /// <param name="colorAsBlack">optional bool ColorAsBlack = false</param>
        /// <param name="includeBackground">optional bool IncludeBackground = true</param>
        /// <param name="includeDocumentProperties">optional bool IncludeDocumentProperties = true</param>
        /// <param name="includeStructureTags">optional bool IncludeStructureTags = true</param>
        /// <param name="useISO19005_1">optional bool UseISO19005_1 = false</param>
        /// <param name="fixedFormatExtClass">optional object fixedFormatExtClass</param>
        [SupportByVersion("Visio", 12, 14, 15, 16)]
        void ExportAsFixedFormat(NetOffice.VisioApi.Enums.VisFixedFormatTypes fixedFormat, string outputFileName, NetOffice.VisioApi.Enums.VisDocExIntent intent, NetOffice.VisioApi.Enums.VisPrintOutRange printRange, object fromPage, object toPage, object colorAsBlack, object includeBackground, object includeDocumentProperties, object includeStructureTags, object useISO19005_1, object fixedFormatExtClass);

        /// <summary>
        /// SupportByVersion Visio 12, 14, 15, 16
        /// </summary>
        /// <param name="fixedFormat">NetOffice.VisioApi.Enums.VisFixedFormatTypes fixedFormat</param>
        /// <param name="outputFileName">string outputFileName</param>
        /// <param name="intent">NetOffice.VisioApi.Enums.VisDocExIntent intent</param>
        /// <param name="printRange">NetOffice.VisioApi.Enums.VisPrintOutRange printRange</param>
        [CustomMethod]
        [SupportByVersion("Visio", 12, 14, 15, 16)]
        void ExportAsFixedFormat(NetOffice.VisioApi.Enums.VisFixedFormatTypes fixedFormat, string outputFileName, NetOffice.VisioApi.Enums.VisDocExIntent intent, NetOffice.VisioApi.Enums.VisPrintOutRange printRange);

        /// <summary>
        /// SupportByVersion Visio 12, 14, 15, 16
        /// </summary>
        /// <param name="fixedFormat">NetOffice.VisioApi.Enums.VisFixedFormatTypes fixedFormat</param>
        /// <param name="outputFileName">string outputFileName</param>
        /// <param name="intent">NetOffice.VisioApi.Enums.VisDocExIntent intent</param>
        /// <param name="printRange">NetOffice.VisioApi.Enums.VisPrintOutRange printRange</param>
        /// <param name="fromPage">optional Int32 FromPage = 1</param>
        [CustomMethod]
        [SupportByVersion("Visio", 12, 14, 15, 16)]
        void ExportAsFixedFormat(NetOffice.VisioApi.Enums.VisFixedFormatTypes fixedFormat, string outputFileName, NetOffice.VisioApi.Enums.VisDocExIntent intent, NetOffice.VisioApi.Enums.VisPrintOutRange printRange, object fromPage);

        /// <summary>
        /// SupportByVersion Visio 12, 14, 15, 16
        /// </summary>
        /// <param name="fixedFormat">NetOffice.VisioApi.Enums.VisFixedFormatTypes fixedFormat</param>
        /// <param name="outputFileName">string outputFileName</param>
        /// <param name="intent">NetOffice.VisioApi.Enums.VisDocExIntent intent</param>
        /// <param name="printRange">NetOffice.VisioApi.Enums.VisPrintOutRange printRange</param>
        /// <param name="fromPage">optional Int32 FromPage = 1</param>
        /// <param name="toPage">optional Int32 ToPage = -1</param>
        [CustomMethod]
        [SupportByVersion("Visio", 12, 14, 15, 16)]
        void ExportAsFixedFormat(NetOffice.VisioApi.Enums.VisFixedFormatTypes fixedFormat, string outputFileName, NetOffice.VisioApi.Enums.VisDocExIntent intent, NetOffice.VisioApi.Enums.VisPrintOutRange printRange, object fromPage, object toPage);

        /// <summary>
        /// SupportByVersion Visio 12, 14, 15, 16
        /// </summary>
        /// <param name="fixedFormat">NetOffice.VisioApi.Enums.VisFixedFormatTypes fixedFormat</param>
        /// <param name="outputFileName">string outputFileName</param>
        /// <param name="intent">NetOffice.VisioApi.Enums.VisDocExIntent intent</param>
        /// <param name="printRange">NetOffice.VisioApi.Enums.VisPrintOutRange printRange</param>
        /// <param name="fromPage">optional Int32 FromPage = 1</param>
        /// <param name="toPage">optional Int32 ToPage = -1</param>
        /// <param name="colorAsBlack">optional bool ColorAsBlack = false</param>
        [CustomMethod]
        [SupportByVersion("Visio", 12, 14, 15, 16)]
        void ExportAsFixedFormat(NetOffice.VisioApi.Enums.VisFixedFormatTypes fixedFormat, string outputFileName, NetOffice.VisioApi.Enums.VisDocExIntent intent, NetOffice.VisioApi.Enums.VisPrintOutRange printRange, object fromPage, object toPage, object colorAsBlack);

        /// <summary>
        /// SupportByVersion Visio 12, 14, 15, 16
        /// </summary>
        /// <param name="fixedFormat">NetOffice.VisioApi.Enums.VisFixedFormatTypes fixedFormat</param>
        /// <param name="outputFileName">string outputFileName</param>
        /// <param name="intent">NetOffice.VisioApi.Enums.VisDocExIntent intent</param>
        /// <param name="printRange">NetOffice.VisioApi.Enums.VisPrintOutRange printRange</param>
        /// <param name="fromPage">optional Int32 FromPage = 1</param>
        /// <param name="toPage">optional Int32 ToPage = -1</param>
        /// <param name="colorAsBlack">optional bool ColorAsBlack = false</param>
        /// <param name="includeBackground">optional bool IncludeBackground = true</param>
        [CustomMethod]
        [SupportByVersion("Visio", 12, 14, 15, 16)]
        void ExportAsFixedFormat(NetOffice.VisioApi.Enums.VisFixedFormatTypes fixedFormat, string outputFileName, NetOffice.VisioApi.Enums.VisDocExIntent intent, NetOffice.VisioApi.Enums.VisPrintOutRange printRange, object fromPage, object toPage, object colorAsBlack, object includeBackground);

        /// <summary>
        /// SupportByVersion Visio 12, 14, 15, 16
        /// </summary>
        /// <param name="fixedFormat">NetOffice.VisioApi.Enums.VisFixedFormatTypes fixedFormat</param>
        /// <param name="outputFileName">string outputFileName</param>
        /// <param name="intent">NetOffice.VisioApi.Enums.VisDocExIntent intent</param>
        /// <param name="printRange">NetOffice.VisioApi.Enums.VisPrintOutRange printRange</param>
        /// <param name="fromPage">optional Int32 FromPage = 1</param>
        /// <param name="toPage">optional Int32 ToPage = -1</param>
        /// <param name="colorAsBlack">optional bool ColorAsBlack = false</param>
        /// <param name="includeBackground">optional bool IncludeBackground = true</param>
        /// <param name="includeDocumentProperties">optional bool IncludeDocumentProperties = true</param>
        [CustomMethod]
        [SupportByVersion("Visio", 12, 14, 15, 16)]
        void ExportAsFixedFormat(NetOffice.VisioApi.Enums.VisFixedFormatTypes fixedFormat, string outputFileName, NetOffice.VisioApi.Enums.VisDocExIntent intent, NetOffice.VisioApi.Enums.VisPrintOutRange printRange, object fromPage, object toPage, object colorAsBlack, object includeBackground, object includeDocumentProperties);

        /// <summary>
        /// SupportByVersion Visio 12, 14, 15, 16
        /// </summary>
        /// <param name="fixedFormat">NetOffice.VisioApi.Enums.VisFixedFormatTypes fixedFormat</param>
        /// <param name="outputFileName">string outputFileName</param>
        /// <param name="intent">NetOffice.VisioApi.Enums.VisDocExIntent intent</param>
        /// <param name="printRange">NetOffice.VisioApi.Enums.VisPrintOutRange printRange</param>
        /// <param name="fromPage">optional Int32 FromPage = 1</param>
        /// <param name="toPage">optional Int32 ToPage = -1</param>
        /// <param name="colorAsBlack">optional bool ColorAsBlack = false</param>
        /// <param name="includeBackground">optional bool IncludeBackground = true</param>
        /// <param name="includeDocumentProperties">optional bool IncludeDocumentProperties = true</param>
        /// <param name="includeStructureTags">optional bool IncludeStructureTags = true</param>
        [CustomMethod]
        [SupportByVersion("Visio", 12, 14, 15, 16)]
        void ExportAsFixedFormat(NetOffice.VisioApi.Enums.VisFixedFormatTypes fixedFormat, string outputFileName, NetOffice.VisioApi.Enums.VisDocExIntent intent, NetOffice.VisioApi.Enums.VisPrintOutRange printRange, object fromPage, object toPage, object colorAsBlack, object includeBackground, object includeDocumentProperties, object includeStructureTags);

        /// <summary>
        /// SupportByVersion Visio 12, 14, 15, 16
        /// </summary>
        /// <param name="fixedFormat">NetOffice.VisioApi.Enums.VisFixedFormatTypes fixedFormat</param>
        /// <param name="outputFileName">string outputFileName</param>
        /// <param name="intent">NetOffice.VisioApi.Enums.VisDocExIntent intent</param>
        /// <param name="printRange">NetOffice.VisioApi.Enums.VisPrintOutRange printRange</param>
        /// <param name="fromPage">optional Int32 FromPage = 1</param>
        /// <param name="toPage">optional Int32 ToPage = -1</param>
        /// <param name="colorAsBlack">optional bool ColorAsBlack = false</param>
        /// <param name="includeBackground">optional bool IncludeBackground = true</param>
        /// <param name="includeDocumentProperties">optional bool IncludeDocumentProperties = true</param>
        /// <param name="includeStructureTags">optional bool IncludeStructureTags = true</param>
        /// <param name="useISO19005_1">optional bool UseISO19005_1 = false</param>
        [CustomMethod]
        [SupportByVersion("Visio", 12, 14, 15, 16)]
        void ExportAsFixedFormat(NetOffice.VisioApi.Enums.VisFixedFormatTypes fixedFormat, string outputFileName, NetOffice.VisioApi.Enums.VisDocExIntent intent, NetOffice.VisioApi.Enums.VisPrintOutRange printRange, object fromPage, object toPage, object colorAsBlack, object includeBackground, object includeDocumentProperties, object includeStructureTags, object useISO19005_1);

        #endregion
    }
}
