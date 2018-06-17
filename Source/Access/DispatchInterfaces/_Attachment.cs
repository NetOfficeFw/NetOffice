using System;
using NetRuntimeSystem = System;
using System.ComponentModel;
using NetOffice.Attributes;

namespace NetOffice.AccessApi
{
    /// <summary>
    /// _Attachment
    /// </summary>
    [SyntaxBypass]
    public interface _Attachment_ : NetOffice.OfficeApi.IAccessible
    {
        #region Properties

        /// <summary>
        /// SupportByVersion Access 12, 14, 15, 16
        /// Get
        /// </summary>
        /// <param name="var">optional object var</param>
        /// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff193567.aspx
        [SupportByVersion("Access", 12, 14, 15, 16)]
        [EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
        string get_FileName(object var);

        /// <summary>
        /// SupportByVersion Access 12, 14, 15, 16
        /// Alias for get_FileName
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff193567.aspx </remarks>
        /// <param name="var">optional object var</param>
        [SupportByVersion("Access", 12, 14, 15, 16), Redirect("get_FileName")]
        string FileName(object var);

        /// <summary>
        /// SupportByVersion Access 12, 14, 15, 16
        /// Get
        /// </summary>
        /// <param name="var">optional object var</param>
        /// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff845243.aspx
        [SupportByVersion("Access", 12, 14, 15, 16)]
        [EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
        string get_FileType(object var);

        /// <summary>
        /// SupportByVersion Access 12, 14, 15, 16
        /// Alias for get_FileType
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff845243.aspx </remarks>
        /// <param name="var">optional object var</param>
        [SupportByVersion("Access", 12, 14, 15, 16), Redirect("get_FileType")]
        string FileType(object var);

        /// <summary>
        /// SupportByVersion Access 12, 14, 15, 16
        /// Get
        /// </summary>
        /// <param name="var">optional object var</param>
        /// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff195201.aspx
        [SupportByVersion("Access", 12, 14, 15, 16)]
        [EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
        string get_FileURL(object var);

        /// <summary>
        /// SupportByVersion Access 12, 14, 15, 16
        /// Alias for get_FileURL
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff195201.aspx </remarks>
        /// <param name="var">optional object var</param>
        [SupportByVersion("Access", 12, 14, 15, 16), Redirect("get_FileURL")]
        string FileURL(object var);

        /// <summary>
        /// SupportByVersion Access 12, 14, 15, 16
        /// Get
        /// </summary>
        /// <param name="var">optional object var</param>
        [SupportByVersion("Access", 12, 14, 15, 16)]
        [EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
        object get_FileData(object var);

        /// <summary>
        /// SupportByVersion Access 12, 14, 15, 16
        /// Alias for get_FileData
        /// </summary>
        /// <param name="var">optional object var</param>
        [SupportByVersion("Access", 12, 14, 15, 16), Redirect("get_FileData")]
        object FileData(object var);

        /// <summary>
        /// SupportByVersion Access 12, 14, 15, 16
        /// Get
        /// </summary>
        /// <param name="var">optional object var</param>
        [SupportByVersion("Access", 12, 14, 15, 16)]
        [EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
        object get_PictureDisp(object var);

        /// <summary>
        /// SupportByVersion Access 12, 14, 15, 16
        /// Alias for get_PictureDisp
        /// </summary>
        /// <param name="var">optional object var</param>
        [SupportByVersion("Access", 12, 14, 15, 16), Redirect("get_PictureDisp")]
        object PictureDisp(object var);

        #endregion
    }

    /// <summary>
    /// DispatchInterface _Attachment 
    /// SupportByVersion Access, 12,14,15,16
    /// </summary>
    [SupportByVersion("Access", 12, 14, 15, 16)]
    [EntityType(EntityType.IsDispatchInterface), BaseType]
	[TypeId("3B06E980-E47C-11CD-8701-00AA003F0F07")]
    [CoClassSource(typeof(NetOffice.AccessApi.Attachment))]
    public interface _Attachment : _Attachment_
    {
        #region Properties

        /// <summary>
        /// SupportByVersion Access 12, 14, 15, 16
        /// Get
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff835342.aspx </remarks>
        [SupportByVersion("Access", 12, 14, 15, 16)]
        NetOffice.AccessApi.Application Application { get; }

        /// <summary>
        /// SupportByVersion Access 12, 14, 15, 16
        /// Get
        /// Unknown COM Proxy
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff820946.aspx </remarks>
        [SupportByVersion("Access", 12, 14, 15, 16), ProxyResult]
        object Parent { get; }

        /// <summary>
        /// SupportByVersion Access 12, 14, 15, 16
        /// Get
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff836659.aspx </remarks>
        [SupportByVersion("Access", 12, 14, 15, 16)]
        object OldValue { get; }

        /// <summary>
        /// SupportByVersion Access 12, 14, 15, 16
        /// Get
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff192518.aspx </remarks>
        [SupportByVersion("Access", 12, 14, 15, 16)]
        NetOffice.AccessApi.Properties Properties { get; }

        /// <summary>
        /// SupportByVersion Access 12, 14, 15, 16
        /// Get
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff196010.aspx </remarks>
        [SupportByVersion("Access", 12, 14, 15, 16)]
        NetOffice.AccessApi.Children Controls { get; }

        /// <summary>
        /// SupportByVersion Access 12, 14, 15, 16
        /// Get
        /// </summary>
        [SupportByVersion("Access", 12, 14, 15, 16)]
        [BaseResult]
        [EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
        NetOffice.AccessApi._Hyperlink Hyperlink { get; }

        /// <summary>
        /// SupportByVersion Access 12, 14, 15, 16
        /// Get/Set
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff836727.aspx </remarks>
        [SupportByVersion("Access", 12, 14, 15, 16)]
        string EventProcPrefix { get; set; }

        /// <summary>
        /// SupportByVersion Access 12, 14, 15, 16
        /// Get/Set
        /// </summary>
        [SupportByVersion("Access", 12, 14, 15, 16)]
        [EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
        string _Name { get; set; }

        /// <summary>
        /// SupportByVersion Access 12, 14, 15, 16
        /// Get/Set
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff836871.aspx </remarks>
        [SupportByVersion("Access", 12, 14, 15, 16)]
        byte ControlType { get; set; }

        /// <summary>
        /// SupportByVersion Access 12, 14, 15, 16
        /// Get/Set
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff844989.aspx </remarks>
        [SupportByVersion("Access", 12, 14, 15, 16)]
        byte PictureSizeMode { get; set; }

        /// <summary>
        /// SupportByVersion Access 12, 14, 15, 16
        /// Get/Set
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff193807.aspx </remarks>
        [SupportByVersion("Access", 12, 14, 15, 16)]
        byte PictureAlignment { get; set; }

        /// <summary>
        /// SupportByVersion Access 12, 14, 15, 16
        /// Get/Set
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff835074.aspx </remarks>
        [SupportByVersion("Access", 12, 14, 15, 16)]
        bool PictureTiling { get; set; }

        /// <summary>
        /// SupportByVersion Access 12, 14, 15, 16
        /// Get/Set
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff845492.aspx </remarks>
        [SupportByVersion("Access", 12, 14, 15, 16)]
        bool Visible { get; set; }

        /// <summary>
        /// SupportByVersion Access 12, 14, 15, 16
        /// Get/Set
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff195722.aspx </remarks>
        [SupportByVersion("Access", 12, 14, 15, 16)]
        byte DisplayWhen { get; set; }

        /// <summary>
        /// SupportByVersion Access 12, 14, 15, 16
        /// Get/Set
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff834511.aspx </remarks>
        [SupportByVersion("Access", 12, 14, 15, 16)]
        Int16 Left { get; set; }

        /// <summary>
        /// SupportByVersion Access 12, 14, 15, 16
        /// Get/Set
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff844867.aspx </remarks>
        [SupportByVersion("Access", 12, 14, 15, 16)]
        Int16 Top { get; set; }

        /// <summary>
        /// SupportByVersion Access 12, 14, 15, 16
        /// Get/Set
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff195738.aspx </remarks>
        [SupportByVersion("Access", 12, 14, 15, 16)]
        Int16 Width { get; set; }

        /// <summary>
        /// SupportByVersion Access 12, 14, 15, 16
        /// Get/Set
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff192489.aspx </remarks>
        [SupportByVersion("Access", 12, 14, 15, 16)]
        Int16 Height { get; set; }

        /// <summary>
        /// SupportByVersion Access 12, 14, 15, 16
        /// Get/Set
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff196447.aspx </remarks>
        [SupportByVersion("Access", 12, 14, 15, 16)]
        byte BackStyle { get; set; }

        /// <summary>
        /// SupportByVersion Access 12, 14, 15, 16
        /// Get/Set
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff196184.aspx </remarks>
        [SupportByVersion("Access", 12, 14, 15, 16)]
        Int32 BackColor { get; set; }

        /// <summary>
        /// SupportByVersion Access 12, 14, 15, 16
        /// Get/Set
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff193541.aspx </remarks>
        [SupportByVersion("Access", 12, 14, 15, 16)]
        byte SpecialEffect { get; set; }

        /// <summary>
        /// SupportByVersion Access 12, 14, 15, 16
        /// Get/Set
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff820796.aspx </remarks>
        [SupportByVersion("Access", 12, 14, 15, 16)]
        byte BorderStyle { get; set; }

        /// <summary>
        /// SupportByVersion Access 12, 14, 15, 16
        /// Get/Set
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff821483.aspx </remarks>
        [SupportByVersion("Access", 12, 14, 15, 16)]
        byte OldBorderStyle { get; set; }

        /// <summary>
        /// SupportByVersion Access 12, 14, 15, 16
        /// Get/Set
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff834465.aspx </remarks>
        [SupportByVersion("Access", 12, 14, 15, 16)]
        Int32 BorderColor { get; set; }

        /// <summary>
        /// SupportByVersion Access 12, 14, 15, 16
        /// Get/Set
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff835995.aspx </remarks>
        [SupportByVersion("Access", 12, 14, 15, 16)]
        byte BorderWidth { get; set; }

        /// <summary>
        /// SupportByVersion Access 12, 14, 15, 16
        /// Get/Set
        /// </summary>
        [SupportByVersion("Access", 12, 14, 15, 16)]
        [EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
        byte BorderLineStyle { get; set; }

        /// <summary>
        /// SupportByVersion Access 12, 14, 15, 16
        /// Get/Set
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff823119.aspx </remarks>
        [SupportByVersion("Access", 12, 14, 15, 16)]
        string ControlTipText { get; set; }

        /// <summary>
        /// SupportByVersion Access 12, 14, 15, 16
        /// Get/Set
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff821449.aspx </remarks>
        [SupportByVersion("Access", 12, 14, 15, 16)]
        Int32 HelpContextId { get; set; }

        /// <summary>
        /// SupportByVersion Access 12, 14, 15, 16
        /// Get/Set
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff837166.aspx </remarks>
        [SupportByVersion("Access", 12, 14, 15, 16)]
        Int16 Section { get; set; }

        /// <summary>
        /// SupportByVersion Access 12, 14, 15, 16
        /// Get/Set
        /// </summary>
        [SupportByVersion("Access", 12, 14, 15, 16)]
        [EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
        string ControlName { get; set; }

        /// <summary>
        /// SupportByVersion Access 12, 14, 15, 16
        /// Get/Set
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff194772.aspx </remarks>
        [SupportByVersion("Access", 12, 14, 15, 16)]
        bool IsVisible { get; set; }

        /// <summary>
        /// SupportByVersion Access 12, 14, 15, 16
        /// Get/Set
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff193160.aspx </remarks>
        [SupportByVersion("Access", 12, 14, 15, 16)]
        string BeforeUpdate { get; set; }

        /// <summary>
        /// SupportByVersion Access 12, 14, 15, 16
        /// Get/Set
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff194169.aspx </remarks>
        [SupportByVersion("Access", 12, 14, 15, 16)]
        string AfterUpdate { get; set; }

        /// <summary>
        /// SupportByVersion Access 12, 14, 15, 16
        /// Get/Set
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff194485.aspx </remarks>
        [SupportByVersion("Access", 12, 14, 15, 16)]
        string OnEnter { get; set; }

        /// <summary>
        /// SupportByVersion Access 12, 14, 15, 16
        /// Get/Set
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff194575.aspx </remarks>
        [SupportByVersion("Access", 12, 14, 15, 16)]
        string OnExit { get; set; }

        /// <summary>
        /// SupportByVersion Access 12, 14, 15, 16
        /// Get/Set
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff821070.aspx </remarks>
        [SupportByVersion("Access", 12, 14, 15, 16)]
        string OnDirty { get; set; }

        /// <summary>
        /// SupportByVersion Access 12, 14, 15, 16
        /// Get/Set
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff823022.aspx </remarks>
        [SupportByVersion("Access", 12, 14, 15, 16)]
        string OnChange { get; set; }

        /// <summary>
        /// SupportByVersion Access 12, 14, 15, 16
        /// Get/Set
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff820975.aspx </remarks>
        [SupportByVersion("Access", 12, 14, 15, 16)]
        string OnGotFocus { get; set; }

        /// <summary>
        /// SupportByVersion Access 12, 14, 15, 16
        /// Get/Set
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff194110.aspx </remarks>
        [SupportByVersion("Access", 12, 14, 15, 16)]
        string OnLostFocus { get; set; }

        /// <summary>
        /// SupportByVersion Access 12, 14, 15, 16
        /// Get/Set
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff822868.aspx </remarks>
        [SupportByVersion("Access", 12, 14, 15, 16)]
        string OnClick { get; set; }

        /// <summary>
        /// SupportByVersion Access 12, 14, 15, 16
        /// Get/Set
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff194550.aspx </remarks>
        [SupportByVersion("Access", 12, 14, 15, 16)]
        string OnDblClick { get; set; }

        /// <summary>
        /// SupportByVersion Access 12, 14, 15, 16
        /// Get/Set
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff195807.aspx </remarks>
        [SupportByVersion("Access", 12, 14, 15, 16)]
        string OnMouseDown { get; set; }

        /// <summary>
        /// SupportByVersion Access 12, 14, 15, 16
        /// Get/Set
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff193537.aspx </remarks>
        [SupportByVersion("Access", 12, 14, 15, 16)]
        string OnMouseMove { get; set; }

        /// <summary>
        /// SupportByVersion Access 12, 14, 15, 16
        /// Get/Set
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff844980.aspx </remarks>
        [SupportByVersion("Access", 12, 14, 15, 16)]
        string OnMouseUp { get; set; }

        /// <summary>
        /// SupportByVersion Access 12, 14, 15, 16
        /// Get/Set
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff196490.aspx </remarks>
        [SupportByVersion("Access", 12, 14, 15, 16)]
        string OnKeyDown { get; set; }

        /// <summary>
        /// SupportByVersion Access 12, 14, 15, 16
        /// Get/Set
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff194243.aspx </remarks>
        [SupportByVersion("Access", 12, 14, 15, 16)]
        string OnKeyUp { get; set; }

        /// <summary>
        /// SupportByVersion Access 12, 14, 15, 16
        /// Get/Set
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff198361.aspx </remarks>
        [SupportByVersion("Access", 12, 14, 15, 16)]
        string OnKeyPress { get; set; }

        /// <summary>
        /// SupportByVersion Access 12, 14, 15, 16
        /// Get/Set
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff196159.aspx </remarks>
        [SupportByVersion("Access", 12, 14, 15, 16)]
        string OnAttachmentCurrent { get; set; }

        /// <summary>
        /// SupportByVersion Access 12, 14, 15, 16
        /// Get/Set
        /// </summary>
        [SupportByVersion("Access", 12, 14, 15, 16)]
        [EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
        string BeforeUpdateMacro { get; set; }

        /// <summary>
        /// SupportByVersion Access 12, 14, 15, 16
        /// Get/Set
        /// </summary>
        [SupportByVersion("Access", 12, 14, 15, 16)]
        [EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
        string AfterUpdateMacro { get; set; }

        /// <summary>
        /// SupportByVersion Access 12, 14, 15, 16
        /// Get/Set
        /// </summary>
        [SupportByVersion("Access", 12, 14, 15, 16)]
        [EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
        string OnEnterMacro { get; set; }

        /// <summary>
        /// SupportByVersion Access 12, 14, 15, 16
        /// Get/Set
        /// </summary>
        [SupportByVersion("Access", 12, 14, 15, 16)]
        [EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
        string OnExitMacro { get; set; }

        /// <summary>
        /// SupportByVersion Access 12, 14, 15, 16
        /// Get/Set
        /// </summary>
        [SupportByVersion("Access", 12, 14, 15, 16)]
        [EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
        string OnDirtyMacro { get; set; }

        /// <summary>
        /// SupportByVersion Access 12, 14, 15, 16
        /// Get/Set
        /// </summary>
        [SupportByVersion("Access", 12, 14, 15, 16)]
        [EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
        string OnChangeMacro { get; set; }

        /// <summary>
        /// SupportByVersion Access 12, 14, 15, 16
        /// Get/Set
        /// </summary>
        [SupportByVersion("Access", 12, 14, 15, 16)]
        [EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
        string OnGotFocusMacro { get; set; }

        /// <summary>
        /// SupportByVersion Access 12, 14, 15, 16
        /// Get/Set
        /// </summary>
        [SupportByVersion("Access", 12, 14, 15, 16)]
        [EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
        string OnLostFocusMacro { get; set; }

        /// <summary>
        /// SupportByVersion Access 12, 14, 15, 16
        /// Get/Set
        /// </summary>
        [SupportByVersion("Access", 12, 14, 15, 16)]
        [EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
        string OnClickMacro { get; set; }

        /// <summary>
        /// SupportByVersion Access 12, 14, 15, 16
        /// Get/Set
        /// </summary>
        [SupportByVersion("Access", 12, 14, 15, 16)]
        [EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
        string OnDblClickMacro { get; set; }

        /// <summary>
        /// SupportByVersion Access 12, 14, 15, 16
        /// Get/Set
        /// </summary>
        [SupportByVersion("Access", 12, 14, 15, 16)]
        [EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
        string OnMouseDownMacro { get; set; }

        /// <summary>
        /// SupportByVersion Access 12, 14, 15, 16
        /// Get/Set
        /// </summary>
        [SupportByVersion("Access", 12, 14, 15, 16)]
        [EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
        string OnMouseMoveMacro { get; set; }

        /// <summary>
        /// SupportByVersion Access 12, 14, 15, 16
        /// Get/Set
        /// </summary>
        [SupportByVersion("Access", 12, 14, 15, 16)]
        [EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
        string OnMouseUpMacro { get; set; }

        /// <summary>
        /// SupportByVersion Access 12, 14, 15, 16
        /// Get/Set
        /// </summary>
        [SupportByVersion("Access", 12, 14, 15, 16)]
        [EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
        string OnKeyDownMacro { get; set; }

        /// <summary>
        /// SupportByVersion Access 12, 14, 15, 16
        /// Get/Set
        /// </summary>
        [SupportByVersion("Access", 12, 14, 15, 16)]
        [EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
        string OnKeyUpMacro { get; set; }

        /// <summary>
        /// SupportByVersion Access 12, 14, 15, 16
        /// Get/Set
        /// </summary>
        [SupportByVersion("Access", 12, 14, 15, 16)]
        [EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
        string OnKeyPressMacro { get; set; }

        /// <summary>
        /// SupportByVersion Access 12, 14, 15, 16
        /// Get/Set
        /// </summary>
        [SupportByVersion("Access", 12, 14, 15, 16)]
        [EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
        string OnAttachmentCurrentMacro { get; set; }

        /// <summary>
        /// SupportByVersion Access 12, 14, 15, 16
        /// Get/Set
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff822748.aspx </remarks>
        [SupportByVersion("Access", 12, 14, 15, 16)]
        string ShortcutMenuBar { get; set; }

        /// <summary>
        /// SupportByVersion Access 12, 14, 15, 16
        /// Get/Set
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff835668.aspx </remarks>
        [SupportByVersion("Access", 12, 14, 15, 16)]
        [EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
        bool InSelection { get; set; }

        /// <summary>
        /// SupportByVersion Access 12, 14, 15, 16
        /// Get/Set
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff195530.aspx </remarks>
        [SupportByVersion("Access", 12, 14, 15, 16)]
        string Tag { get; set; }

        /// <summary>
        /// SupportByVersion Access 12, 14, 15, 16
        /// Get/Set
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff194767.aspx </remarks>
        [SupportByVersion("Access", 12, 14, 15, 16)]
        string Name { get; set; }

        /// <summary>
        /// SupportByVersion Access 12, 14, 15, 16
        /// Get/Set
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff821403.aspx </remarks>
        [SupportByVersion("Access", 12, 14, 15, 16)]
        NetOffice.AccessApi.Enums.AcDisplayAs DisplayAs { get; set; }

        /// <summary>
        /// SupportByVersion Access 12, 14, 15, 16
        /// Get
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff192274.aspx </remarks>
        [SupportByVersion("Access", 12, 14, 15, 16)]
        Int32 AttachmentCount { get; }

        /// <summary>
        /// SupportByVersion Access 12, 14, 15, 16
        /// Get/Set
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff197073.aspx </remarks>
        [SupportByVersion("Access", 12, 14, 15, 16)]
        Int32 CurrentAttachment { get; set; }

        /// <summary>
        /// SupportByVersion Access 12, 14, 15, 16
        /// Get
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff193567.aspx </remarks>
        [SupportByVersion("Access", 12, 14, 15, 16)]
        new string FileName { get; }

        /// <summary>
        /// SupportByVersion Access 12, 14, 15, 16
        /// Get
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff845243.aspx </remarks>
        [SupportByVersion("Access", 12, 14, 15, 16)]
        new string FileType { get; }

        /// <summary>
        /// SupportByVersion Access 12, 14, 15, 16
        /// Get
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff195201.aspx </remarks>
        [SupportByVersion("Access", 12, 14, 15, 16)]
        new string FileURL { get; }

        /// <summary>
        /// SupportByVersion Access 12, 14, 15, 16
        /// Get/Set
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff845195.aspx </remarks>
        [SupportByVersion("Access", 12, 14, 15, 16)]
        NetOffice.AccessApi.Enums.AcHorizontalAnchor HorizontalAnchor { get; set; }

        /// <summary>
        /// SupportByVersion Access 12, 14, 15, 16
        /// Get/Set
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff822708.aspx </remarks>
        [SupportByVersion("Access", 12, 14, 15, 16)]
        NetOffice.AccessApi.Enums.AcVerticalAnchor VerticalAnchor { get; set; }

        /// <summary>
        /// SupportByVersion Access 12, 14, 15, 16
        /// Get
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff845398.aspx </remarks>
        [SupportByVersion("Access", 12, 14, 15, 16)]
        NetOffice.AccessApi.Enums.AcLayoutType Layout { get; }

        /// <summary>
        /// SupportByVersion Access 12, 14, 15, 16
        /// Get/Set
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff195463.aspx </remarks>
        [SupportByVersion("Access", 12, 14, 15, 16)]
        Int16 LeftPadding { get; set; }

        /// <summary>
        /// SupportByVersion Access 12, 14, 15, 16
        /// Get/Set
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff845619.aspx </remarks>
        [SupportByVersion("Access", 12, 14, 15, 16)]
        Int16 TopPadding { get; set; }

        /// <summary>
        /// SupportByVersion Access 12, 14, 15, 16
        /// Get/Set
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff821735.aspx </remarks>
        [SupportByVersion("Access", 12, 14, 15, 16)]
        Int16 RightPadding { get; set; }

        /// <summary>
        /// SupportByVersion Access 12, 14, 15, 16
        /// Get/Set
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff822438.aspx </remarks>
        [SupportByVersion("Access", 12, 14, 15, 16)]
        Int16 BottomPadding { get; set; }

        /// <summary>
        /// SupportByVersion Access 12, 14, 15, 16
        /// Get/Set
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff845104.aspx </remarks>
        [SupportByVersion("Access", 12, 14, 15, 16)]
        byte GridlineStyleLeft { get; set; }

        /// <summary>
        /// SupportByVersion Access 12, 14, 15, 16
        /// Get/Set
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff822793.aspx </remarks>
        [SupportByVersion("Access", 12, 14, 15, 16)]
        byte GridlineStyleTop { get; set; }

        /// <summary>
        /// SupportByVersion Access 12, 14, 15, 16
        /// Get/Set
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff823207.aspx </remarks>
        [SupportByVersion("Access", 12, 14, 15, 16)]
        byte GridlineStyleRight { get; set; }

        /// <summary>
        /// SupportByVersion Access 12, 14, 15, 16
        /// Get/Set
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff192328.aspx </remarks>
        [SupportByVersion("Access", 12, 14, 15, 16)]
        byte GridlineStyleBottom { get; set; }

        /// <summary>
        /// SupportByVersion Access 12, 14, 15, 16
        /// Get/Set
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff192080.aspx </remarks>
        [SupportByVersion("Access", 12, 14, 15, 16)]
        byte GridlineWidthLeft { get; set; }

        /// <summary>
        /// SupportByVersion Access 12, 14, 15, 16
        /// Get/Set
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff195855.aspx </remarks>
        [SupportByVersion("Access", 12, 14, 15, 16)]
        byte GridlineWidthTop { get; set; }

        /// <summary>
        /// SupportByVersion Access 12, 14, 15, 16
        /// Get/Set
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff195871.aspx </remarks>
        [SupportByVersion("Access", 12, 14, 15, 16)]
        byte GridlineWidthRight { get; set; }

        /// <summary>
        /// SupportByVersion Access 12, 14, 15, 16
        /// Get/Set
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff193208.aspx </remarks>
        [SupportByVersion("Access", 12, 14, 15, 16)]
        byte GridlineWidthBottom { get; set; }

        /// <summary>
        /// SupportByVersion Access 12, 14, 15, 16
        /// Get/Set
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff836375.aspx </remarks>
        [SupportByVersion("Access", 12, 14, 15, 16)]
        Int32 GridlineColor { get; set; }

        /// <summary>
        /// SupportByVersion Access 12, 14, 15, 16
        /// Get/Set
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff197948.aspx </remarks>
        [SupportByVersion("Access", 12, 14, 15, 16)]
        string DefaultPicture { get; set; }

        /// <summary>
        /// SupportByVersion Access 12, 14, 15, 16
        /// Get
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff835967.aspx </remarks>
        [SupportByVersion("Access", 12, 14, 15, 16)]
        Int32 LayoutID { get; }

        /// <summary>
        /// SupportByVersion Access 12, 14, 15, 16
        /// Get/Set
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff845041.aspx </remarks>
        [SupportByVersion("Access", 12, 14, 15, 16)]
        bool AutoLabel { get; set; }

        /// <summary>
        /// SupportByVersion Access 12, 14, 15, 16
        /// Get/Set
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff197385.aspx </remarks>
        [SupportByVersion("Access", 12, 14, 15, 16)]
        bool AddColon { get; set; }

        /// <summary>
        /// SupportByVersion Access 12, 14, 15, 16
        /// Get/Set
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff195238.aspx </remarks>
        [SupportByVersion("Access", 12, 14, 15, 16)]
        Int16 LabelX { get; set; }

        /// <summary>
        /// SupportByVersion Access 12, 14, 15, 16
        /// Get/Set
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff193492.aspx </remarks>
        [SupportByVersion("Access", 12, 14, 15, 16)]
        Int16 LabelY { get; set; }

        /// <summary>
        /// SupportByVersion Access 12, 14, 15, 16
        /// Get/Set
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff192309.aspx </remarks>
        [SupportByVersion("Access", 12, 14, 15, 16)]
        byte LabelAlign { get; set; }

        /// <summary>
        /// SupportByVersion Access 12, 14, 15, 16
        /// Get/Set
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff845588.aspx </remarks>
        [SupportByVersion("Access", 12, 14, 15, 16)]
        Int16 ColumnWidth { get; set; }

        /// <summary>
        /// SupportByVersion Access 12, 14, 15, 16
        /// Get/Set
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff835664.aspx </remarks>
        [SupportByVersion("Access", 12, 14, 15, 16)]
        Int16 ColumnOrder { get; set; }

        /// <summary>
        /// SupportByVersion Access 12, 14, 15, 16
        /// Get/Set
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff196054.aspx </remarks>
        [SupportByVersion("Access", 12, 14, 15, 16)]
        bool ColumnHidden { get; set; }

        /// <summary>
        /// SupportByVersion Access 12, 14, 15, 16
        /// Get/Set
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff195499.aspx </remarks>
        [SupportByVersion("Access", 12, 14, 15, 16)]
        string ControlSource { get; set; }

        /// <summary>
        /// SupportByVersion Access 12, 14, 15, 16
        /// Get/Set
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff198248.aspx </remarks>
        [SupportByVersion("Access", 12, 14, 15, 16)]
        string StatusBarText { get; set; }

        /// <summary>
        /// SupportByVersion Access 12, 14, 15, 16
        /// Get/Set
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff197633.aspx </remarks>
        [SupportByVersion("Access", 12, 14, 15, 16)]
        bool TabStop { get; set; }

        /// <summary>
        /// SupportByVersion Access 12, 14, 15, 16
        /// Get/Set
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff822722.aspx </remarks>
        [SupportByVersion("Access", 12, 14, 15, 16)]
        Int16 TabIndex { get; set; }

        /// <summary>
        /// SupportByVersion Access 12, 14, 15, 16
        /// Get/Set
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff834720.aspx </remarks>
        [SupportByVersion("Access", 12, 14, 15, 16)]
        bool Enabled { get; set; }

        /// <summary>
        /// SupportByVersion Access 12, 14, 15, 16
        /// Get/Set
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff834789.aspx </remarks>
        [SupportByVersion("Access", 12, 14, 15, 16)]
        bool Locked { get; set; }

        /// <summary>
        /// SupportByVersion Access 12, 14, 15, 16
        /// Get
        /// </summary>
        [SupportByVersion("Access", 12, 14, 15, 16)]
        [EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
        new object FileData { get; }

        /// <summary>
        /// SupportByVersion Access 12, 14, 15, 16
        /// Get
        /// </summary>
        [SupportByVersion("Access", 12, 14, 15, 16)]
        [EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
        new object PictureDisp { get; }

        /// <summary>
        /// SupportByVersion Access 14, 15, 16
        /// Get/Set
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff823010.aspx </remarks>
        [SupportByVersion("Access", 14, 15, 16)]
        Int32 BackThemeColorIndex { get; set; }

        /// <summary>
        /// SupportByVersion Access 14, 15, 16
        /// Get/Set
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff836938.aspx </remarks>
        [SupportByVersion("Access", 14, 15, 16)]
        Single BackTint { get; set; }

        /// <summary>
        /// SupportByVersion Access 14, 15, 16
        /// Get/Set
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff191790.aspx </remarks>
        [SupportByVersion("Access", 14, 15, 16)]
        Single BackShade { get; set; }

        /// <summary>
        /// SupportByVersion Access 14, 15, 16
        /// Get/Set
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff820935.aspx </remarks>
        [SupportByVersion("Access", 14, 15, 16)]
        Int32 BorderThemeColorIndex { get; set; }

        /// <summary>
        /// SupportByVersion Access 14, 15, 16
        /// Get/Set
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff196496.aspx </remarks>
        [SupportByVersion("Access", 14, 15, 16)]
        Single BorderTint { get; set; }

        /// <summary>
        /// SupportByVersion Access 14, 15, 16
        /// Get/Set
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff192532.aspx </remarks>
        [SupportByVersion("Access", 14, 15, 16)]
        Single BorderShade { get; set; }

        /// <summary>
        /// SupportByVersion Access 14, 15, 16
        /// Get/Set
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff845642.aspx </remarks>
        [SupportByVersion("Access", 14, 15, 16)]
        Int32 GridlineThemeColorIndex { get; set; }

        /// <summary>
        /// SupportByVersion Access 14, 15, 16
        /// Get/Set
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff822861.aspx </remarks>
        [SupportByVersion("Access", 14, 15, 16)]
        Single GridlineTint { get; set; }

        /// <summary>
        /// SupportByVersion Access 14, 15, 16
        /// Get/Set
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff191829.aspx </remarks>
        [SupportByVersion("Access", 14, 15, 16)]
        Single GridlineShade { get; set; }

        /// <summary>
        /// SupportByVersion Access 14, 15, 16
        /// Get/Set
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff196043.aspx </remarks>
        [SupportByVersion("Access", 14, 15, 16)]
        byte DefaultPictureType { get; set; }

        #endregion

        #region Methods

        /// <summary>
        /// SupportByVersion Access 12, 14, 15, 16
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff198323.aspx </remarks>
        [SupportByVersion("Access", 12, 14, 15, 16)]
        void SizeToFit();

        /// <summary>
        /// SupportByVersion Access 12, 14, 15, 16
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff195435.aspx </remarks>
        [SupportByVersion("Access", 12, 14, 15, 16)]
        void Requery();

        /// <summary>
        /// SupportByVersion Access 12, 14, 15, 16
        /// </summary>
        [EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
        [SupportByVersion("Access", 12, 14, 15, 16)]
        void Goto();

        /// <summary>
        /// SupportByVersion Access 12, 14, 15, 16
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff194081.aspx </remarks>
        [SupportByVersion("Access", 12, 14, 15, 16)]
        void SetFocus();

        /// <summary>
        /// SupportByVersion Access 12, 14, 15, 16
        /// </summary>
        /// <param name="bstrExpr">string bstrExpr</param>
        /// <param name="ppsa">optional object[] ppsa</param>
        [EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
        [SupportByVersion("Access", 12, 14, 15, 16)]
        object _Evaluate(string bstrExpr, object[] ppsa);
		
		/// <summary>
		/// SupportByVersion Access 12, 14, 15, 16
		/// </summary>
		/// <param name="bstrExpr">string bstrExpr</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
        [CustomMethod]
        [SupportByVersion("Access", 12, 14, 15, 16)]
        object _Evaluate(string bstrExpr);

        /// <summary>
        /// SupportByVersion Access 12, 14, 15, 16
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff834482.aspx </remarks>
        /// <param name="left">object left</param>
        /// <param name="top">optional object top</param>
        /// <param name="width">optional object width</param>
        /// <param name="height">optional object height</param>
        [SupportByVersion("Access", 12, 14, 15, 16)]
        void Move(object left, object top, object width, object height);

        /// <summary>
        /// SupportByVersion Access 12, 14, 15, 16
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff834482.aspx </remarks>
        /// <param name="left">object left</param>
        [CustomMethod]
        [SupportByVersion("Access", 12, 14, 15, 16)]
        void Move(object left);

        /// <summary>
        /// SupportByVersion Access 12, 14, 15, 16
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff834482.aspx </remarks>
        /// <param name="left">object left</param>
        /// <param name="top">optional object top</param>
        [CustomMethod]
        [SupportByVersion("Access", 12, 14, 15, 16)]
        void Move(object left, object top);

        /// <summary>
        /// SupportByVersion Access 12, 14, 15, 16
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff834482.aspx </remarks>
        /// <param name="left">object left</param>
        /// <param name="top">optional object top</param>
        /// <param name="width">optional object width</param>
        [CustomMethod]
        [SupportByVersion("Access", 12, 14, 15, 16)]
        void Move(object left, object top, object width);

        /// <summary>
        /// SupportByVersion Access 12, 14, 15, 16
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff845316.aspx </remarks>
        [SupportByVersion("Access", 12, 14, 15, 16)]
        void Forward();

        /// <summary>
        /// SupportByVersion Access 12, 14, 15, 16
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff197753.aspx </remarks>
        [SupportByVersion("Access", 12, 14, 15, 16)]
        void Back();

        /// <summary>
        /// SupportByVersion Access 12, 14, 15, 16
        /// </summary>
        /// <param name="dispid">Int32 dispid</param>
        [EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
        [SupportByVersion("Access", 12, 14, 15, 16)]
        bool IsMemberSafe(Int32 dispid);

        #endregion
    } 
}
