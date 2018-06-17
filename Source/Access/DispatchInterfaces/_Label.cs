using System;
using NetRuntimeSystem = System;
using System.ComponentModel;
using NetOffice.Attributes;

namespace NetOffice.AccessApi
{
	/// <summary>
	/// DispatchInterface _Label 
	/// SupportByVersion Access, 9,10,11,12,14,15,16
	/// </summary>
	[SupportByVersion("Access", 9,10,11,12,14,15,16)]
	[EntityType(EntityType.IsDispatchInterface), BaseType]
	[TypeId("3B06E948-E47C-11CD-8701-00AA003F0F07")]
    [CoClassSource(typeof(NetOffice.AccessApi._ChildLabel), typeof(NetOffice.AccessApi.Label))]
    public interface _Label : NetOffice.OfficeApi.IAccessible
	{
		#region Properties

		/// <summary>
		/// SupportByVersion Access 9, 10, 11, 12, 14, 15, 16
		/// Get
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff197732.aspx </remarks>
		[SupportByVersion("Access", 9,10,11,12,14,15,16)]
		NetOffice.AccessApi.Application Application { get; }

		/// <summary>
		/// SupportByVersion Access 9, 10, 11, 12, 14, 15, 16
		/// Get
		/// Unknown COM Proxy
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff197063.aspx </remarks>
		[SupportByVersion("Access", 9,10,11,12,14,15,16), ProxyResult]
		object Parent { get; }

		/// <summary>
		/// SupportByVersion Access 9, 10, 11, 12, 14, 15, 16
		/// Get
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff192695.aspx </remarks>
		[SupportByVersion("Access", 9,10,11,12,14,15,16)]
		NetOffice.AccessApi.Properties Properties { get; }

		/// <summary>
		/// SupportByVersion Access 9, 10, 11, 12, 14, 15, 16
		/// Get
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff836961.aspx </remarks>
		[SupportByVersion("Access", 9,10,11,12,14,15,16)]
		[BaseResult]
		NetOffice.AccessApi._Hyperlink Hyperlink { get; }

		/// <summary>
		/// SupportByVersion Access 9, 10, 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff845019.aspx </remarks>
		[SupportByVersion("Access", 9,10,11,12,14,15,16)]
		string EventProcPrefix { get; set; }

		/// <summary>
		/// SupportByVersion Access 9, 10, 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		[SupportByVersion("Access", 9,10,11,12,14,15,16)]
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		string _Name { get; set; }

		/// <summary>
		/// SupportByVersion Access 9, 10, 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff836622.aspx </remarks>
		[SupportByVersion("Access", 9,10,11,12,14,15,16)]
		byte ControlType { get; set; }

		/// <summary>
		/// SupportByVersion Access 9, 10, 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff193237.aspx </remarks>
		[SupportByVersion("Access", 9,10,11,12,14,15,16)]
		string Caption { get; set; }

		/// <summary>
		/// SupportByVersion Access 9, 10, 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff836335.aspx </remarks>
		[SupportByVersion("Access", 9,10,11,12,14,15,16)]
		string HyperlinkAddress { get; set; }

		/// <summary>
		/// SupportByVersion Access 9, 10, 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff192850.aspx </remarks>
		[SupportByVersion("Access", 9,10,11,12,14,15,16)]
		string HyperlinkSubAddress { get; set; }

		/// <summary>
		/// SupportByVersion Access 9, 10, 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff822738.aspx </remarks>
		[SupportByVersion("Access", 9,10,11,12,14,15,16)]
		bool Visible { get; set; }

		/// <summary>
		/// SupportByVersion Access 9, 10, 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff836078.aspx </remarks>
		[SupportByVersion("Access", 9,10,11,12,14,15,16)]
		byte DisplayWhen { get; set; }

		/// <summary>
		/// SupportByVersion Access 9, 10, 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff195518.aspx </remarks>
		[SupportByVersion("Access", 9,10,11,12,14,15,16)]
		bool Vertical { get; set; }

		/// <summary>
		/// SupportByVersion Access 9, 10, 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff822023.aspx </remarks>
		[SupportByVersion("Access", 9,10,11,12,14,15,16)]
		Int16 Left { get; set; }

		/// <summary>
		/// SupportByVersion Access 9, 10, 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff823200.aspx </remarks>
		[SupportByVersion("Access", 9,10,11,12,14,15,16)]
		Int16 Top { get; set; }

		/// <summary>
		/// SupportByVersion Access 9, 10, 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff821109.aspx </remarks>
		[SupportByVersion("Access", 9,10,11,12,14,15,16)]
		Int16 Width { get; set; }

		/// <summary>
		/// SupportByVersion Access 9, 10, 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff193454.aspx </remarks>
		[SupportByVersion("Access", 9,10,11,12,14,15,16)]
		Int16 Height { get; set; }

		/// <summary>
		/// SupportByVersion Access 9, 10, 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff191911.aspx </remarks>
		[SupportByVersion("Access", 9,10,11,12,14,15,16)]
		byte BackStyle { get; set; }

		/// <summary>
		/// SupportByVersion Access 9, 10, 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff822069.aspx </remarks>
		[SupportByVersion("Access", 9,10,11,12,14,15,16)]
		Int32 BackColor { get; set; }

		/// <summary>
		/// SupportByVersion Access 9, 10, 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff192098.aspx </remarks>
		[SupportByVersion("Access", 9,10,11,12,14,15,16)]
		byte SpecialEffect { get; set; }

		/// <summary>
		/// SupportByVersion Access 9, 10, 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff194644.aspx </remarks>
		[SupportByVersion("Access", 9,10,11,12,14,15,16)]
		byte BorderStyle { get; set; }

		/// <summary>
		/// SupportByVersion Access 9, 10, 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff821073.aspx </remarks>
		[SupportByVersion("Access", 9,10,11,12,14,15,16)]
		byte OldBorderStyle { get; set; }

		/// <summary>
		/// SupportByVersion Access 9, 10, 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff845367.aspx </remarks>
		[SupportByVersion("Access", 9,10,11,12,14,15,16)]
		Int32 BorderColor { get; set; }

		/// <summary>
		/// SupportByVersion Access 9, 10, 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff191816.aspx </remarks>
		[SupportByVersion("Access", 9,10,11,12,14,15,16)]
		byte BorderWidth { get; set; }

		/// <summary>
		/// SupportByVersion Access 9, 10, 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		[SupportByVersion("Access", 9,10,11,12,14,15,16)]
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		byte BorderLineStyle { get; set; }

		/// <summary>
		/// SupportByVersion Access 9, 10, 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff835355.aspx </remarks>
		[SupportByVersion("Access", 9,10,11,12,14,15,16)]
		Int32 ForeColor { get; set; }

		/// <summary>
		/// SupportByVersion Access 9, 10, 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff192083.aspx </remarks>
		[SupportByVersion("Access", 9,10,11,12,14,15,16)]
		string FontName { get; set; }

		/// <summary>
		/// SupportByVersion Access 9, 10, 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff198044.aspx </remarks>
		[SupportByVersion("Access", 9,10,11,12,14,15,16)]
		Int16 FontSize { get; set; }

		/// <summary>
		/// SupportByVersion Access 9, 10, 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff196716.aspx </remarks>
		[SupportByVersion("Access", 9,10,11,12,14,15,16)]
		Int16 FontWeight { get; set; }

		/// <summary>
		/// SupportByVersion Access 9, 10, 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff196449.aspx </remarks>
		[SupportByVersion("Access", 9,10,11,12,14,15,16)]
		bool FontItalic { get; set; }

		/// <summary>
		/// SupportByVersion Access 9, 10, 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff845189.aspx </remarks>
		[SupportByVersion("Access", 9,10,11,12,14,15,16)]
		bool FontUnderline { get; set; }

		/// <summary>
		/// SupportByVersion Access 9, 10, 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		[SupportByVersion("Access", 9,10,11,12,14,15,16)]
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		byte TextFontCharSet { get; set; }

		/// <summary>
		/// SupportByVersion Access 9, 10, 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff845014.aspx </remarks>
		[SupportByVersion("Access", 9,10,11,12,14,15,16)]
		byte TextAlign { get; set; }

		/// <summary>
		/// SupportByVersion Access 9, 10, 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff197689.aspx </remarks>
		[SupportByVersion("Access", 9,10,11,12,14,15,16)]
		Int16 FontBold { get; set; }

		/// <summary>
		/// SupportByVersion Access 9, 10, 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff192886.aspx </remarks>
		[SupportByVersion("Access", 9,10,11,12,14,15,16)]
		string ControlTipText { get; set; }

		/// <summary>
		/// SupportByVersion Access 9, 10, 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff837156.aspx </remarks>
		[SupportByVersion("Access", 9,10,11,12,14,15,16)]
		Int32 HelpContextId { get; set; }

		/// <summary>
		/// SupportByVersion Access 9, 10, 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff845271.aspx </remarks>
		[SupportByVersion("Access", 9,10,11,12,14,15,16)]
		Int16 Section { get; set; }

		/// <summary>
		/// SupportByVersion Access 9, 10, 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		[SupportByVersion("Access", 9,10,11,12,14,15,16)]
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		string ControlName { get; set; }

		/// <summary>
		/// SupportByVersion Access 9, 10, 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff845369.aspx </remarks>
		[SupportByVersion("Access", 9,10,11,12,14,15,16)]
		bool IsVisible { get; set; }

		/// <summary>
		/// SupportByVersion Access 9, 10, 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff197791.aspx </remarks>
		[SupportByVersion("Access", 9,10,11,12,14,15,16)]
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		bool InSelection { get; set; }

		/// <summary>
		/// SupportByVersion Access 9, 10, 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff821103.aspx </remarks>
		[SupportByVersion("Access", 9,10,11,12,14,15,16)]
		string OnClick { get; set; }

		/// <summary>
		/// SupportByVersion Access 9, 10, 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff845658.aspx </remarks>
		[SupportByVersion("Access", 9,10,11,12,14,15,16)]
		string OnDblClick { get; set; }

		/// <summary>
		/// SupportByVersion Access 9, 10, 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff837304.aspx </remarks>
		[SupportByVersion("Access", 9,10,11,12,14,15,16)]
		string OnMouseDown { get; set; }

		/// <summary>
		/// SupportByVersion Access 9, 10, 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff192295.aspx </remarks>
		[SupportByVersion("Access", 9,10,11,12,14,15,16)]
		string OnMouseMove { get; set; }

		/// <summary>
		/// SupportByVersion Access 9, 10, 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff845384.aspx </remarks>
		[SupportByVersion("Access", 9,10,11,12,14,15,16)]
		string OnMouseUp { get; set; }

		/// <summary>
		/// SupportByVersion Access 9, 10, 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff191854.aspx </remarks>
		[SupportByVersion("Access", 9,10,11,12,14,15,16)]
		string ShortcutMenuBar { get; set; }

		/// <summary>
		/// SupportByVersion Access 9, 10, 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff844941.aspx </remarks>
		[SupportByVersion("Access", 9,10,11,12,14,15,16)]
		string Tag { get; set; }

		/// <summary>
		/// SupportByVersion Access 9, 10, 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff836884.aspx </remarks>
		[SupportByVersion("Access", 9,10,11,12,14,15,16)]
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		byte ReadingOrder { get; set; }

		/// <summary>
		/// SupportByVersion Access 9, 10, 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff192715.aspx </remarks>
		[SupportByVersion("Access", 9,10,11,12,14,15,16)]
		byte NumeralShapes { get; set; }

		/// <summary>
		/// SupportByVersion Access 9, 10, 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff194609.aspx </remarks>
		[SupportByVersion("Access", 9,10,11,12,14,15,16)]
		string Name { get; set; }

		/// <summary>
		/// SupportByVersion Access 9, 10, 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff196451.aspx </remarks>
		[SupportByVersion("Access", 9,10,11,12,14,15,16)]
		Int16 LeftMargin { get; set; }

		/// <summary>
		/// SupportByVersion Access 9, 10, 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff197691.aspx </remarks>
		[SupportByVersion("Access", 9,10,11,12,14,15,16)]
		Int16 TopMargin { get; set; }

		/// <summary>
		/// SupportByVersion Access 9, 10, 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff197026.aspx </remarks>
		[SupportByVersion("Access", 9,10,11,12,14,15,16)]
		Int16 LineSpacing { get; set; }

		/// <summary>
		/// SupportByVersion Access 9, 10, 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff844807.aspx </remarks>
		[SupportByVersion("Access", 9,10,11,12,14,15,16)]
		Int16 RightMargin { get; set; }

		/// <summary>
		/// SupportByVersion Access 9, 10, 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff845191.aspx </remarks>
		[SupportByVersion("Access", 9,10,11,12,14,15,16)]
		Int16 BottomMargin { get; set; }

		/// <summary>
		/// SupportByVersion Access 11, 12, 14, 15, 16
		/// Get
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff845730.aspx </remarks>
		[SupportByVersion("Access", 11,12,14,15,16)]
		[BaseResult]
		NetOffice.AccessApi._SmartTags SmartTags { get; }

		/// <summary>
		/// SupportByVersion Access 12, 14, 15, 16
		/// Get
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff836377.aspx </remarks>
		[SupportByVersion("Access", 12,14,15,16)]
		NetOffice.AccessApi.Enums.AcLayoutType Layout { get; }

		/// <summary>
		/// SupportByVersion Access 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff835335.aspx </remarks>
		[SupportByVersion("Access", 12,14,15,16)]
		Int16 LeftPadding { get; set; }

		/// <summary>
		/// SupportByVersion Access 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff197935.aspx </remarks>
		[SupportByVersion("Access", 12,14,15,16)]
		Int16 TopPadding { get; set; }

		/// <summary>
		/// SupportByVersion Access 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff194934.aspx </remarks>
		[SupportByVersion("Access", 12,14,15,16)]
		Int16 RightPadding { get; set; }

		/// <summary>
		/// SupportByVersion Access 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff845886.aspx </remarks>
		[SupportByVersion("Access", 12,14,15,16)]
		Int16 BottomPadding { get; set; }

		/// <summary>
		/// SupportByVersion Access 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff822711.aspx </remarks>
		[SupportByVersion("Access", 12,14,15,16)]
		byte GridlineStyleLeft { get; set; }

		/// <summary>
		/// SupportByVersion Access 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff192244.aspx </remarks>
		[SupportByVersion("Access", 12,14,15,16)]
		byte GridlineStyleTop { get; set; }

		/// <summary>
		/// SupportByVersion Access 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff196672.aspx </remarks>
		[SupportByVersion("Access", 12,14,15,16)]
		byte GridlineStyleRight { get; set; }

		/// <summary>
		/// SupportByVersion Access 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff845470.aspx </remarks>
		[SupportByVersion("Access", 12,14,15,16)]
		byte GridlineStyleBottom { get; set; }

		/// <summary>
		/// SupportByVersion Access 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff835639.aspx </remarks>
		[SupportByVersion("Access", 12,14,15,16)]
		byte GridlineWidthLeft { get; set; }

		/// <summary>
		/// SupportByVersion Access 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff820758.aspx </remarks>
		[SupportByVersion("Access", 12,14,15,16)]
		byte GridlineWidthTop { get; set; }

		/// <summary>
		/// SupportByVersion Access 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff835756.aspx </remarks>
		[SupportByVersion("Access", 12,14,15,16)]
		byte GridlineWidthRight { get; set; }

		/// <summary>
		/// SupportByVersion Access 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff822404.aspx </remarks>
		[SupportByVersion("Access", 12,14,15,16)]
		byte GridlineWidthBottom { get; set; }

		/// <summary>
		/// SupportByVersion Access 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff836879.aspx </remarks>
		[SupportByVersion("Access", 12,14,15,16)]
		Int32 GridlineColor { get; set; }

		/// <summary>
		/// SupportByVersion Access 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff834433.aspx </remarks>
		[SupportByVersion("Access", 12,14,15,16)]
		NetOffice.AccessApi.Enums.AcHorizontalAnchor HorizontalAnchor { get; set; }

		/// <summary>
		/// SupportByVersion Access 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff194232.aspx </remarks>
		[SupportByVersion("Access", 12,14,15,16)]
		NetOffice.AccessApi.Enums.AcVerticalAnchor VerticalAnchor { get; set; }

		/// <summary>
		/// SupportByVersion Access 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		[SupportByVersion("Access", 12,14,15,16)]
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		string OnClickMacro { get; set; }

		/// <summary>
		/// SupportByVersion Access 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		[SupportByVersion("Access", 12,14,15,16)]
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		string OnDblClickMacro { get; set; }

		/// <summary>
		/// SupportByVersion Access 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		[SupportByVersion("Access", 12,14,15,16)]
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		string OnMouseDownMacro { get; set; }

		/// <summary>
		/// SupportByVersion Access 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		[SupportByVersion("Access", 12,14,15,16)]
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		string OnMouseMoveMacro { get; set; }

		/// <summary>
		/// SupportByVersion Access 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		[SupportByVersion("Access", 12,14,15,16)]
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		string OnMouseUpMacro { get; set; }

		/// <summary>
		/// SupportByVersion Access 12, 14, 15, 16
		/// Get
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff845247.aspx </remarks>
		[SupportByVersion("Access", 12,14,15,16)]
		Int32 LayoutID { get; }

		/// <summary>
		/// SupportByVersion Access 14, 15, 16
		/// Get/Set
		/// </summary>
		[SupportByVersion("Access", 14,15,16)]
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		string Target { get; set; }

		/// <summary>
		/// SupportByVersion Access 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff196483.aspx </remarks>
		[SupportByVersion("Access", 14,15,16)]
		Int32 BackThemeColorIndex { get; set; }

		/// <summary>
		/// SupportByVersion Access 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff194959.aspx </remarks>
		[SupportByVersion("Access", 14,15,16)]
		Single BackTint { get; set; }

		/// <summary>
		/// SupportByVersion Access 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff191762.aspx </remarks>
		[SupportByVersion("Access", 14,15,16)]
		Single BackShade { get; set; }

		/// <summary>
		/// SupportByVersion Access 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff836009.aspx </remarks>
		[SupportByVersion("Access", 14,15,16)]
		Int32 BorderThemeColorIndex { get; set; }

		/// <summary>
		/// SupportByVersion Access 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff196389.aspx </remarks>
		[SupportByVersion("Access", 14,15,16)]
		Single BorderTint { get; set; }

		/// <summary>
		/// SupportByVersion Access 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff194952.aspx </remarks>
		[SupportByVersion("Access", 14,15,16)]
		Single BorderShade { get; set; }

		/// <summary>
		/// SupportByVersion Access 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff195247.aspx </remarks>
		[SupportByVersion("Access", 14,15,16)]
		Int32 ForeThemeColorIndex { get; set; }

		/// <summary>
		/// SupportByVersion Access 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff834437.aspx </remarks>
		[SupportByVersion("Access", 14,15,16)]
		Single ForeTint { get; set; }

		/// <summary>
		/// SupportByVersion Access 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff193557.aspx </remarks>
		[SupportByVersion("Access", 14,15,16)]
		Single ForeShade { get; set; }

		/// <summary>
		/// SupportByVersion Access 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff194239.aspx </remarks>
		[SupportByVersion("Access", 14,15,16)]
		Int32 ThemeFontIndex { get; set; }

		/// <summary>
		/// SupportByVersion Access 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff192688.aspx </remarks>
		[SupportByVersion("Access", 14,15,16)]
		Int32 GridlineThemeColorIndex { get; set; }

		/// <summary>
		/// SupportByVersion Access 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff192843.aspx </remarks>
		[SupportByVersion("Access", 14,15,16)]
		Single GridlineTint { get; set; }

		/// <summary>
		/// SupportByVersion Access 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff834779.aspx </remarks>
		[SupportByVersion("Access", 14,15,16)]
		Single GridlineShade { get; set; }

		#endregion

		#region Methods

		/// <summary>
		/// SupportByVersion Access 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff837020.aspx </remarks>
		[SupportByVersion("Access", 9,10,11,12,14,15,16)]
		void SizeToFit();

		/// <summary>
		/// SupportByVersion Access 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <param name="bstrExpr">string bstrExpr</param>
		/// <param name="ppsa">optional object[] ppsa</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		[SupportByVersion("Access", 9,10,11,12,14,15,16)]
		object _Evaluate(string bstrExpr, object[] ppsa);

		/// <summary>
		/// SupportByVersion Access 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <param name="bstrExpr">string bstrExpr</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		[CustomMethod]
		[SupportByVersion("Access", 9,10,11,12,14,15,16)]
		object _Evaluate(string bstrExpr);

		/// <summary>
		/// SupportByVersion Access 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff845043.aspx </remarks>
		/// <param name="left">object left</param>
		/// <param name="top">optional object top</param>
		/// <param name="width">optional object width</param>
		/// <param name="height">optional object height</param>
		[SupportByVersion("Access", 10,11,12,14,15,16)]
		void Move(object left, object top, object width, object height);

		/// <summary>
		/// SupportByVersion Access 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff845043.aspx </remarks>
		/// <param name="left">object left</param>
		[CustomMethod]
		[SupportByVersion("Access", 10,11,12,14,15,16)]
		void Move(object left);

		/// <summary>
		/// SupportByVersion Access 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff845043.aspx </remarks>
		/// <param name="left">object left</param>
		/// <param name="top">optional object top</param>
		[CustomMethod]
		[SupportByVersion("Access", 10,11,12,14,15,16)]
		void Move(object left, object top);

		/// <summary>
		/// SupportByVersion Access 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff845043.aspx </remarks>
		/// <param name="left">object left</param>
		/// <param name="top">optional object top</param>
		/// <param name="width">optional object width</param>
		[CustomMethod]
		[SupportByVersion("Access", 10,11,12,14,15,16)]
		void Move(object left, object top, object width);

		/// <summary>
		/// SupportByVersion Access 11, 12, 14, 15, 16
		/// </summary>
		/// <param name="dispid">Int32 dispid</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		[SupportByVersion("Access", 11,12,14,15,16)]
		bool IsMemberSafe(Int32 dispid);

		#endregion
	}
}
