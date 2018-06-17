using System;
using NetRuntimeSystem = System;
using System.ComponentModel;
using NetOffice.Attributes;

namespace NetOffice.AccessApi
{
	/// <summary>
	/// DispatchInterface _ListBox 
	/// SupportByVersion Access, 9,10,11,12,14,15,16
	/// </summary>
	[SupportByVersion("Access", 9,10,11,12,14,15,16)]
	[EntityType(EntityType.IsDispatchInterface), BaseType]
	[TypeId("3B06E95A-E47C-11CD-8701-00AA003F0F07")]
    [CoClassSource(typeof(NetOffice.AccessApi.ListBox))]
    public interface _ListBox : NetOffice.OfficeApi.IAccessible
	{
		#region Properties

		/// <summary>
		/// SupportByVersion Access 9, 10, 11, 12, 14, 15, 16
		/// Get
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff835405.aspx </remarks>
		[SupportByVersion("Access", 9,10,11,12,14,15,16)]
		NetOffice.AccessApi.Application Application { get; }

		/// <summary>
		/// SupportByVersion Access 9, 10, 11, 12, 14, 15, 16
		/// Get
		/// Unknown COM Proxy
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff823103.aspx </remarks>
		[SupportByVersion("Access", 9,10,11,12,14,15,16), ProxyResult]
		object Parent { get; }

		/// <summary>
		/// SupportByVersion Access 9, 10, 11, 12, 14, 15, 16
		/// Get
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff834793.aspx </remarks>
		/// <param name="index">Int32 index</param>
		/// <param name="row">optional object row</param>
		[SupportByVersion("Access", 9,10,11,12,14,15,16)]
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		object get_Column(Int32 index, object row);

		/// <summary>
		/// SupportByVersion Access 9, 10, 11, 12, 14, 15, 16
		/// Alias for get_Column
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff834793.aspx </remarks>
		/// <param name="index">Int32 index</param>
		/// <param name="row">optional object row</param>
		[SupportByVersion("Access", 9,10,11,12,14,15,16), Redirect("get_Column")]
		object Column(Int32 index, object row);

		/// <summary>
		/// SupportByVersion Access 9, 10, 11, 12, 14, 15, 16
		/// Get
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff834793.aspx </remarks>
		/// <param name="index">Int32 index</param>
		[SupportByVersion("Access", 9,10,11,12,14,15,16)]
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		object get_Column(Int32 index);

		/// <summary>
		/// SupportByVersion Access 9, 10, 11, 12, 14, 15, 16
		/// Alias for get_Column
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff834793.aspx </remarks>
		/// <param name="index">Int32 index</param>
		[SupportByVersion("Access", 9,10,11,12,14,15,16), Redirect("get_Column")]
		object Column(Int32 index);

		/// <summary>
		/// SupportByVersion Access 9, 10, 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff835338.aspx </remarks>
		/// <param name="lRow">Int32 lRow</param>
		[SupportByVersion("Access", 9,10,11,12,14,15,16)]
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		Int32 get_Selected(Int32 lRow);

        /// <summary>
        /// SupportByVersion Access 9, 10, 11, 12, 14, 15, 16
        /// Get/Set
        /// </summary>
        /// <param name="lRow">Int32 lRow</param>
        /// <param name="value">Int32 value</param>
        [SupportByVersion("Access", 9,10,11,12,14,15,16)]
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		void set_Selected(Int32 lRow, Int32 value);

		/// <summary>
		/// SupportByVersion Access 9, 10, 11, 12, 14, 15, 16
		/// Alias for get_Selected
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff835338.aspx </remarks>
		/// <param name="lRow">Int32 lRow</param>
		[SupportByVersion("Access", 9,10,11,12,14,15,16), Redirect("get_Selected")]
		Int32 Selected(Int32 lRow);

		/// <summary>
		/// SupportByVersion Access 9, 10, 11, 12, 14, 15, 16
		/// Get
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff194010.aspx </remarks>
		[SupportByVersion("Access", 9,10,11,12,14,15,16)]
		object OldValue { get; }

		/// <summary>
		/// SupportByVersion Access 9, 10, 11, 12, 14, 15, 16
		/// Get
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff820776.aspx </remarks>
		/// <param name="index">Int32 index</param>
		[SupportByVersion("Access", 9,10,11,12,14,15,16)]
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		object get_ItemData(Int32 index);

		/// <summary>
		/// SupportByVersion Access 9, 10, 11, 12, 14, 15, 16
		/// Alias for get_ItemData
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff820776.aspx </remarks>
		/// <param name="index">Int32 index</param>
		[SupportByVersion("Access", 9,10,11,12,14,15,16), Redirect("get_ItemData")]
		object ItemData(Int32 index);

		/// <summary>
		/// SupportByVersion Access 9, 10, 11, 12, 14, 15, 16
		/// Get
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff837302.aspx </remarks>
		[SupportByVersion("Access", 9,10,11,12,14,15,16)]
		NetOffice.AccessApi.Properties Properties { get; }

		/// <summary>
		/// SupportByVersion Access 9, 10, 11, 12, 14, 15, 16
		/// Get
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff845521.aspx </remarks>
		[SupportByVersion("Access", 9,10,11,12,14,15,16)]
		NetOffice.AccessApi.Children Controls { get; }

		/// <summary>
		/// SupportByVersion Access 9, 10, 11, 12, 14, 15, 16
		/// Get
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff823015.aspx </remarks>
		[SupportByVersion("Access", 9,10,11,12,14,15,16)]
		NetOffice.AccessApi._ItemsSelected ItemsSelected { get; }

		/// <summary>
		/// SupportByVersion Access 9, 10, 11, 12, 14, 15, 16
		/// Get
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff837056.aspx </remarks>
		[SupportByVersion("Access", 9,10,11,12,14,15,16)]
		[BaseResult]
		NetOffice.AccessApi._Hyperlink Hyperlink { get; }

		/// <summary>
		/// SupportByVersion Access 9, 10, 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff193151.aspx </remarks>
		[SupportByVersion("Access", 9,10,11,12,14,15,16)]
		object Value { get; set; }

		/// <summary>
		/// SupportByVersion Access 9, 10, 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff192015.aspx </remarks>
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
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff192266.aspx </remarks>
		[SupportByVersion("Access", 9,10,11,12,14,15,16)]
		byte ControlType { get; set; }

		/// <summary>
		/// SupportByVersion Access 9, 10, 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff192283.aspx </remarks>
		[SupportByVersion("Access", 9,10,11,12,14,15,16)]
		string ControlSource { get; set; }

		/// <summary>
		/// SupportByVersion Access 9, 10, 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff821440.aspx </remarks>
		[SupportByVersion("Access", 9,10,11,12,14,15,16)]
		string RowSourceType { get; set; }

		/// <summary>
		/// SupportByVersion Access 9, 10, 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff196460.aspx </remarks>
		[SupportByVersion("Access", 9,10,11,12,14,15,16)]
		string RowSource { get; set; }

		/// <summary>
		/// SupportByVersion Access 9, 10, 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff820820.aspx </remarks>
		[SupportByVersion("Access", 9,10,11,12,14,15,16)]
		Int16 ColumnCount { get; set; }

		/// <summary>
		/// SupportByVersion Access 9, 10, 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff834477.aspx </remarks>
		[SupportByVersion("Access", 9,10,11,12,14,15,16)]
		bool ColumnHeads { get; set; }

		/// <summary>
		/// SupportByVersion Access 9, 10, 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff193491.aspx </remarks>
		[SupportByVersion("Access", 9,10,11,12,14,15,16)]
		string ColumnWidths { get; set; }

		/// <summary>
		/// SupportByVersion Access 9, 10, 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff836882.aspx </remarks>
		[SupportByVersion("Access", 9,10,11,12,14,15,16)]
		Int32 BoundColumn { get; set; }

		/// <summary>
		/// SupportByVersion Access 9, 10, 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff835379.aspx </remarks>
		[SupportByVersion("Access", 9,10,11,12,14,15,16)]
		string DefaultValue { get; set; }

		/// <summary>
		/// SupportByVersion Access 9, 10, 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff191760.aspx </remarks>
		[SupportByVersion("Access", 9,10,11,12,14,15,16)]
		bool IMEHold { get; set; }

		/// <summary>
		/// SupportByVersion Access 9, 10, 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff834794.aspx </remarks>
		[SupportByVersion("Access", 9,10,11,12,14,15,16)]
		string ValidationRule { get; set; }

		/// <summary>
		/// SupportByVersion Access 9, 10, 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff822752.aspx </remarks>
		[SupportByVersion("Access", 9,10,11,12,14,15,16)]
		string ValidationText { get; set; }

		/// <summary>
		/// SupportByVersion Access 9, 10, 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff191982.aspx </remarks>
		[SupportByVersion("Access", 9,10,11,12,14,15,16)]
		string StatusBarText { get; set; }

		/// <summary>
		/// SupportByVersion Access 9, 10, 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff195987.aspx </remarks>
		[SupportByVersion("Access", 9,10,11,12,14,15,16)]
		bool Visible { get; set; }

		/// <summary>
		/// SupportByVersion Access 9, 10, 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff835089.aspx </remarks>
		[SupportByVersion("Access", 9,10,11,12,14,15,16)]
		byte DisplayWhen { get; set; }

		/// <summary>
		/// SupportByVersion Access 9, 10, 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff836886.aspx </remarks>
		[SupportByVersion("Access", 9,10,11,12,14,15,16)]
		bool Enabled { get; set; }

		/// <summary>
		/// SupportByVersion Access 9, 10, 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff822435.aspx </remarks>
		[SupportByVersion("Access", 9,10,11,12,14,15,16)]
		bool Locked { get; set; }

		/// <summary>
		/// SupportByVersion Access 9, 10, 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff195771.aspx </remarks>
		[SupportByVersion("Access", 9,10,11,12,14,15,16)]
		byte MultiSelect { get; set; }

		/// <summary>
		/// SupportByVersion Access 9, 10, 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff844922.aspx </remarks>
		[SupportByVersion("Access", 9,10,11,12,14,15,16)]
		bool TabStop { get; set; }

		/// <summary>
		/// SupportByVersion Access 9, 10, 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff193826.aspx </remarks>
		[SupportByVersion("Access", 9,10,11,12,14,15,16)]
		Int16 TabIndex { get; set; }

		/// <summary>
		/// SupportByVersion Access 9, 10, 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff821181.aspx </remarks>
		[SupportByVersion("Access", 9,10,11,12,14,15,16)]
		bool HideDuplicates { get; set; }

		/// <summary>
		/// SupportByVersion Access 9, 10, 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff192307.aspx </remarks>
		[SupportByVersion("Access", 9,10,11,12,14,15,16)]
		Int16 Left { get; set; }

		/// <summary>
		/// SupportByVersion Access 9, 10, 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff845686.aspx </remarks>
		[SupportByVersion("Access", 9,10,11,12,14,15,16)]
		Int16 Top { get; set; }

		/// <summary>
		/// SupportByVersion Access 9, 10, 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff192685.aspx </remarks>
		[SupportByVersion("Access", 9,10,11,12,14,15,16)]
		Int16 Width { get; set; }

		/// <summary>
		/// SupportByVersion Access 9, 10, 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff822454.aspx </remarks>
		[SupportByVersion("Access", 9,10,11,12,14,15,16)]
		Int16 Height { get; set; }

		/// <summary>
		/// SupportByVersion Access 9, 10, 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff196798.aspx </remarks>
		[SupportByVersion("Access", 9,10,11,12,14,15,16)]
		Int32 BackColor { get; set; }

		/// <summary>
		/// SupportByVersion Access 9, 10, 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff196380.aspx </remarks>
		[SupportByVersion("Access", 9,10,11,12,14,15,16)]
		byte SpecialEffect { get; set; }

		/// <summary>
		/// SupportByVersion Access 9, 10, 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff195459.aspx </remarks>
		[SupportByVersion("Access", 9,10,11,12,14,15,16)]
		byte BorderStyle { get; set; }

		/// <summary>
		/// SupportByVersion Access 9, 10, 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff193587.aspx </remarks>
		[SupportByVersion("Access", 9,10,11,12,14,15,16)]
		byte OldBorderStyle { get; set; }

		/// <summary>
		/// SupportByVersion Access 9, 10, 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff192725.aspx </remarks>
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
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff194153.aspx </remarks>
		[SupportByVersion("Access", 9,10,11,12,14,15,16)]
		Int32 BorderColor { get; set; }

		/// <summary>
		/// SupportByVersion Access 9, 10, 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff822710.aspx </remarks>
		[SupportByVersion("Access", 9,10,11,12,14,15,16)]
		Int32 ForeColor { get; set; }

		/// <summary>
		/// SupportByVersion Access 9, 10, 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff198062.aspx </remarks>
		[SupportByVersion("Access", 9,10,11,12,14,15,16)]
		string FontName { get; set; }

		/// <summary>
		/// SupportByVersion Access 9, 10, 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff193490.aspx </remarks>
		[SupportByVersion("Access", 9,10,11,12,14,15,16)]
		Int16 FontSize { get; set; }

		/// <summary>
		/// SupportByVersion Access 9, 10, 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff196370.aspx </remarks>
		[SupportByVersion("Access", 9,10,11,12,14,15,16)]
		Int16 FontWeight { get; set; }

		/// <summary>
		/// SupportByVersion Access 9, 10, 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff845203.aspx </remarks>
		[SupportByVersion("Access", 9,10,11,12,14,15,16)]
		bool FontItalic { get; set; }

		/// <summary>
		/// SupportByVersion Access 9, 10, 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff845713.aspx </remarks>
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
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff821795.aspx </remarks>
		[SupportByVersion("Access", 9,10,11,12,14,15,16)]
		Int16 FontBold { get; set; }

		/// <summary>
		/// SupportByVersion Access 9, 10, 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff194639.aspx </remarks>
		[SupportByVersion("Access", 9,10,11,12,14,15,16)]
		string ShortcutMenuBar { get; set; }

		/// <summary>
		/// SupportByVersion Access 9, 10, 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff191917.aspx </remarks>
		[SupportByVersion("Access", 9,10,11,12,14,15,16)]
		string ControlTipText { get; set; }

		/// <summary>
		/// SupportByVersion Access 9, 10, 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff835636.aspx </remarks>
		[SupportByVersion("Access", 9,10,11,12,14,15,16)]
		Int32 HelpContextId { get; set; }

		/// <summary>
		/// SupportByVersion Access 9, 10, 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff195887.aspx </remarks>
		[SupportByVersion("Access", 9,10,11,12,14,15,16)]
		Int16 ColumnWidth { get; set; }

		/// <summary>
		/// SupportByVersion Access 9, 10, 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff196766.aspx </remarks>
		[SupportByVersion("Access", 9,10,11,12,14,15,16)]
		Int16 ColumnOrder { get; set; }

		/// <summary>
		/// SupportByVersion Access 9, 10, 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff194134.aspx </remarks>
		[SupportByVersion("Access", 9,10,11,12,14,15,16)]
		bool ColumnHidden { get; set; }

		/// <summary>
		/// SupportByVersion Access 9, 10, 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff836855.aspx </remarks>
		[SupportByVersion("Access", 9,10,11,12,14,15,16)]
		bool AutoLabel { get; set; }

		/// <summary>
		/// SupportByVersion Access 9, 10, 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff836586.aspx </remarks>
		[SupportByVersion("Access", 9,10,11,12,14,15,16)]
		bool AddColon { get; set; }

		/// <summary>
		/// SupportByVersion Access 9, 10, 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff845394.aspx </remarks>
		[SupportByVersion("Access", 9,10,11,12,14,15,16)]
		Int16 LabelX { get; set; }

		/// <summary>
		/// SupportByVersion Access 9, 10, 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff196166.aspx </remarks>
		[SupportByVersion("Access", 9,10,11,12,14,15,16)]
		Int16 LabelY { get; set; }

		/// <summary>
		/// SupportByVersion Access 9, 10, 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff822055.aspx </remarks>
		[SupportByVersion("Access", 9,10,11,12,14,15,16)]
		byte LabelAlign { get; set; }

		/// <summary>
		/// SupportByVersion Access 9, 10, 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff192036.aspx </remarks>
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
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff196800.aspx </remarks>
		[SupportByVersion("Access", 9,10,11,12,14,15,16)]
		string Tag { get; set; }

		/// <summary>
		/// SupportByVersion Access 9, 10, 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff845049.aspx </remarks>
		[SupportByVersion("Access", 9,10,11,12,14,15,16)]
		Int32 ListCount { get; set; }

		/// <summary>
		/// SupportByVersion Access 9, 10, 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff193579.aspx </remarks>
		[SupportByVersion("Access", 9,10,11,12,14,15,16)]
		Int32 ListIndex { get; set; }

		/// <summary>
		/// SupportByVersion Access 9, 10, 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff194845.aspx </remarks>
		[SupportByVersion("Access", 9,10,11,12,14,15,16)]
		bool IsVisible { get; set; }

		/// <summary>
		/// SupportByVersion Access 9, 10, 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff193539.aspx </remarks>
		[SupportByVersion("Access", 9,10,11,12,14,15,16)]
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		bool InSelection { get; set; }

		/// <summary>
		/// SupportByVersion Access 9, 10, 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff822432.aspx </remarks>
		[SupportByVersion("Access", 9,10,11,12,14,15,16)]
		string BeforeUpdate { get; set; }

		/// <summary>
		/// SupportByVersion Access 9, 10, 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff822417.aspx </remarks>
		[SupportByVersion("Access", 9,10,11,12,14,15,16)]
		string AfterUpdate { get; set; }

		/// <summary>
		/// SupportByVersion Access 9, 10, 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff192523.aspx </remarks>
		[SupportByVersion("Access", 9,10,11,12,14,15,16)]
		string OnEnter { get; set; }

		/// <summary>
		/// SupportByVersion Access 9, 10, 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff197036.aspx </remarks>
		[SupportByVersion("Access", 9,10,11,12,14,15,16)]
		string OnExit { get; set; }

		/// <summary>
		/// SupportByVersion Access 9, 10, 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff835432.aspx </remarks>
		[SupportByVersion("Access", 9,10,11,12,14,15,16)]
		string OnGotFocus { get; set; }

		/// <summary>
		/// SupportByVersion Access 9, 10, 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff834495.aspx </remarks>
		[SupportByVersion("Access", 9,10,11,12,14,15,16)]
		string OnLostFocus { get; set; }

		/// <summary>
		/// SupportByVersion Access 9, 10, 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff845684.aspx </remarks>
		[SupportByVersion("Access", 9,10,11,12,14,15,16)]
		string OnClick { get; set; }

		/// <summary>
		/// SupportByVersion Access 9, 10, 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff192914.aspx </remarks>
		[SupportByVersion("Access", 9,10,11,12,14,15,16)]
		string OnDblClick { get; set; }

		/// <summary>
		/// SupportByVersion Access 9, 10, 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff191714.aspx </remarks>
		[SupportByVersion("Access", 9,10,11,12,14,15,16)]
		string OnMouseDown { get; set; }

		/// <summary>
		/// SupportByVersion Access 9, 10, 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff192063.aspx </remarks>
		[SupportByVersion("Access", 9,10,11,12,14,15,16)]
		string OnMouseMove { get; set; }

		/// <summary>
		/// SupportByVersion Access 9, 10, 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff820793.aspx </remarks>
		[SupportByVersion("Access", 9,10,11,12,14,15,16)]
		string OnMouseUp { get; set; }

		/// <summary>
		/// SupportByVersion Access 9, 10, 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff836570.aspx </remarks>
		[SupportByVersion("Access", 9,10,11,12,14,15,16)]
		string OnKeyDown { get; set; }

		/// <summary>
		/// SupportByVersion Access 9, 10, 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff836424.aspx </remarks>
		[SupportByVersion("Access", 9,10,11,12,14,15,16)]
		string OnKeyUp { get; set; }

		/// <summary>
		/// SupportByVersion Access 9, 10, 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff836928.aspx </remarks>
		[SupportByVersion("Access", 9,10,11,12,14,15,16)]
		string OnKeyPress { get; set; }

		/// <summary>
		/// SupportByVersion Access 9, 10, 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff198002.aspx </remarks>
		[SupportByVersion("Access", 9,10,11,12,14,15,16)]
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		byte ReadingOrder { get; set; }

		/// <summary>
		/// SupportByVersion Access 9, 10, 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff195567.aspx </remarks>
		[SupportByVersion("Access", 9,10,11,12,14,15,16)]
		byte ScrollBarAlign { get; set; }

		/// <summary>
		/// SupportByVersion Access 9, 10, 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		[SupportByVersion("Access", 9,10,11,12,14,15,16)]
		byte TextAlign { get; set; }

		/// <summary>
		/// SupportByVersion Access 9, 10, 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff822449.aspx </remarks>
		[SupportByVersion("Access", 9,10,11,12,14,15,16)]
		byte NumeralShapes { get; set; }

		/// <summary>
		/// SupportByVersion Access 9, 10, 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff196727.aspx </remarks>
		[SupportByVersion("Access", 9,10,11,12,14,15,16)]
		NetOffice.AccessApi.Enums.AcImeMode IMEMode { get; set; }

		/// <summary>
		/// SupportByVersion Access 9, 10, 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff196804.aspx </remarks>
		[SupportByVersion("Access", 9,10,11,12,14,15,16)]
		string Name { get; set; }

		/// <summary>
		/// SupportByVersion Access 9, 10, 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff197032.aspx </remarks>
		[SupportByVersion("Access", 9,10,11,12,14,15,16)]
		NetOffice.AccessApi.Enums.AcImeSentenceMode IMESentenceMode { get; set; }

		/// <summary>
		/// SupportByVersion Access 10, 11, 12, 14, 15, 16
		/// Get/Set
		/// Unknown COM Proxy
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff836245.aspx </remarks>
		[SupportByVersion("Access", 10,11,12,14,15,16), ProxyResult]
		object Recordset { get; set; }

		/// <summary>
		/// SupportByVersion Access 11, 12, 14, 15, 16
		/// Get
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff845835.aspx </remarks>
		[SupportByVersion("Access", 11,12,14,15,16)]
		[BaseResult]
		NetOffice.AccessApi._SmartTags SmartTags { get; }

		/// <summary>
		/// SupportByVersion Access 12, 14, 15, 16
		/// Get
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff845595.aspx </remarks>
		[SupportByVersion("Access", 12,14,15,16)]
		NetOffice.AccessApi.Enums.AcLayoutType Layout { get; }

		/// <summary>
		/// SupportByVersion Access 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff194962.aspx </remarks>
		[SupportByVersion("Access", 12,14,15,16)]
		Int16 LeftPadding { get; set; }

		/// <summary>
		/// SupportByVersion Access 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff822063.aspx </remarks>
		[SupportByVersion("Access", 12,14,15,16)]
		Int16 TopPadding { get; set; }

		/// <summary>
		/// SupportByVersion Access 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff821465.aspx </remarks>
		[SupportByVersion("Access", 12,14,15,16)]
		Int16 RightPadding { get; set; }

		/// <summary>
		/// SupportByVersion Access 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff196482.aspx </remarks>
		[SupportByVersion("Access", 12,14,15,16)]
		Int16 BottomPadding { get; set; }

		/// <summary>
		/// SupportByVersion Access 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff834334.aspx </remarks>
		[SupportByVersion("Access", 12,14,15,16)]
		byte GridlineStyleLeft { get; set; }

		/// <summary>
		/// SupportByVersion Access 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff835020.aspx </remarks>
		[SupportByVersion("Access", 12,14,15,16)]
		byte GridlineStyleTop { get; set; }

		/// <summary>
		/// SupportByVersion Access 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff193849.aspx </remarks>
		[SupportByVersion("Access", 12,14,15,16)]
		byte GridlineStyleRight { get; set; }

		/// <summary>
		/// SupportByVersion Access 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff845671.aspx </remarks>
		[SupportByVersion("Access", 12,14,15,16)]
		byte GridlineStyleBottom { get; set; }

		/// <summary>
		/// SupportByVersion Access 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff194863.aspx </remarks>
		[SupportByVersion("Access", 12,14,15,16)]
		byte GridlineWidthLeft { get; set; }

		/// <summary>
		/// SupportByVersion Access 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff822774.aspx </remarks>
		[SupportByVersion("Access", 12,14,15,16)]
		byte GridlineWidthTop { get; set; }

		/// <summary>
		/// SupportByVersion Access 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff821362.aspx </remarks>
		[SupportByVersion("Access", 12,14,15,16)]
		byte GridlineWidthRight { get; set; }

		/// <summary>
		/// SupportByVersion Access 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff822791.aspx </remarks>
		[SupportByVersion("Access", 12,14,15,16)]
		byte GridlineWidthBottom { get; set; }

		/// <summary>
		/// SupportByVersion Access 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff192033.aspx </remarks>
		[SupportByVersion("Access", 12,14,15,16)]
		Int32 GridlineColor { get; set; }

		/// <summary>
		/// SupportByVersion Access 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff836611.aspx </remarks>
		[SupportByVersion("Access", 12,14,15,16)]
		NetOffice.AccessApi.Enums.AcHorizontalAnchor HorizontalAnchor { get; set; }

		/// <summary>
		/// SupportByVersion Access 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff192092.aspx </remarks>
		[SupportByVersion("Access", 12,14,15,16)]
		NetOffice.AccessApi.Enums.AcVerticalAnchor VerticalAnchor { get; set; }

		/// <summary>
		/// SupportByVersion Access 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		[SupportByVersion("Access", 12,14,15,16)]
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		string BeforeUpdateMacro { get; set; }

		/// <summary>
		/// SupportByVersion Access 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		[SupportByVersion("Access", 12,14,15,16)]
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		string AfterUpdateMacro { get; set; }

		/// <summary>
		/// SupportByVersion Access 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		[SupportByVersion("Access", 12,14,15,16)]
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		string OnEnterMacro { get; set; }

		/// <summary>
		/// SupportByVersion Access 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		[SupportByVersion("Access", 12,14,15,16)]
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		string OnExitMacro { get; set; }

		/// <summary>
		/// SupportByVersion Access 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		[SupportByVersion("Access", 12,14,15,16)]
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		string OnGotFocusMacro { get; set; }

		/// <summary>
		/// SupportByVersion Access 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		[SupportByVersion("Access", 12,14,15,16)]
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		string OnLostFocusMacro { get; set; }

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
		/// Get/Set
		/// </summary>
		[SupportByVersion("Access", 12,14,15,16)]
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		string OnKeyDownMacro { get; set; }

		/// <summary>
		/// SupportByVersion Access 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		[SupportByVersion("Access", 12,14,15,16)]
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		string OnKeyUpMacro { get; set; }

		/// <summary>
		/// SupportByVersion Access 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		[SupportByVersion("Access", 12,14,15,16)]
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		string OnKeyPressMacro { get; set; }

		/// <summary>
		/// SupportByVersion Access 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff834354.aspx </remarks>
		[SupportByVersion("Access", 12,14,15,16)]
		bool AllowValueListEdits { get; set; }

		/// <summary>
		/// SupportByVersion Access 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff836908.aspx </remarks>
		[SupportByVersion("Access", 12,14,15,16)]
		string ListItemsEditForm { get; set; }

		/// <summary>
		/// SupportByVersion Access 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff196196.aspx </remarks>
		[SupportByVersion("Access", 12,14,15,16)]
		bool InheritValueList { get; set; }

		/// <summary>
		/// SupportByVersion Access 12, 14, 15, 16
		/// Get
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff845047.aspx </remarks>
		[SupportByVersion("Access", 12,14,15,16)]
		Int32 LayoutID { get; }

		/// <summary>
		/// SupportByVersion Access 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff192468.aspx </remarks>
		[SupportByVersion("Access", 12,14,15,16)]
		bool ShowOnlyRowSourceValues { get; set; }

		/// <summary>
		/// SupportByVersion Access 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff835060.aspx </remarks>
		[SupportByVersion("Access", 14,15,16)]
		Int32 BackThemeColorIndex { get; set; }

		/// <summary>
		/// SupportByVersion Access 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff196681.aspx </remarks>
		[SupportByVersion("Access", 14,15,16)]
		Single BackTint { get; set; }

		/// <summary>
		/// SupportByVersion Access 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff195189.aspx </remarks>
		[SupportByVersion("Access", 14,15,16)]
		Single BackShade { get; set; }

		/// <summary>
		/// SupportByVersion Access 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff192718.aspx </remarks>
		[SupportByVersion("Access", 14,15,16)]
		Int32 BorderThemeColorIndex { get; set; }

		/// <summary>
		/// SupportByVersion Access 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff196023.aspx </remarks>
		[SupportByVersion("Access", 14,15,16)]
		Single BorderTint { get; set; }

		/// <summary>
		/// SupportByVersion Access 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff836684.aspx </remarks>
		[SupportByVersion("Access", 14,15,16)]
		Single BorderShade { get; set; }

		/// <summary>
		/// SupportByVersion Access 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff197663.aspx </remarks>
		[SupportByVersion("Access", 14,15,16)]
		Int32 ForeThemeColorIndex { get; set; }

		/// <summary>
		/// SupportByVersion Access 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff192702.aspx </remarks>
		[SupportByVersion("Access", 14,15,16)]
		Single ForeTint { get; set; }

		/// <summary>
		/// SupportByVersion Access 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff193265.aspx </remarks>
		[SupportByVersion("Access", 14,15,16)]
		Single ForeShade { get; set; }

		/// <summary>
		/// SupportByVersion Access 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff196465.aspx </remarks>
		[SupportByVersion("Access", 14,15,16)]
		Int32 ThemeFontIndex { get; set; }

		/// <summary>
		/// SupportByVersion Access 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff194917.aspx </remarks>
		[SupportByVersion("Access", 14,15,16)]
		Int32 GridlineThemeColorIndex { get; set; }

		/// <summary>
		/// SupportByVersion Access 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff192876.aspx </remarks>
		[SupportByVersion("Access", 14,15,16)]
		Single GridlineTint { get; set; }

		/// <summary>
		/// SupportByVersion Access 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff191885.aspx </remarks>
		[SupportByVersion("Access", 14,15,16)]
		Single GridlineShade { get; set; }

		#endregion

		#region Methods

		/// <summary>
		/// SupportByVersion Access 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff836296.aspx </remarks>
		[SupportByVersion("Access", 9,10,11,12,14,15,16)]
		void Undo();

		/// <summary>
		/// SupportByVersion Access 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff845136.aspx </remarks>
		[SupportByVersion("Access", 9,10,11,12,14,15,16)]
		void SizeToFit();

		/// <summary>
		/// SupportByVersion Access 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff197359.aspx </remarks>
		[SupportByVersion("Access", 9,10,11,12,14,15,16)]
		void Requery();

		/// <summary>
		/// SupportByVersion Access 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		[SupportByVersion("Access", 9,10,11,12,14,15,16)]
		void Goto();

		/// <summary>
		/// SupportByVersion Access 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff821412.aspx </remarks>
		[SupportByVersion("Access", 9,10,11,12,14,15,16)]
		void SetFocus();

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
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff820814.aspx </remarks>
		/// <param name="left">object left</param>
		/// <param name="top">optional object top</param>
		/// <param name="width">optional object width</param>
		/// <param name="height">optional object height</param>
		[SupportByVersion("Access", 10,11,12,14,15,16)]
		void Move(object left, object top, object width, object height);

		/// <summary>
		/// SupportByVersion Access 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff820814.aspx </remarks>
		/// <param name="left">object left</param>
		[CustomMethod]
		[SupportByVersion("Access", 10,11,12,14,15,16)]
		void Move(object left);

		/// <summary>
		/// SupportByVersion Access 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff820814.aspx </remarks>
		/// <param name="left">object left</param>
		/// <param name="top">optional object top</param>
		[CustomMethod]
		[SupportByVersion("Access", 10,11,12,14,15,16)]
		void Move(object left, object top);

		/// <summary>
		/// SupportByVersion Access 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff820814.aspx </remarks>
		/// <param name="left">object left</param>
		/// <param name="top">optional object top</param>
		/// <param name="width">optional object width</param>
		[CustomMethod]
		[SupportByVersion("Access", 10,11,12,14,15,16)]
		void Move(object left, object top, object width);

		/// <summary>
		/// SupportByVersion Access 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff835328.aspx </remarks>
		/// <param name="item">string item</param>
		/// <param name="index">optional object index</param>
		[SupportByVersion("Access", 10,11,12,14,15,16)]
		void AddItem(string item, object index);

		/// <summary>
		/// SupportByVersion Access 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff835328.aspx </remarks>
		/// <param name="item">string item</param>
		[CustomMethod]
		[SupportByVersion("Access", 10,11,12,14,15,16)]
		void AddItem(string item);

		/// <summary>
		/// SupportByVersion Access 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff194439.aspx </remarks>
		/// <param name="index">object index</param>
		[SupportByVersion("Access", 10,11,12,14,15,16)]
		void RemoveItem(object index);

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
