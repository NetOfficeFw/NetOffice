using System;
using NetRuntimeSystem = System;
using System.ComponentModel;
using NetOffice.Attributes;

namespace NetOffice.AccessApi
{
	/// <summary>
	/// DispatchInterface _ObjectFrame 
	/// SupportByVersion Access, 9,10,11,12,14,15,16
	/// </summary>
	[SupportByVersion("Access", 9,10,11,12,14,15,16)]
	[EntityType(EntityType.IsDispatchInterface), BaseType]
	[TypeId("3B06E95E-E47C-11CD-8701-00AA003F0F07")]
    [CoClassSource(typeof(NetOffice.AccessApi.ObjectFrame))]
    public interface _ObjectFrame : NetOffice.OfficeApi.IAccessible
	{
		#region Properties

		/// <summary>
		/// SupportByVersion Access 9, 10, 11, 12, 14, 15, 16
		/// Get
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff192054.aspx </remarks>
		[SupportByVersion("Access", 9,10,11,12,14,15,16)]
		NetOffice.AccessApi.Application Application { get; }

		/// <summary>
		/// SupportByVersion Access 9, 10, 11, 12, 14, 15, 16
		/// Get
		/// Unknown COM Proxy
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff193535.aspx </remarks>
		[SupportByVersion("Access", 9,10,11,12,14,15,16), ProxyResult]
		object Parent { get; }

		/// <summary>
		/// SupportByVersion Access 9, 10, 11, 12, 14, 15, 16
		/// Get
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff836395.aspx </remarks>
		[SupportByVersion("Access", 9,10,11,12,14,15,16)]
		object OldValue { get; }

		/// <summary>
		/// SupportByVersion Access 9, 10, 11, 12, 14, 15, 16
		/// Get
		/// Unknown COM Proxy
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff835340.aspx </remarks>
		[SupportByVersion("Access", 9,10,11,12,14,15,16), ProxyResult]
		object Object { get; }

		/// <summary>
		/// SupportByVersion Access 9, 10, 11, 12, 14, 15, 16
		/// Get
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff835659.aspx </remarks>
		/// <param name="index">Int32 index</param>
		[SupportByVersion("Access", 9,10,11,12,14,15,16)]
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		string get_ObjectVerbs(Int32 index);

		/// <summary>
		/// SupportByVersion Access 9, 10, 11, 12, 14, 15, 16
		/// Alias for get_ObjectVerbs
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff835659.aspx </remarks>
		/// <param name="index">Int32 index</param>
		[SupportByVersion("Access", 9,10,11,12,14,15,16), Redirect("get_ObjectVerbs")]
		string ObjectVerbs(Int32 index);

		/// <summary>
		/// SupportByVersion Access 9, 10, 11, 12, 14, 15, 16
		/// Get
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff835069.aspx </remarks>
		[SupportByVersion("Access", 9,10,11,12,14,15,16)]
		NetOffice.AccessApi.Properties Properties { get; }

		/// <summary>
		/// SupportByVersion Access 9, 10, 11, 12, 14, 15, 16
		/// Get
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff192915.aspx </remarks>
		[SupportByVersion("Access", 9,10,11,12,14,15,16)]
		NetOffice.AccessApi.Children Controls { get; }

		/// <summary>
		/// SupportByVersion Access 9, 10, 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff821045.aspx </remarks>
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
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff192655.aspx </remarks>
		[SupportByVersion("Access", 9,10,11,12,14,15,16)]
		byte ControlType { get; set; }

		/// <summary>
		/// SupportByVersion Access 9, 10, 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff192070.aspx </remarks>
		[SupportByVersion("Access", 9,10,11,12,14,15,16)]
		byte SizeMode { get; set; }

		/// <summary>
		/// SupportByVersion Access 9, 10, 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff836327.aspx </remarks>
		[SupportByVersion("Access", 9,10,11,12,14,15,16)]
		string OLEClass { get; set; }

		/// <summary>
		/// SupportByVersion Access 9, 10, 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff845599.aspx </remarks>
		[SupportByVersion("Access", 9,10,11,12,14,15,16)]
		string Item { get; set; }

		/// <summary>
		/// SupportByVersion Access 9, 10, 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff834811.aspx </remarks>
		[SupportByVersion("Access", 9,10,11,12,14,15,16)]
		string RowSourceType { get; set; }

		/// <summary>
		/// SupportByVersion Access 9, 10, 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff835403.aspx </remarks>
		[SupportByVersion("Access", 9,10,11,12,14,15,16)]
		string RowSource { get; set; }

		/// <summary>
		/// SupportByVersion Access 9, 10, 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff836940.aspx </remarks>
		[SupportByVersion("Access", 9,10,11,12,14,15,16)]
		string LinkChildFields { get; set; }

		/// <summary>
		/// SupportByVersion Access 9, 10, 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff845815.aspx </remarks>
		[SupportByVersion("Access", 9,10,11,12,14,15,16)]
		string LinkMasterFields { get; set; }

		/// <summary>
		/// SupportByVersion Access 9, 10, 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff835991.aspx </remarks>
		[SupportByVersion("Access", 9,10,11,12,14,15,16)]
		Int16 AutoActivate { get; set; }

		/// <summary>
		/// SupportByVersion Access 9, 10, 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff192277.aspx </remarks>
		[SupportByVersion("Access", 9,10,11,12,14,15,16)]
		bool DisplayType { get; set; }

		/// <summary>
		/// SupportByVersion Access 9, 10, 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff192058.aspx </remarks>
		[SupportByVersion("Access", 9,10,11,12,14,15,16)]
		Int16 UpdateOptions { get; set; }

		/// <summary>
		/// SupportByVersion Access 9, 10, 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff836339.aspx </remarks>
		[SupportByVersion("Access", 9,10,11,12,14,15,16)]
		Int32 Verb { get; set; }

		/// <summary>
		/// SupportByVersion Access 9, 10, 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff836267.aspx </remarks>
		[SupportByVersion("Access", 9,10,11,12,14,15,16)]
		byte OLEType { get; set; }

		/// <summary>
		/// SupportByVersion Access 9, 10, 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff834345.aspx </remarks>
		[SupportByVersion("Access", 9,10,11,12,14,15,16)]
		byte OLETypeAllowed { get; set; }

		/// <summary>
		/// SupportByVersion Access 9, 10, 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff197939.aspx </remarks>
		[SupportByVersion("Access", 9,10,11,12,14,15,16)]
		string SourceObject { get; set; }

		/// <summary>
		/// SupportByVersion Access 9, 10, 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff192519.aspx </remarks>
		[SupportByVersion("Access", 9,10,11,12,14,15,16)]
		string Class { get; set; }

		/// <summary>
		/// SupportByVersion Access 9, 10, 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff191792.aspx </remarks>
		[SupportByVersion("Access", 9,10,11,12,14,15,16)]
		string SourceDoc { get; set; }

		/// <summary>
		/// SupportByVersion Access 9, 10, 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff197008.aspx </remarks>
		[SupportByVersion("Access", 9,10,11,12,14,15,16)]
		string SourceItem { get; set; }

		/// <summary>
		/// SupportByVersion Access 9, 10, 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff822758.aspx </remarks>
		[SupportByVersion("Access", 9,10,11,12,14,15,16)]
		Int16 ColumnCount { get; set; }

		/// <summary>
		/// SupportByVersion Access 9, 10, 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff836651.aspx </remarks>
		[SupportByVersion("Access", 9,10,11,12,14,15,16)]
		bool ColumnHeads { get; set; }

		/// <summary>
		/// SupportByVersion Access 9, 10, 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff191818.aspx </remarks>
		[SupportByVersion("Access", 9,10,11,12,14,15,16)]
		bool Visible { get; set; }

		/// <summary>
		/// SupportByVersion Access 9, 10, 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff192500.aspx </remarks>
		[SupportByVersion("Access", 9,10,11,12,14,15,16)]
		byte DisplayWhen { get; set; }

		/// <summary>
		/// SupportByVersion Access 9, 10, 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		[SupportByVersion("Access", 9,10,11,12,14,15,16)]
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		object OleData { get; set; }

		/// <summary>
		/// SupportByVersion Access 9, 10, 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff192679.aspx </remarks>
		[SupportByVersion("Access", 9,10,11,12,14,15,16)]
		Int16 UpdateMethod { get; set; }

		/// <summary>
		/// SupportByVersion Access 9, 10, 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff845706.aspx </remarks>
		[SupportByVersion("Access", 9,10,11,12,14,15,16)]
		bool Enabled { get; set; }

		/// <summary>
		/// SupportByVersion Access 9, 10, 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff844972.aspx </remarks>
		[SupportByVersion("Access", 9,10,11,12,14,15,16)]
		bool Locked { get; set; }

		/// <summary>
		/// SupportByVersion Access 9, 10, 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff193892.aspx </remarks>
		[SupportByVersion("Access", 9,10,11,12,14,15,16)]
		string StatusBarText { get; set; }

		/// <summary>
		/// SupportByVersion Access 9, 10, 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff192024.aspx </remarks>
		[SupportByVersion("Access", 9,10,11,12,14,15,16)]
		bool TabStop { get; set; }

		/// <summary>
		/// SupportByVersion Access 9, 10, 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff192839.aspx </remarks>
		[SupportByVersion("Access", 9,10,11,12,14,15,16)]
		Int16 TabIndex { get; set; }

		/// <summary>
		/// SupportByVersion Access 9, 10, 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff192272.aspx </remarks>
		[SupportByVersion("Access", 9,10,11,12,14,15,16)]
		Int16 Left { get; set; }

		/// <summary>
		/// SupportByVersion Access 9, 10, 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff822470.aspx </remarks>
		[SupportByVersion("Access", 9,10,11,12,14,15,16)]
		Int16 Top { get; set; }

		/// <summary>
		/// SupportByVersion Access 9, 10, 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff196176.aspx </remarks>
		[SupportByVersion("Access", 9,10,11,12,14,15,16)]
		Int16 Width { get; set; }

		/// <summary>
		/// SupportByVersion Access 9, 10, 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff194077.aspx </remarks>
		[SupportByVersion("Access", 9,10,11,12,14,15,16)]
		Int16 Height { get; set; }

		/// <summary>
		/// SupportByVersion Access 9, 10, 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff193583.aspx </remarks>
		[SupportByVersion("Access", 9,10,11,12,14,15,16)]
		byte BackStyle { get; set; }

		/// <summary>
		/// SupportByVersion Access 9, 10, 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff823149.aspx </remarks>
		[SupportByVersion("Access", 9,10,11,12,14,15,16)]
		Int32 BackColor { get; set; }

		/// <summary>
		/// SupportByVersion Access 9, 10, 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff192867.aspx </remarks>
		[SupportByVersion("Access", 9,10,11,12,14,15,16)]
		byte SpecialEffect { get; set; }

		/// <summary>
		/// SupportByVersion Access 9, 10, 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff191813.aspx </remarks>
		[SupportByVersion("Access", 9,10,11,12,14,15,16)]
		byte BorderStyle { get; set; }

		/// <summary>
		/// SupportByVersion Access 9, 10, 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff835749.aspx </remarks>
		[SupportByVersion("Access", 9,10,11,12,14,15,16)]
		byte OldBorderStyle { get; set; }

		/// <summary>
		/// SupportByVersion Access 9, 10, 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff196757.aspx </remarks>
		[SupportByVersion("Access", 9,10,11,12,14,15,16)]
		Int32 BorderColor { get; set; }

		/// <summary>
		/// SupportByVersion Access 9, 10, 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff192255.aspx </remarks>
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
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff844945.aspx </remarks>
		[SupportByVersion("Access", 9,10,11,12,14,15,16)]
		string ShortcutMenuBar { get; set; }

		/// <summary>
		/// SupportByVersion Access 9, 10, 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff195454.aspx </remarks>
		[SupportByVersion("Access", 9,10,11,12,14,15,16)]
		string ControlTipText { get; set; }

		/// <summary>
		/// SupportByVersion Access 9, 10, 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff196802.aspx </remarks>
		[SupportByVersion("Access", 9,10,11,12,14,15,16)]
		Int32 HelpContextId { get; set; }

		/// <summary>
		/// SupportByVersion Access 9, 10, 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		[SupportByVersion("Access", 9,10,11,12,14,15,16)]
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		Int16 Data { get; set; }

		/// <summary>
		/// SupportByVersion Access 9, 10, 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff822742.aspx </remarks>
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
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff845142.aspx </remarks>
		[SupportByVersion("Access", 9,10,11,12,14,15,16)]
		string Tag { get; set; }

		/// <summary>
		/// SupportByVersion Access 9, 10, 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff845409.aspx </remarks>
		[SupportByVersion("Access", 9,10,11,12,14,15,16)]
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		object ObjectPalette { get; set; }

		/// <summary>
		/// SupportByVersion Access 9, 10, 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		[SupportByVersion("Access", 9,10,11,12,14,15,16)]
		Int32 LpOleObject { get; set; }

		/// <summary>
		/// SupportByVersion Access 9, 10, 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff197338.aspx </remarks>
		[SupportByVersion("Access", 9,10,11,12,14,15,16)]
		Int32 ObjectVerbsCount { get; set; }

		/// <summary>
		/// SupportByVersion Access 9, 10, 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff844827.aspx </remarks>
		[SupportByVersion("Access", 9,10,11,12,14,15,16)]
		Int16 Action { get; set; }

		/// <summary>
		/// SupportByVersion Access 9, 10, 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff836289.aspx </remarks>
		[SupportByVersion("Access", 9,10,11,12,14,15,16)]
		byte Scaling { get; set; }

		/// <summary>
		/// SupportByVersion Access 9, 10, 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff845300.aspx </remarks>
		[SupportByVersion("Access", 9,10,11,12,14,15,16)]
		bool IsVisible { get; set; }

		/// <summary>
		/// SupportByVersion Access 9, 10, 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff194505.aspx </remarks>
		[SupportByVersion("Access", 9,10,11,12,14,15,16)]
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		bool InSelection { get; set; }

		/// <summary>
		/// SupportByVersion Access 9, 10, 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff834765.aspx </remarks>
		[SupportByVersion("Access", 9,10,11,12,14,15,16)]
		string OnUpdated { get; set; }

		/// <summary>
		/// SupportByVersion Access 9, 10, 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff192512.aspx </remarks>
		[SupportByVersion("Access", 9,10,11,12,14,15,16)]
		string OnEnter { get; set; }

		/// <summary>
		/// SupportByVersion Access 9, 10, 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff193221.aspx </remarks>
		[SupportByVersion("Access", 9,10,11,12,14,15,16)]
		string OnExit { get; set; }

		/// <summary>
		/// SupportByVersion Access 9, 10, 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff192514.aspx </remarks>
		[SupportByVersion("Access", 9,10,11,12,14,15,16)]
		string OnGotFocus { get; set; }

		/// <summary>
		/// SupportByVersion Access 9, 10, 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff193234.aspx </remarks>
		[SupportByVersion("Access", 9,10,11,12,14,15,16)]
		string OnLostFocus { get; set; }

		/// <summary>
		/// SupportByVersion Access 9, 10, 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff193910.aspx </remarks>
		[SupportByVersion("Access", 9,10,11,12,14,15,16)]
		string OnClick { get; set; }

		/// <summary>
		/// SupportByVersion Access 9, 10, 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff192458.aspx </remarks>
		[SupportByVersion("Access", 9,10,11,12,14,15,16)]
		string OnDblClick { get; set; }

		/// <summary>
		/// SupportByVersion Access 9, 10, 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff835392.aspx </remarks>
		[SupportByVersion("Access", 9,10,11,12,14,15,16)]
		string OnMouseDown { get; set; }

		/// <summary>
		/// SupportByVersion Access 9, 10, 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff822085.aspx </remarks>
		[SupportByVersion("Access", 9,10,11,12,14,15,16)]
		string OnMouseMove { get; set; }

		/// <summary>
		/// SupportByVersion Access 9, 10, 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff845164.aspx </remarks>
		[SupportByVersion("Access", 9,10,11,12,14,15,16)]
		string OnMouseUp { get; set; }

		/// <summary>
		/// SupportByVersion Access 9, 10, 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff835093.aspx </remarks>
		[SupportByVersion("Access", 9,10,11,12,14,15,16)]
		string Name { get; set; }

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
		/// Get
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff822733.aspx </remarks>
		[SupportByVersion("Access", 12,14,15,16)]
		NetOffice.AccessApi.Enums.AcLayoutType Layout { get; }

		/// <summary>
		/// SupportByVersion Access 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff822002.aspx </remarks>
		[SupportByVersion("Access", 12,14,15,16)]
		Int16 LeftPadding { get; set; }

		/// <summary>
		/// SupportByVersion Access 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff834395.aspx </remarks>
		[SupportByVersion("Access", 12,14,15,16)]
		Int16 TopPadding { get; set; }

		/// <summary>
		/// SupportByVersion Access 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff197942.aspx </remarks>
		[SupportByVersion("Access", 12,14,15,16)]
		Int16 RightPadding { get; set; }

		/// <summary>
		/// SupportByVersion Access 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff835660.aspx </remarks>
		[SupportByVersion("Access", 12,14,15,16)]
		Int16 BottomPadding { get; set; }

		/// <summary>
		/// SupportByVersion Access 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff845065.aspx </remarks>
		[SupportByVersion("Access", 12,14,15,16)]
		byte GridlineStyleLeft { get; set; }

		/// <summary>
		/// SupportByVersion Access 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff834679.aspx </remarks>
		[SupportByVersion("Access", 12,14,15,16)]
		byte GridlineStyleTop { get; set; }

		/// <summary>
		/// SupportByVersion Access 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff197955.aspx </remarks>
		[SupportByVersion("Access", 12,14,15,16)]
		byte GridlineStyleRight { get; set; }

		/// <summary>
		/// SupportByVersion Access 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff195270.aspx </remarks>
		[SupportByVersion("Access", 12,14,15,16)]
		byte GridlineStyleBottom { get; set; }

		/// <summary>
		/// SupportByVersion Access 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff197015.aspx </remarks>
		[SupportByVersion("Access", 12,14,15,16)]
		byte GridlineWidthLeft { get; set; }

		/// <summary>
		/// SupportByVersion Access 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff195842.aspx </remarks>
		[SupportByVersion("Access", 12,14,15,16)]
		byte GridlineWidthTop { get; set; }

		/// <summary>
		/// SupportByVersion Access 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff193249.aspx </remarks>
		[SupportByVersion("Access", 12,14,15,16)]
		byte GridlineWidthRight { get; set; }

		/// <summary>
		/// SupportByVersion Access 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff836743.aspx </remarks>
		[SupportByVersion("Access", 12,14,15,16)]
		byte GridlineWidthBottom { get; set; }

		/// <summary>
		/// SupportByVersion Access 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff195875.aspx </remarks>
		[SupportByVersion("Access", 12,14,15,16)]
		Int32 GridlineColor { get; set; }

		/// <summary>
		/// SupportByVersion Access 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff196767.aspx </remarks>
		[SupportByVersion("Access", 12,14,15,16)]
		NetOffice.AccessApi.Enums.AcHorizontalAnchor HorizontalAnchor { get; set; }

		/// <summary>
		/// SupportByVersion Access 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff197784.aspx </remarks>
		[SupportByVersion("Access", 12,14,15,16)]
		NetOffice.AccessApi.Enums.AcVerticalAnchor VerticalAnchor { get; set; }

		/// <summary>
		/// SupportByVersion Access 12, 14, 15, 16
		/// Get
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff194370.aspx </remarks>
		[SupportByVersion("Access", 12,14,15,16)]
		Int32 LayoutID { get; }

		/// <summary>
		/// SupportByVersion Access 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff198105.aspx </remarks>
		[SupportByVersion("Access", 14,15,16)]
		Int32 BackThemeColorIndex { get; set; }

		/// <summary>
		/// SupportByVersion Access 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff196481.aspx </remarks>
		[SupportByVersion("Access", 14,15,16)]
		Single BackTint { get; set; }

		/// <summary>
		/// SupportByVersion Access 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff195254.aspx </remarks>
		[SupportByVersion("Access", 14,15,16)]
		Single BackShade { get; set; }

		/// <summary>
		/// SupportByVersion Access 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff196476.aspx </remarks>
		[SupportByVersion("Access", 14,15,16)]
		Int32 BorderThemeColorIndex { get; set; }

		/// <summary>
		/// SupportByVersion Access 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff836588.aspx </remarks>
		[SupportByVersion("Access", 14,15,16)]
		Single BorderTint { get; set; }

		/// <summary>
		/// SupportByVersion Access 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff835423.aspx </remarks>
		[SupportByVersion("Access", 14,15,16)]
		Single BorderShade { get; set; }

		/// <summary>
		/// SupportByVersion Access 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff193207.aspx </remarks>
		[SupportByVersion("Access", 14,15,16)]
		Int32 GridlineThemeColorIndex { get; set; }

		/// <summary>
		/// SupportByVersion Access 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff836652.aspx </remarks>
		[SupportByVersion("Access", 14,15,16)]
		Single GridlineTint { get; set; }

		/// <summary>
		/// SupportByVersion Access 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff196748.aspx </remarks>
		[SupportByVersion("Access", 14,15,16)]
		Single GridlineShade { get; set; }

		/// <summary>
		/// SupportByVersion Access 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff835651.aspx </remarks>
		[SupportByVersion("Access", 14,15,16)]
		object VarOleObject { get; set; }

		#endregion

		#region Methods

		/// <summary>
		/// SupportByVersion Access 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff836575.aspx </remarks>
		[SupportByVersion("Access", 9,10,11,12,14,15,16)]
		void SizeToFit();

		/// <summary>
		/// SupportByVersion Access 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff193199.aspx </remarks>
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
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff196788.aspx </remarks>
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
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff194963.aspx </remarks>
		/// <param name="left">object left</param>
		/// <param name="top">optional object top</param>
		/// <param name="width">optional object width</param>
		/// <param name="height">optional object height</param>
		[SupportByVersion("Access", 10,11,12,14,15,16)]
		void Move(object left, object top, object width, object height);

		/// <summary>
		/// SupportByVersion Access 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff194963.aspx </remarks>
		/// <param name="left">object left</param>
		[CustomMethod]
		[SupportByVersion("Access", 10,11,12,14,15,16)]
		void Move(object left);

		/// <summary>
		/// SupportByVersion Access 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff194963.aspx </remarks>
		/// <param name="left">object left</param>
		/// <param name="top">optional object top</param>
		[CustomMethod]
		[SupportByVersion("Access", 10,11,12,14,15,16)]
		void Move(object left, object top);

		/// <summary>
		/// SupportByVersion Access 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff194963.aspx </remarks>
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
