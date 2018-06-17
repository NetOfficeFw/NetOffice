using System.Reflection;
using System;
using NetRuntimeSystem = System;
using System.ComponentModel;
using NetOffice.Attributes;

namespace NetOffice.AccessApi
{
	/// <summary>
	/// DispatchInterface Module 
	/// SupportByVersion Access, 9,10,11,12,14,15,16
	/// </summary>
	/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff835649.aspx </remarks>
	[SupportByVersion("Access", 9,10,11,12,14,15,16)]
	[EntityType(EntityType.IsDispatchInterface)]
	[TypeId("331FDCFE-CF31-11CD-8701-00AA003F0F07")]
	public interface Module : ICOMObject
	{
		#region Properties

		/// <summary>
		/// SupportByVersion Access 9, 10, 11, 12, 14, 15, 16
		/// Get
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff197648.aspx </remarks>
		[SupportByVersion("Access", 9,10,11,12,14,15,16)]
		NetOffice.AccessApi.Application Application { get; }

		/// <summary>
		/// SupportByVersion Access 9, 10, 11, 12, 14, 15, 16
		/// Get
		/// Unknown COM Proxy
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff845790.aspx </remarks>
		[SupportByVersion("Access", 9,10,11,12,14,15,16), ProxyResult]
		object Parent { get; }

		/// <summary>
		/// SupportByVersion Access 9, 10, 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff192086.aspx </remarks>
		[SupportByVersion("Access", 9,10,11,12,14,15,16)]
		string Name { get; set; }

		/// <summary>
		/// SupportByVersion Access 9, 10, 11, 12, 14, 15, 16
		/// Get
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff820960.aspx </remarks>
		/// <param name="line">Int32 line</param>
		/// <param name="numLines">Int32 numLines</param>
		[SupportByVersion("Access", 9,10,11,12,14,15,16)]
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		string get_Lines(Int32 line, Int32 numLines);

		/// <summary>
		/// SupportByVersion Access 9, 10, 11, 12, 14, 15, 16
		/// Alias for get_Lines
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff820960.aspx </remarks>
		/// <param name="line">Int32 line</param>
		/// <param name="numLines">Int32 numLines</param>
		[SupportByVersion("Access", 9,10,11,12,14,15,16), Redirect("get_Lines")]
		string Lines(Int32 line, Int32 numLines);

		/// <summary>
		/// SupportByVersion Access 9, 10, 11, 12, 14, 15, 16
		/// Get
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff195500.aspx </remarks>
		[SupportByVersion("Access", 9,10,11,12,14,15,16)]
		Int32 CountOfLines { get; }

		/// <summary>
		/// SupportByVersion Access 9, 10, 11, 12, 14, 15, 16
		/// Get
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff836419.aspx </remarks>
		/// <param name="procName">string procName</param>
		/// <param name="procKind">NetOffice.VBIDEApi.Enums.vbext_ProcKind procKind</param>
		[SupportByVersion("Access", 9,10,11,12,14,15,16)]
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		Int32 get_ProcStartLine(string procName, NetOffice.VBIDEApi.Enums.vbext_ProcKind procKind);

		/// <summary>
		/// SupportByVersion Access 9, 10, 11, 12, 14, 15, 16
		/// Alias for get_ProcStartLine
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff836419.aspx </remarks>
		/// <param name="procName">string procName</param>
		/// <param name="procKind">NetOffice.VBIDEApi.Enums.vbext_ProcKind procKind</param>
		[SupportByVersion("Access", 9,10,11,12,14,15,16), Redirect("get_ProcStartLine")]
		Int32 ProcStartLine(string procName, NetOffice.VBIDEApi.Enums.vbext_ProcKind procKind);

		/// <summary>
		/// SupportByVersion Access 9, 10, 11, 12, 14, 15, 16
		/// Get
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff835086.aspx </remarks>
		/// <param name="procName">string procName</param>
		/// <param name="procKind">NetOffice.VBIDEApi.Enums.vbext_ProcKind procKind</param>
		[SupportByVersion("Access", 9,10,11,12,14,15,16)]
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		Int32 get_ProcCountLines(string procName, NetOffice.VBIDEApi.Enums.vbext_ProcKind procKind);

		/// <summary>
		/// SupportByVersion Access 9, 10, 11, 12, 14, 15, 16
		/// Alias for get_ProcCountLines
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff835086.aspx </remarks>
		/// <param name="procName">string procName</param>
		/// <param name="procKind">NetOffice.VBIDEApi.Enums.vbext_ProcKind procKind</param>
		[SupportByVersion("Access", 9,10,11,12,14,15,16), Redirect("get_ProcCountLines")]
		Int32 ProcCountLines(string procName, NetOffice.VBIDEApi.Enums.vbext_ProcKind procKind);

		/// <summary>
		/// SupportByVersion Access 9, 10, 11, 12, 14, 15, 16
		/// Get
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff822434.aspx </remarks>
		/// <param name="procName">string procName</param>
		/// <param name="procKind">NetOffice.VBIDEApi.Enums.vbext_ProcKind procKind</param>
		[SupportByVersion("Access", 9,10,11,12,14,15,16)]
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		Int32 get_ProcBodyLine(string procName, NetOffice.VBIDEApi.Enums.vbext_ProcKind procKind);

		/// <summary>
		/// SupportByVersion Access 9, 10, 11, 12, 14, 15, 16
		/// Alias for get_ProcBodyLine
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff822434.aspx </remarks>
		/// <param name="procName">string procName</param>
		/// <param name="procKind">NetOffice.VBIDEApi.Enums.vbext_ProcKind procKind</param>
		[SupportByVersion("Access", 9,10,11,12,14,15,16), Redirect("get_ProcBodyLine")]
		Int32 ProcBodyLine(string procName, NetOffice.VBIDEApi.Enums.vbext_ProcKind procKind);

		/// <summary>
		/// SupportByVersion Access 9, 10, 11, 12, 14, 15, 16
		/// Get
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff195085.aspx </remarks>
		/// <param name="line">Int32 line</param>
		/// <param name="pprockind">NetOffice.VBIDEApi.Enums.vbext_ProcKind pprockind</param>
		[SupportByVersion("Access", 9,10,11,12,14,15,16)]
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		string get_ProcOfLine(Int32 line, out NetOffice.VBIDEApi.Enums.vbext_ProcKind pprockind);

		/// <summary>
		/// SupportByVersion Access 9, 10, 11, 12, 14, 15, 16
		/// Alias for get_ProcOfLine
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff195085.aspx </remarks>
		/// <param name="line">Int32 line</param>
		/// <param name="pprockind">NetOffice.VBIDEApi.Enums.vbext_ProcKind pprockind</param>
		[SupportByVersion("Access", 9,10,11,12,14,15,16), Redirect("get_ProcOfLine")]
		string ProcOfLine(Int32 line, out NetOffice.VBIDEApi.Enums.vbext_ProcKind pprockind);

		/// <summary>
		/// SupportByVersion Access 9, 10, 11, 12, 14, 15, 16
		/// Get
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff837190.aspx </remarks>
		[SupportByVersion("Access", 9,10,11,12,14,15,16)]
		Int32 CountOfDeclarationLines { get; }

		/// <summary>
		/// SupportByVersion Access 9, 10, 11, 12, 14, 15, 16
		/// Get
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff835633.aspx </remarks>
		[SupportByVersion("Access", 9,10,11,12,14,15,16)]
		NetOffice.AccessApi.Enums.AcModuleType Type { get; }

		#endregion

		#region Methods

		/// <summary>
		/// SupportByVersion Access 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff845332.aspx </remarks>
		/// <param name="text">string text</param>
		[SupportByVersion("Access", 9,10,11,12,14,15,16)]
		void InsertText(string text);

		/// <summary>
		/// SupportByVersion Access 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff845379.aspx </remarks>
		/// <param name="_string">string string</param>
		[SupportByVersion("Access", 9,10,11,12,14,15,16)]
		void AddFromString(string _string);

		/// <summary>
		/// SupportByVersion Access 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff821352.aspx </remarks>
		/// <param name="fileName">string fileName</param>
		[SupportByVersion("Access", 9,10,11,12,14,15,16)]
		void AddFromFile(string fileName);

		/// <summary>
		/// SupportByVersion Access 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff194137.aspx </remarks>
		/// <param name="line">Int32 line</param>
		/// <param name="_string">string string</param>
		[SupportByVersion("Access", 9,10,11,12,14,15,16)]
		void InsertLines(Int32 line, string _string);

		/// <summary>
		/// SupportByVersion Access 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff194301.aspx </remarks>
		/// <param name="startLine">Int32 startLine</param>
		/// <param name="count">Int32 count</param>
		[SupportByVersion("Access", 9,10,11,12,14,15,16)]
		void DeleteLines(Int32 startLine, Int32 count);

		/// <summary>
		/// SupportByVersion Access 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff198276.aspx </remarks>
		/// <param name="line">Int32 line</param>
		/// <param name="_string">string string</param>
		[SupportByVersion("Access", 9,10,11,12,14,15,16)]
		void ReplaceLine(Int32 line, string _string);

		/// <summary>
		/// SupportByVersion Access 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff845440.aspx </remarks>
		/// <param name="eventName">string eventName</param>
		/// <param name="objectName">string objectName</param>
		[SupportByVersion("Access", 9,10,11,12,14,15,16)]
		Int32 CreateEventProc(string eventName, string objectName);

		/// <summary>
		/// SupportByVersion Access 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff195471.aspx </remarks>
		/// <param name="target">string target</param>
		/// <param name="startLine">Int32 startLine</param>
		/// <param name="startColumn">Int32 startColumn</param>
		/// <param name="endLine">Int32 endLine</param>
		/// <param name="endColumn">Int32 endColumn</param>
		/// <param name="wholeWord">optional bool WholeWord = false</param>
		/// <param name="matchCase">optional bool MatchCase = false</param>
		/// <param name="patternSearch">optional bool PatternSearch = false</param>
		[SupportByVersion("Access", 9,10,11,12,14,15,16)]
		bool Find(string target, Int32 startLine, Int32 startColumn, Int32 endLine, Int32 endColumn, object wholeWord, object matchCase, object patternSearch);

		/// <summary>
		/// SupportByVersion Access 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff195471.aspx </remarks>
		/// <param name="target">string target</param>
		/// <param name="startLine">Int32 startLine</param>
		/// <param name="startColumn">Int32 startColumn</param>
		/// <param name="endLine">Int32 endLine</param>
		/// <param name="endColumn">Int32 endColumn</param>
		[CustomMethod]
		[SupportByVersion("Access", 9,10,11,12,14,15,16)]
		bool Find(string target, Int32 startLine, Int32 startColumn, Int32 endLine, Int32 endColumn);

		/// <summary>
		/// SupportByVersion Access 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff195471.aspx </remarks>
		/// <param name="target">string target</param>
		/// <param name="startLine">Int32 startLine</param>
		/// <param name="startColumn">Int32 startColumn</param>
		/// <param name="endLine">Int32 endLine</param>
		/// <param name="endColumn">Int32 endColumn</param>
		/// <param name="wholeWord">optional bool WholeWord = false</param>
		[CustomMethod]
		[SupportByVersion("Access", 9,10,11,12,14,15,16)]
		bool Find(string target, Int32 startLine, Int32 startColumn, Int32 endLine, Int32 endColumn, object wholeWord);

		/// <summary>
		/// SupportByVersion Access 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff195471.aspx </remarks>
		/// <param name="target">string target</param>
		/// <param name="startLine">Int32 startLine</param>
		/// <param name="startColumn">Int32 startColumn</param>
		/// <param name="endLine">Int32 endLine</param>
		/// <param name="endColumn">Int32 endColumn</param>
		/// <param name="wholeWord">optional bool WholeWord = false</param>
		/// <param name="matchCase">optional bool MatchCase = false</param>
		[CustomMethod]
		[SupportByVersion("Access", 9,10,11,12,14,15,16)]
		bool Find(string target, Int32 startLine, Int32 startColumn, Int32 endLine, Int32 endColumn, object wholeWord, object matchCase);

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
