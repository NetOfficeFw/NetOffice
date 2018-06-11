using System.Reflection;
using System;
using NetRuntimeSystem = System;
using System.ComponentModel;
using NetOffice.Attributes;

namespace NetOffice.VBIDEApi
{
    /// <summary>
    /// DispatchInterface _CodeModule
    /// SupportByVersion VBIDE, 12,14,5.3
    /// </summary>
    [SupportByVersion("VBIDE", 12, 14, 5.3)]
    [EntityType(EntityType.IsDispatchInterface), BaseType]
	[TypeId("0002E16E-0000-0000-C000-000000000046")]
    [CoClassSource(typeof(NetOffice.VBIDEApi.CodeModule))]
    public interface _CodeModule : ICOMObject
    {
        #region Properties

        /// <summary>
        /// SupportByVersion VBIDE 12, 14, 5.3
        /// Get
        /// </summary>
        [SupportByVersion("VBIDE", 12, 14, 5.3)]
        NetOffice.VBIDEApi.VBComponent Parent { get; }

        /// <summary>
        /// SupportByVersion VBIDE 12, 14, 5.3
        /// Get
        /// </summary>
        [SupportByVersion("VBIDE", 12, 14, 5.3)]
        NetOffice.VBIDEApi.VBE VBE { get; }

        /// <summary>
		/// SupportByVersion VBIDE 12, 14, 5.3
		/// Get/Set
		/// </summary>
		[SupportByVersion("VBIDE", 12, 14, 5.3)]
        string Name { get; set; }

        /// <summary>
		/// SupportByVersion VBIDE 12, 14, 5.3
		/// Get
		/// </summary>
		/// <param name="startLine">Int32 startLine</param>
		/// <param name="count">Int32 count</param>
		[SupportByVersion("VBIDE", 12, 14, 5.3)]
        [EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
        string get_Lines(Int32 startLine, Int32 count);

        /// <summary>
		/// SupportByVersion VBIDE 12, 14, 5.3
		/// Alias for get_Lines
		/// </summary>
		/// <param name="startLine">Int32 startLine</param>
		/// <param name="count">Int32 count</param>
		[SupportByVersion("VBIDE", 12, 14, 5.3), Redirect("get_Lines")]
        string Lines(Int32 startLine, Int32 count);

        /// <summary>
		/// SupportByVersion VBIDE 12, 14, 5.3
		/// Get
		/// </summary>
		[SupportByVersion("VBIDE", 12, 14, 5.3)]
        Int32 CountOfLines { get; }

        /// <summary>
		/// SupportByVersion VBIDE 12, 14, 5.3
		/// Get
		/// </summary>
		/// <param name="procName">string procName</param>
		/// <param name="procKind">NetOffice.VBIDEApi.Enums.vbext_ProcKind procKind</param>
		[SupportByVersion("VBIDE", 12, 14, 5.3)]
        [EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
        Int32 get_ProcStartLine(string procName, NetOffice.VBIDEApi.Enums.vbext_ProcKind procKind);

        /// <summary>
        /// SupportByVersion VBIDE 12, 14, 5.3
        /// Alias for get_ProcStartLine
        /// </summary>
        /// <param name="procName">string procName</param>
        /// <param name="procKind">NetOffice.VBIDEApi.Enums.vbext_ProcKind procKind</param>
        [SupportByVersion("VBIDE", 12, 14, 5.3), Redirect("get_ProcStartLine")]
        Int32 ProcStartLine(string procName, NetOffice.VBIDEApi.Enums.vbext_ProcKind procKind);

        /// <summary>
		/// SupportByVersion VBIDE 12, 14, 5.3
		/// Get
		/// </summary>
		/// <param name="procName">string procName</param>
		/// <param name="procKind">NetOffice.VBIDEApi.Enums.vbext_ProcKind procKind</param>
		[SupportByVersion("VBIDE", 12, 14, 5.3)]
        [EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
        Int32 get_ProcCountLines(string procName, NetOffice.VBIDEApi.Enums.vbext_ProcKind procKind);

        /// <summary>
        /// SupportByVersion VBIDE 12, 14, 5.3
        /// Alias for get_ProcCountLines
        /// </summary>
        /// <param name="procName">string procName</param>
        /// <param name="procKind">NetOffice.VBIDEApi.Enums.vbext_ProcKind procKind</param>
        [SupportByVersion("VBIDE", 12, 14, 5.3), Redirect("get_ProcCountLines")]
        Int32 ProcCountLines(string procName, NetOffice.VBIDEApi.Enums.vbext_ProcKind procKind);

        /// <summary>
        /// SupportByVersion VBIDE 12, 14, 5.3
        /// Get
        /// </summary>
        /// <param name="procName">string procName</param>
        /// <param name="procKind">NetOffice.VBIDEApi.Enums.vbext_ProcKind procKind</param>
        [SupportByVersion("VBIDE", 12, 14, 5.3)]
        [EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
        Int32 get_ProcBodyLine(string procName, NetOffice.VBIDEApi.Enums.vbext_ProcKind procKind);

        /// <summary>
        /// SupportByVersion VBIDE 12, 14, 5.3
        /// Alias for get_ProcBodyLine
        /// </summary>
        /// <param name="procName">string procName</param>
        /// <param name="procKind">NetOffice.VBIDEApi.Enums.vbext_ProcKind procKind</param>
        [SupportByVersion("VBIDE", 12, 14, 5.3), Redirect("get_ProcBodyLine")]
        Int32 ProcBodyLine(string procName, NetOffice.VBIDEApi.Enums.vbext_ProcKind procKind);

        /// <summary>
        /// SupportByVersion VBIDE 12, 14, 5.3
        /// Get
        /// </summary>
        /// <param name="line">Int32 line</param>
        /// <param name="procKind">NetOffice.VBIDEApi.Enums.vbext_ProcKind procKind</param>
        [SupportByVersion("VBIDE", 12, 14, 5.3)]
        [EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
        string get_ProcOfLine(Int32 line, out NetOffice.VBIDEApi.Enums.vbext_ProcKind procKind);

        /// <summary>
        /// SupportByVersion VBIDE 12, 14, 5.3
        /// Alias for get_ProcOfLine
        /// </summary>
        /// <param name="line">Int32 line</param>
        /// <param name="procKind">NetOffice.VBIDEApi.Enums.vbext_ProcKind procKind</param>
        [SupportByVersion("VBIDE", 12, 14, 5.3), Redirect("get_ProcOfLine")]
        string ProcOfLine(Int32 line, out NetOffice.VBIDEApi.Enums.vbext_ProcKind procKind);

        /// <summary>
        /// SupportByVersion VBIDE 12, 14, 5.3
        /// Get
        /// </summary>
        [SupportByVersion("VBIDE", 12, 14, 5.3)]
        Int32 CountOfDeclarationLines { get; }

        /// <summary>
        /// SupportByVersion VBIDE 12, 14, 5.3
        /// Get
        /// </summary>
        [SupportByVersion("VBIDE", 12, 14, 5.3)]
        NetOffice.VBIDEApi.CodePane CodePane { get; }

        #endregion

        #region Methods

        /// <summary>
		/// SupportByVersion VBIDE 12, 14, 5.3
		/// </summary>
		/// <param name="_string">string string</param>
		[SupportByVersion("VBIDE", 12, 14, 5.3)]
        void AddFromString(string _string);

        /// <summary>
        /// SupportByVersion VBIDE 12, 14, 5.3
        /// </summary>
        /// <param name="fileName">string fileName</param>
        [SupportByVersion("VBIDE", 12, 14, 5.3)]
        void AddFromFile(string fileName);

        /// <summary>
        /// SupportByVersion VBIDE 12, 14, 5.3
        /// </summary>
        /// <param name="line">Int32 line</param>
        /// <param name="_string">string string</param>
        [SupportByVersion("VBIDE", 12, 14, 5.3)]
        void InsertLines(Int32 line, string _string);

        /// <summary>
        /// SupportByVersion VBIDE 12, 14, 5.3
        /// </summary>
        /// <param name="startLine">Int32 startLine</param>
        /// <param name="count">optional Int32 Count = 1</param>
        [SupportByVersion("VBIDE", 12, 14, 5.3)]
        void DeleteLines(Int32 startLine, object count);

        /// <summary>
        /// SupportByVersion VBIDE 12, 14, 5.3
        /// </summary>
        /// <param name="startLine">Int32 startLine</param>
        [CustomMethod]
        [SupportByVersion("VBIDE", 12, 14, 5.3)]
        void DeleteLines(Int32 startLine);

        /// <summary>
        /// SupportByVersion VBIDE 12, 14, 5.3
        /// </summary>
        /// <param name="line">Int32 line</param>
        /// <param name="_string">string string</param>
        [SupportByVersion("VBIDE", 12, 14, 5.3)]
        void ReplaceLine(Int32 line, string _string);

        /// <summary>
        /// SupportByVersion VBIDE 12, 14, 5.3
        /// </summary>
        /// <param name="eventName">string eventName</param>
        /// <param name="objectName">string objectName</param>
        [SupportByVersion("VBIDE", 12, 14, 5.3)]
        Int32 CreateEventProc(string eventName, string objectName);

        /// <summary>
        /// SupportByVersion VBIDE 12, 14, 5.3
        /// </summary>
        /// <param name="target">string target</param>
        /// <param name="startLine">Int32 startLine</param>
        /// <param name="startColumn">Int32 startColumn</param>
        /// <param name="endLine">Int32 endLine</param>
        /// <param name="endColumn">Int32 endColumn</param>
        /// <param name="wholeWord">optional bool WholeWord = false</param>
        /// <param name="matchCase">optional bool MatchCase = false</param>
        /// <param name="patternSearch">optional bool PatternSearch = false</param>
        [SupportByVersion("VBIDE", 12, 14, 5.3)]
        bool Find(string target, Int32 startLine, Int32 startColumn, Int32 endLine, Int32 endColumn, object wholeWord, object matchCase, object patternSearch);

        /// <summary>
        /// SupportByVersion VBIDE 12, 14, 5.3
        /// </summary>
        /// <param name="target">string target</param>
        /// <param name="startLine">Int32 startLine</param>
        /// <param name="startColumn">Int32 startColumn</param>
        /// <param name="endLine">Int32 endLine</param>
        /// <param name="endColumn">Int32 endColumn</param>
        [CustomMethod]
        [SupportByVersion("VBIDE", 12, 14, 5.3)]
        bool Find(string target, Int32 startLine, Int32 startColumn, Int32 endLine, Int32 endColumn);

        /// <summary>
        /// SupportByVersion VBIDE 12, 14, 5.3
        /// </summary>
        /// <param name="target">string target</param>
        /// <param name="startLine">Int32 startLine</param>
        /// <param name="startColumn">Int32 startColumn</param>
        /// <param name="endLine">Int32 endLine</param>
        /// <param name="endColumn">Int32 endColumn</param>
        /// <param name="wholeWord">optional bool WholeWord = false</param>
        [CustomMethod]
        [SupportByVersion("VBIDE", 12, 14, 5.3)]
        bool Find(string target, Int32 startLine, Int32 startColumn, Int32 endLine, Int32 endColumn, object wholeWord);

        /// <summary>
        /// SupportByVersion VBIDE 12, 14, 5.3
        /// </summary>
        /// <param name="target">string target</param>
        /// <param name="startLine">Int32 startLine</param>
        /// <param name="startColumn">Int32 startColumn</param>
        /// <param name="endLine">Int32 endLine</param>
        /// <param name="endColumn">Int32 endColumn</param>
        /// <param name="wholeWord">optional bool WholeWord = false</param>
        /// <param name="matchCase">optional bool MatchCase = false</param>
        [CustomMethod]
        [SupportByVersion("VBIDE", 12, 14, 5.3)]
        bool Find(string target, Int32 startLine, Int32 startColumn, Int32 endLine, Int32 endColumn, object wholeWord, object matchCase);

        #endregion
    }
}
