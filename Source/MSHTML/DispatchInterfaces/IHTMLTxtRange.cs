using System;
using NetRuntimeSystem = System;
using System.ComponentModel;
using NetOffice.Attributes;

namespace NetOffice.MSHTMLApi
{
	/// <summary>
	/// DispatchInterface IHTMLTxtRange 
	/// SupportByVersion MSHTML, 4
	/// </summary>
	[SupportByVersion("MSHTML", 4)]
	[EntityType(EntityType.IsDispatchInterface)]
	[TypeId("3050F220-98B5-11CF-BB82-00AA00BDCE0B")]
	public interface IHTMLTxtRange : ICOMObject
	{
		#region Properties

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// Get
		/// </summary>
		[SupportByVersion("MSHTML", 4)]
		string htmlText { get; }

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// Get/Set
		/// </summary>
		[SupportByVersion("MSHTML", 4)]
		string text { get; set; }

		#endregion

		#region Methods

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// </summary>
		[SupportByVersion("MSHTML", 4)]
		[BaseResult]
		NetOffice.MSHTMLApi.IHTMLElement parentElement();

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// </summary>
		[SupportByVersion("MSHTML", 4)]
		NetOffice.MSHTMLApi.IHTMLTxtRange duplicate();

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// </summary>
		/// <param name="range">NetOffice.MSHTMLApi.IHTMLTxtRange range</param>
		[SupportByVersion("MSHTML", 4)]
		bool inRange(NetOffice.MSHTMLApi.IHTMLTxtRange range);

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// </summary>
		/// <param name="range">NetOffice.MSHTMLApi.IHTMLTxtRange range</param>
		[SupportByVersion("MSHTML", 4)]
		bool isEqual(NetOffice.MSHTMLApi.IHTMLTxtRange range);

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// </summary>
		/// <param name="fStart">optional bool fStart = true</param>
		[SupportByVersion("MSHTML", 4)]
		void scrollIntoView(object fStart);

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// </summary>
		[CustomMethod]
		[SupportByVersion("MSHTML", 4)]
		void scrollIntoView();

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// </summary>
		/// <param name="start">optional bool Start = true</param>
		[SupportByVersion("MSHTML", 4)]
		void collapse(object start);

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// </summary>
		[CustomMethod]
		[SupportByVersion("MSHTML", 4)]
		void collapse();

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// </summary>
		/// <param name="unit">string unit</param>
		[SupportByVersion("MSHTML", 4)]
		bool expand(string unit);

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// </summary>
		/// <param name="unit">string unit</param>
		/// <param name="count">optional Int32 Count = 1</param>
		[SupportByVersion("MSHTML", 4)]
		Int32 move(string unit, object count);

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// </summary>
		/// <param name="unit">string unit</param>
		[CustomMethod]
		[SupportByVersion("MSHTML", 4)]
		Int32 move(string unit);

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// </summary>
		/// <param name="unit">string unit</param>
		/// <param name="count">optional Int32 Count = 1</param>
		[SupportByVersion("MSHTML", 4)]
		Int32 moveStart(string unit, object count);

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// </summary>
		/// <param name="unit">string unit</param>
		[CustomMethod]
		[SupportByVersion("MSHTML", 4)]
		Int32 moveStart(string unit);

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// </summary>
		/// <param name="unit">string unit</param>
		/// <param name="count">optional Int32 Count = 1</param>
		[SupportByVersion("MSHTML", 4)]
		Int32 moveEnd(string unit, object count);

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// </summary>
		/// <param name="unit">string unit</param>
		[CustomMethod]
		[SupportByVersion("MSHTML", 4)]
		Int32 moveEnd(string unit);

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// </summary>
		[SupportByVersion("MSHTML", 4)]
		void select();

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// </summary>
		/// <param name="html">string html</param>
		[SupportByVersion("MSHTML", 4)]
		void pasteHTML(string html);

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// </summary>
		/// <param name="element">NetOffice.MSHTMLApi.IHTMLElement element</param>
		[SupportByVersion("MSHTML", 4)]
		void moveToElementText(NetOffice.MSHTMLApi.IHTMLElement element);

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// </summary>
		/// <param name="how">string how</param>
		/// <param name="sourceRange">NetOffice.MSHTMLApi.IHTMLTxtRange sourceRange</param>
		[SupportByVersion("MSHTML", 4)]
		void setEndPoint(string how, NetOffice.MSHTMLApi.IHTMLTxtRange sourceRange);

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// </summary>
		/// <param name="how">string how</param>
		/// <param name="sourceRange">NetOffice.MSHTMLApi.IHTMLTxtRange sourceRange</param>
		[SupportByVersion("MSHTML", 4)]
		Int32 compareEndPoints(string how, NetOffice.MSHTMLApi.IHTMLTxtRange sourceRange);

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// </summary>
		/// <param name="_string">string string</param>
		/// <param name="count">optional Int32 Count = 1073741823</param>
		/// <param name="flags">optional Int32 Flags = 0</param>
		[SupportByVersion("MSHTML", 4)]
		bool findText(string _string, object count, object flags);

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// </summary>
		/// <param name="_string">string string</param>
		[CustomMethod]
		[SupportByVersion("MSHTML", 4)]
		bool findText(string _string);

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// </summary>
		/// <param name="_string">string string</param>
		/// <param name="count">optional Int32 Count = 1073741823</param>
		[CustomMethod]
		[SupportByVersion("MSHTML", 4)]
		bool findText(string _string, object count);

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// </summary>
		/// <param name="x">Int32 x</param>
		/// <param name="y">Int32 y</param>
		[SupportByVersion("MSHTML", 4)]
		void moveToPoint(Int32 x, Int32 y);

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// </summary>
		[SupportByVersion("MSHTML", 4)]
		string getBookmark();

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// </summary>
		/// <param name="bookmark">string bookmark</param>
		[SupportByVersion("MSHTML", 4)]
		bool moveToBookmark(string bookmark);

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// </summary>
		/// <param name="cmdID">string cmdID</param>
		[SupportByVersion("MSHTML", 4)]
		bool queryCommandSupported(string cmdID);

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// </summary>
		/// <param name="cmdID">string cmdID</param>
		[SupportByVersion("MSHTML", 4)]
		bool queryCommandEnabled(string cmdID);

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// </summary>
		/// <param name="cmdID">string cmdID</param>
		[SupportByVersion("MSHTML", 4)]
		bool queryCommandState(string cmdID);

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// </summary>
		/// <param name="cmdID">string cmdID</param>
		[SupportByVersion("MSHTML", 4)]
		bool queryCommandIndeterm(string cmdID);

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// </summary>
		/// <param name="cmdID">string cmdID</param>
		[SupportByVersion("MSHTML", 4)]
		string queryCommandText(string cmdID);

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// </summary>
		/// <param name="cmdID">string cmdID</param>
		[SupportByVersion("MSHTML", 4)]
		object queryCommandValue(string cmdID);

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// </summary>
		/// <param name="cmdID">string cmdID</param>
		/// <param name="showUI">optional bool showUI = false</param>
		/// <param name="value">optional object value</param>
		[SupportByVersion("MSHTML", 4)]
		bool execCommand(string cmdID, object showUI, object value);

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// </summary>
		/// <param name="cmdID">string cmdID</param>
		[CustomMethod]
		[SupportByVersion("MSHTML", 4)]
		bool execCommand(string cmdID);

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// </summary>
		/// <param name="cmdID">string cmdID</param>
		/// <param name="showUI">optional bool showUI = false</param>
		[CustomMethod]
		[SupportByVersion("MSHTML", 4)]
		bool execCommand(string cmdID, object showUI);

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// </summary>
		/// <param name="cmdID">string cmdID</param>
		[SupportByVersion("MSHTML", 4)]
		bool execCommandShowHelp(string cmdID);

		#endregion
	}
}
