using System.Reflection;
using System;
using NetRuntimeSystem = System;
using System.ComponentModel;
using NetOffice.Attributes;

namespace NetOffice.MSHTMLApi
{
	/// <summary>
	/// Interface IMarkupTextFrags 
	/// SupportByVersion MSHTML, 4
	/// </summary>
	[SupportByVersion("MSHTML", 4)]
	[EntityType(EntityType.IsInterface)]
	[TypeId("3050F5FA-98B5-11CF-BB82-00AA00BDCE0B")]
	public interface IMarkupTextFrags : ICOMObject
	{
		#region Methods

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// </summary>
		/// <param name="pcFrags">Int32 pcFrags</param>
		[SupportByVersion("MSHTML", 4)]
		Int32 GetTextFragCount(out Int32 pcFrags);

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// </summary>
		/// <param name="iFrag">Int32 iFrag</param>
		/// <param name="pbstrFrag">string pbstrFrag</param>
		/// <param name="pPointerFrag">NetOffice.MSHTMLApi.IMarkupPointer pPointerFrag</param>
		[SupportByVersion("MSHTML", 4)]
		Int32 GetTextFrag(Int32 iFrag, out string pbstrFrag, NetOffice.MSHTMLApi.IMarkupPointer pPointerFrag);

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// </summary>
		/// <param name="iFrag">Int32 iFrag</param>
		[SupportByVersion("MSHTML", 4)]
		Int32 RemoveTextFrag(Int32 iFrag);

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// </summary>
		/// <param name="iFrag">Int32 iFrag</param>
		/// <param name="bstrInsert">string bstrInsert</param>
		/// <param name="pPointerInsert">NetOffice.MSHTMLApi.IMarkupPointer pPointerInsert</param>
		[SupportByVersion("MSHTML", 4)]
		Int32 InsertTextFrag(Int32 iFrag, string bstrInsert, NetOffice.MSHTMLApi.IMarkupPointer pPointerInsert);

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// </summary>
		/// <param name="pPointerFind">NetOffice.MSHTMLApi.IMarkupPointer pPointerFind</param>
		/// <param name="piFrag">Int32 piFrag</param>
		/// <param name="pfFragFound">Int32 pfFragFound</param>
		[SupportByVersion("MSHTML", 4)]
		Int32 FindTextFragFromMarkupPointer(NetOffice.MSHTMLApi.IMarkupPointer pPointerFind, out Int32 piFrag, out Int32 pfFragFound);

		#endregion
	}
}
