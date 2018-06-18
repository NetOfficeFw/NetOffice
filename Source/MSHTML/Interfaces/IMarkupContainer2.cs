using System.Reflection;
using System;
using NetRuntimeSystem = System;
using System.ComponentModel;
using NetOffice.Attributes;

namespace NetOffice.MSHTMLApi
{
	/// <summary>
	/// Interface IMarkupContainer2 
	/// SupportByVersion MSHTML, 4
	/// </summary>
	[SupportByVersion("MSHTML", 4)]
	[EntityType(EntityType.IsInterface)]
	[TypeId("3050F648-98B5-11CF-BB82-00AA00BDCE0B")]
	public interface IMarkupContainer2 : IMarkupContainer
	{
		#region Methods

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// </summary>
		/// <param name="pChangeSink">NetOffice.MSHTMLApi.IHTMLChangeSink pChangeSink</param>
		/// <param name="ppChangeLog">NetOffice.MSHTMLApi.IHTMLChangeLog ppChangeLog</param>
		/// <param name="fForward">Int32 fForward</param>
		/// <param name="fBackward">Int32 fBackward</param>
		[SupportByVersion("MSHTML", 4)]
		Int32 CreateChangeLog(NetOffice.MSHTMLApi.IHTMLChangeSink pChangeSink, out NetOffice.MSHTMLApi.IHTMLChangeLog ppChangeLog, Int32 fForward, Int32 fBackward);

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// </summary>
		/// <param name="pChangeSink">NetOffice.MSHTMLApi.IHTMLChangeSink pChangeSink</param>
		/// <param name="pdwCookie">Int32 pdwCookie</param>
		[SupportByVersion("MSHTML", 4)]
		Int32 RegisterForDirtyRange(NetOffice.MSHTMLApi.IHTMLChangeSink pChangeSink, out Int32 pdwCookie);

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// </summary>
		/// <param name="dwCookie">Int32 dwCookie</param>
		[SupportByVersion("MSHTML", 4)]
		Int32 UnRegisterForDirtyRange(Int32 dwCookie);

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// </summary>
		/// <param name="dwCookie">Int32 dwCookie</param>
		/// <param name="pIPointerBegin">NetOffice.MSHTMLApi.IMarkupPointer pIPointerBegin</param>
		/// <param name="pIPointerEnd">NetOffice.MSHTMLApi.IMarkupPointer pIPointerEnd</param>
		[SupportByVersion("MSHTML", 4)]
		Int32 GetAndClearDirtyRange(Int32 dwCookie, NetOffice.MSHTMLApi.IMarkupPointer pIPointerBegin, NetOffice.MSHTMLApi.IMarkupPointer pIPointerEnd);

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// </summary>
		[SupportByVersion("MSHTML", 4)]
		Int32 GetVersionNumber();

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// </summary>
		/// <param name="ppElementMaster">NetOffice.MSHTMLApi.IHTMLElement ppElementMaster</param>
		[SupportByVersion("MSHTML", 4)]
		Int32 GetMasterElement(out NetOffice.MSHTMLApi.IHTMLElement ppElementMaster);

		#endregion
	}
}
