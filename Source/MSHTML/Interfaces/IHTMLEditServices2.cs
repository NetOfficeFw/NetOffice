using System;
using NetRuntimeSystem = System;
using System.ComponentModel;
using NetOffice.Attributes;

namespace NetOffice.MSHTMLApi
{
	/// <summary>
	/// Interface IHTMLEditServices2 
	/// SupportByVersion MSHTML, 4
	/// </summary>
	[SupportByVersion("MSHTML", 4)]
	[EntityType(EntityType.IsInterface)]
	[TypeId("3050F812-98B5-11CF-BB82-00AA00BDCE0B")]
	public interface IHTMLEditServices2 : IHTMLEditServices
	{
		#region Methods

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// </summary>
		/// <param name="pIStartAnchor">NetOffice.MSHTMLApi.IDisplayPointer pIStartAnchor</param>
		[SupportByVersion("MSHTML", 4)]
		Int32 MoveToSelectionAnchorEx(NetOffice.MSHTMLApi.IDisplayPointer pIStartAnchor);

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// </summary>
		/// <param name="pIEndAnchor">NetOffice.MSHTMLApi.IDisplayPointer pIEndAnchor</param>
		[SupportByVersion("MSHTML", 4)]
		Int32 MoveToSelectionEndEx(NetOffice.MSHTMLApi.IDisplayPointer pIEndAnchor);

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// </summary>
		/// <param name="fReCompute">Int32 fReCompute</param>
		[SupportByVersion("MSHTML", 4)]
		Int32 FreezeVirtualCaretPos(Int32 fReCompute);

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// </summary>
		/// <param name="fReset">Int32 fReset</param>
		[SupportByVersion("MSHTML", 4)]
		Int32 UnFreezeVirtualCaretPos(Int32 fReset);

		#endregion
	}
}
