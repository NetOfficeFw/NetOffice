using System.Reflection;
using System;
using NetRuntimeSystem = System;
using System.ComponentModel;
using NetOffice.Attributes;

namespace NetOffice.MSHTMLApi
{
	/// <summary>
	/// Interface IElementSegment 
	/// SupportByVersion MSHTML, 4
	/// </summary>
	[SupportByVersion("MSHTML", 4)]
	[EntityType(EntityType.IsInterface)]
	[TypeId("3050F68F-98B5-11CF-BB82-00AA00BDCE0B")]
	public interface IElementSegment : ISegment
	{
		#region Methods

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// </summary>
		/// <param name="ppIElement">NetOffice.MSHTMLApi.IHTMLElement ppIElement</param>
		[SupportByVersion("MSHTML", 4)]
		Int32 GetElement(out NetOffice.MSHTMLApi.IHTMLElement ppIElement);

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// </summary>
		/// <param name="fPrimary">Int32 fPrimary</param>
		[SupportByVersion("MSHTML", 4)]
		Int32 SetPrimary(Int32 fPrimary);

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// </summary>
		/// <param name="pfPrimary">Int32 pfPrimary</param>
		[SupportByVersion("MSHTML", 4)]
		Int32 IsPrimary(out Int32 pfPrimary);

		#endregion
	}
}
