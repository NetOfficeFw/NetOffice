using System;
using NetRuntimeSystem = System;
using System.ComponentModel;
using NetOffice.Attributes;

namespace NetOffice.MSHTMLApi
{
	/// <summary>
	/// Interface ISegment 
	/// SupportByVersion MSHTML, 4
	/// </summary>
	[SupportByVersion("MSHTML", 4)]
	[EntityType(EntityType.IsInterface), BaseType]
	[TypeId("3050F683-98B5-11CF-BB82-00AA00BDCE0B")]
	public interface ISegment : ICOMObject
	{
		#region Methods

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// </summary>
		/// <param name="pIStart">NetOffice.MSHTMLApi.IMarkupPointer pIStart</param>
		/// <param name="pIEnd">NetOffice.MSHTMLApi.IMarkupPointer pIEnd</param>
		[SupportByVersion("MSHTML", 4)]
		Int32 GetPointers(NetOffice.MSHTMLApi.IMarkupPointer pIStart, NetOffice.MSHTMLApi.IMarkupPointer pIEnd);

		#endregion
	}
}
