using System.Reflection;
using System;
using NetRuntimeSystem = System;
using System.ComponentModel;
using NetOffice.Attributes;

namespace NetOffice.MSHTMLApi
{
	/// <summary>
	/// Interface ISegmentList 
	/// SupportByVersion MSHTML, 4
	/// </summary>
	[SupportByVersion("MSHTML", 4)]
	[EntityType(EntityType.IsInterface)]
	[TypeId("3050F605-98B5-11CF-BB82-00AA00BDCE0B")]
	public interface ISegmentList : ICOMObject
	{
		#region Methods

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// </summary>
		/// <param name="ppIIter">NetOffice.MSHTMLApi.ISegmentListIterator ppIIter</param>
		[SupportByVersion("MSHTML", 4)]
		Int32 CreateIterator(out NetOffice.MSHTMLApi.ISegmentListIterator ppIIter);

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// </summary>
		/// <param name="peType">NetOffice.MSHTMLApi.Enums._SELECTION_TYPE peType</param>
		[SupportByVersion("MSHTML", 4)]
		Int32 GetType(out NetOffice.MSHTMLApi.Enums._SELECTION_TYPE peType);

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// </summary>
		/// <param name="pfEmpty">Int32 pfEmpty</param>
		[SupportByVersion("MSHTML", 4)]
		Int32 IsEmpty(out Int32 pfEmpty);

		#endregion
	}
}
