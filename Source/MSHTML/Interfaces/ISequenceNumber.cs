using System.Reflection;
using System;
using NetRuntimeSystem = System;
using System.ComponentModel;
using NetOffice.Attributes;

namespace NetOffice.MSHTMLApi
{
	/// <summary>
	/// Interface ISequenceNumber 
	/// SupportByVersion MSHTML, 4
	/// </summary>
	[SupportByVersion("MSHTML", 4)]
	[EntityType(EntityType.IsInterface)]
	[TypeId("3050F6C1-98B5-11CF-BB82-00AA00BDCE0B")]
	public interface ISequenceNumber : ICOMObject
	{
		#region Methods

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// </summary>
		/// <param name="nCurrent">Int32 nCurrent</param>
		/// <param name="pnNew">Int32 pnNew</param>
		[SupportByVersion("MSHTML", 4)]
		Int32 GetSequenceNumber(Int32 nCurrent, out Int32 pnNew);

		#endregion
	}
}
