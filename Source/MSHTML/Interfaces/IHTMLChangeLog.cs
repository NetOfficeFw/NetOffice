using System.Reflection;
using System;
using NetRuntimeSystem = System;
using System.ComponentModel;
using NetOffice.Attributes;

namespace NetOffice.MSHTMLApi
{
	/// <summary>
	/// Interface IHTMLChangeLog 
	/// SupportByVersion MSHTML, 4
	/// </summary>
	[SupportByVersion("MSHTML", 4)]
	[EntityType(EntityType.IsInterface)]
	[TypeId("3050F649-98B5-11CF-BB82-00AA00BDCE0B")]
	public interface IHTMLChangeLog : ICOMObject
	{
		#region Methods

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// </summary>
		/// <param name="pbBuffer">byte pbBuffer</param>
		/// <param name="nBufferSize">Int32 nBufferSize</param>
		/// <param name="pnRecordLength">Int32 pnRecordLength</param>
		[SupportByVersion("MSHTML", 4)]
		Int32 GetNextChange(byte pbBuffer, Int32 nBufferSize, out Int32 pnRecordLength);

		#endregion
	}
}
