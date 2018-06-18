using System;
using NetRuntimeSystem = System;
using System.ComponentModel;
using NetOffice.Attributes;

namespace NetOffice.MSHTMLApi
{
	/// <summary>
	/// Interface IHTMLChangePlayback 
	/// SupportByVersion MSHTML, 4
	/// </summary>
	[SupportByVersion("MSHTML", 4)]
	[EntityType(EntityType.IsInterface)]
	[TypeId("3050F6E0-98B5-11CF-BB82-00AA00BDCE0B")]
	public interface IHTMLChangePlayback : ICOMObject
	{
		#region Methods

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// </summary>
		/// <param name="pbRecord">byte pbRecord</param>
		/// <param name="fForward">Int32 fForward</param>
		[SupportByVersion("MSHTML", 4)]
		Int32 ExecChange(byte pbRecord, Int32 fForward);

		#endregion
	}
}
