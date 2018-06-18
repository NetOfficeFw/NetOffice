using System.Reflection;
using System;
using NetRuntimeSystem = System;
using System.ComponentModel;
using NetOffice.Attributes;

namespace NetOffice.MSHTMLApi
{
	/// <summary>
	/// Interface IIMEServices 
	/// SupportByVersion MSHTML, 4
	/// </summary>
	[SupportByVersion("MSHTML", 4)]
	[EntityType(EntityType.IsInterface)]
	[TypeId("3050F6CA-98B5-11CF-BB82-00AA00BDCE0B")]
	public interface IIMEServices : ICOMObject
	{
		#region Methods

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// </summary>
		/// <param name="ppActiveIMM">NetOffice.MSHTMLApi.IActiveIMMApp ppActiveIMM</param>
		[SupportByVersion("MSHTML", 4)]
		Int32 GetActiveIMM(out NetOffice.MSHTMLApi.IActiveIMMApp ppActiveIMM);

		#endregion
	}
}
