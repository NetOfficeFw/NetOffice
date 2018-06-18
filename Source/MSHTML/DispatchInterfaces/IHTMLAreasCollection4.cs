using System;
using NetRuntimeSystem = System;
using System.ComponentModel;
using NetOffice.Attributes;

namespace NetOffice.MSHTMLApi
{
	/// <summary>
	/// DispatchInterface IHTMLAreasCollection4 
	/// SupportByVersion MSHTML, 4
	/// </summary>
	[SupportByVersion("MSHTML", 4)]
	[EntityType(EntityType.IsDispatchInterface)]
	[TypeId("30510492-98B5-11CF-BB82-00AA00BDCE0B")]
	public interface IHTMLAreasCollection4 : IHTMLAreasCollection3
	{
		#region Methods

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// </summary>
		/// <param name="index">Int32 index</param>
		[SupportByVersion("MSHTML", 4)]
		[BaseResult]
		NetOffice.MSHTMLApi.IHTMLElement2 item(Int32 index);

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// </summary>
		/// <param name="name">string name</param>
		[SupportByVersion("MSHTML", 4)]
		[BaseResult]
        new NetOffice.MSHTMLApi.IHTMLElement2 namedItem(string name);

		#endregion
	}
}
