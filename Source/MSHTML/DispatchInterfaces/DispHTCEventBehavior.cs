using System;
using NetRuntimeSystem = System;
using System.ComponentModel;
using NetOffice.Attributes;

namespace NetOffice.MSHTMLApi
{
	/// <summary>
	/// DispatchInterface DispHTCEventBehavior 
	/// SupportByVersion MSHTML, 4
	/// </summary>
	[SupportByVersion("MSHTML", 4)]
	[EntityType(EntityType.IsDispatchInterface), BaseType]
	[TypeId("3050F574-98B5-11CF-BB82-00AA00BDCE0B")]
    [CoClassSource(typeof(NetOffice.MSHTMLApi.HTCEventBehavior))]
    public interface DispHTCEventBehavior : ICOMObject
	{
		#region Methods

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// </summary>
		/// <param name="pVar">NetOffice.MSHTMLApi.IHTMLEventObj pVar</param>
		[SupportByVersion("MSHTML", 4)]
		void fire(NetOffice.MSHTMLApi.IHTMLEventObj pVar);

		#endregion
	}
}
