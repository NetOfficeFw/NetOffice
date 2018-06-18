using System;
using NetRuntimeSystem = System;
using System.ComponentModel;
using NetOffice.Attributes;

namespace NetOffice.MSHTMLApi
{
	/// <summary>
	/// DispatchInterface IHTMLWindow4 
	/// SupportByVersion MSHTML, 4
	/// </summary>
	[SupportByVersion("MSHTML", 4)]
	[EntityType(EntityType.IsDispatchInterface), BaseType]
	[TypeId("3050F6CF-98B5-11CF-BB82-00AA00BDCE0B")]
	public interface IHTMLWindow4 : IHTMLWindow3
	{
		#region Properties

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// Get
		/// </summary>
		[SupportByVersion("MSHTML", 4)]
		[BaseResult]
		NetOffice.MSHTMLApi.IHTMLFrameBase frameElement { get; }

		#endregion

		#region Methods

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// </summary>
		/// <param name="varArgIn">optional object varArgIn</param>
		[SupportByVersion("MSHTML", 4)]
		object createPopup(object varArgIn);

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// </summary>
		[CustomMethod]
		[SupportByVersion("MSHTML", 4)]
		object createPopup();

		#endregion
	}
}
