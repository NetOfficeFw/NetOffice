using System;
using NetRuntimeSystem = System;
using System.ComponentModel;
using NetOffice.Attributes;

namespace NetOffice.MSHTMLApi
{
	/// <summary>
	/// DispatchInterface IHTMLImageElementFactory 
	/// SupportByVersion MSHTML, 4
	/// </summary>
	[SupportByVersion("MSHTML", 4)]
	[EntityType(EntityType.IsDispatchInterface), BaseType]
	[TypeId("3050F38E-98B5-11CF-BB82-00AA00BDCE0B")]
    [CoClassSource(typeof(NetOffice.MSHTMLApi.HTMLImageElementFactory))]
    public interface IHTMLImageElementFactory : ICOMObject
	{
		#region Methods

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// </summary>
		/// <param name="width">optional object width</param>
		/// <param name="height">optional object height</param>
		[SupportByVersion("MSHTML", 4)]
		[BaseResult]
		NetOffice.MSHTMLApi.IHTMLImgElement create(object width, object height);

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// </summary>
		[CustomMethod]
		[BaseResult]
		[SupportByVersion("MSHTML", 4)]
		NetOffice.MSHTMLApi.IHTMLImgElement create();

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// </summary>
		/// <param name="width">optional object width</param>
		[CustomMethod]
		[BaseResult]
		[SupportByVersion("MSHTML", 4)]
		NetOffice.MSHTMLApi.IHTMLImgElement create(object width);

		#endregion
	}
}
