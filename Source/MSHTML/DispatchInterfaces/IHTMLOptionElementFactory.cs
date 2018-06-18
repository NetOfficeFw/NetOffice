using System;
using NetRuntimeSystem = System;
using System.ComponentModel;
using NetOffice.Attributes;

namespace NetOffice.MSHTMLApi
{
	/// <summary>
	/// DispatchInterface IHTMLOptionElementFactory 
	/// SupportByVersion MSHTML, 4
	/// </summary>
	[SupportByVersion("MSHTML", 4)]
	[EntityType(EntityType.IsDispatchInterface), BaseType]
	[TypeId("3050F38C-98B5-11CF-BB82-00AA00BDCE0B")]
    [CoClassSource(typeof(NetOffice.MSHTMLApi.HTMLOptionElementFactory))]
    public interface IHTMLOptionElementFactory : ICOMObject
	{
		#region Methods

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// </summary>
		/// <param name="text">optional object text</param>
		/// <param name="value">optional object value</param>
		/// <param name="defaultSelected">optional object defaultSelected</param>
		/// <param name="selected">optional object selected</param>
		[SupportByVersion("MSHTML", 4)]
		[BaseResult]
		NetOffice.MSHTMLApi.IHTMLOptionElement create(object text, object value, object defaultSelected, object selected);

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// </summary>
		[CustomMethod]
		[BaseResult]
		[SupportByVersion("MSHTML", 4)]
		NetOffice.MSHTMLApi.IHTMLOptionElement create();

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// </summary>
		/// <param name="text">optional object text</param>
		[CustomMethod]
		[BaseResult]
		[SupportByVersion("MSHTML", 4)]
		NetOffice.MSHTMLApi.IHTMLOptionElement create(object text);

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// </summary>
		/// <param name="text">optional object text</param>
		/// <param name="value">optional object value</param>
		[CustomMethod]
		[BaseResult]
		[SupportByVersion("MSHTML", 4)]
		NetOffice.MSHTMLApi.IHTMLOptionElement create(object text, object value);

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// </summary>
		/// <param name="text">optional object text</param>
		/// <param name="value">optional object value</param>
		/// <param name="defaultSelected">optional object defaultSelected</param>
		[CustomMethod]
		[BaseResult]
		[SupportByVersion("MSHTML", 4)]
		NetOffice.MSHTMLApi.IHTMLOptionElement create(object text, object value, object defaultSelected);

		#endregion
	}
}
