using System;
using NetRuntimeSystem = System;
using System.ComponentModel;
using NetOffice.Attributes;

namespace NetOffice.MSHTMLApi
{
	/// <summary>
	/// DispatchInterface DispHTMLAttributeCollection 
	/// SupportByVersion MSHTML, 4
	/// </summary>
	[SupportByVersion("MSHTML", 4)]
	[EntityType(EntityType.IsDispatchInterface), BaseType]
	[TypeId("3050F56C-98B5-11CF-BB82-00AA00BDCE0B")]
    [CoClassSource(typeof(NetOffice.MSHTMLApi.HTMLAttributeCollection))]
    public interface DispHTMLAttributeCollection : ICOMObject
	{
		#region Properties

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// Get
		/// </summary>
		[SupportByVersion("MSHTML", 4)]
		Int32 length { get; }

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// Get
		/// Unknown COM Proxy
		/// </summary>
		[SupportByVersion("MSHTML", 4), ProxyResult]
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		object _newEnum { get; }

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// Get
		/// </summary>
		[SupportByVersion("MSHTML", 4)]
		Int32 ie8_length { get; }

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// Get
		/// Unknown COM Proxy
		/// </summary>
		[SupportByVersion("MSHTML", 4), ProxyResult]
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		object constructor { get; }

		#endregion

		#region Methods

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// </summary>
		/// <param name="name">optional object name</param>
		[SupportByVersion("MSHTML", 4)]
		object item(object name);

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// </summary>
		[CustomMethod]
		[SupportByVersion("MSHTML", 4)]
		object item();

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// </summary>
		/// <param name="bstrName">string bstrName</param>
		[SupportByVersion("MSHTML", 4)]
		[BaseResult]
		NetOffice.MSHTMLApi.IHTMLDOMAttribute getNamedItem(string bstrName);

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// </summary>
		/// <param name="ppNode">NetOffice.MSHTMLApi.IHTMLDOMAttribute ppNode</param>
		[SupportByVersion("MSHTML", 4)]
		[BaseResult]
		NetOffice.MSHTMLApi.IHTMLDOMAttribute setNamedItem(NetOffice.MSHTMLApi.IHTMLDOMAttribute ppNode);

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// </summary>
		/// <param name="bstrName">string bstrName</param>
		[SupportByVersion("MSHTML", 4)]
		[BaseResult]
		NetOffice.MSHTMLApi.IHTMLDOMAttribute removeNamedItem(string bstrName);

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// </summary>
		/// <param name="bstrName">string bstrName</param>
		[SupportByVersion("MSHTML", 4)]
		[BaseResult]
		NetOffice.MSHTMLApi.IHTMLDOMAttribute ie8_getNamedItem(string bstrName);

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// </summary>
		/// <param name="pNodeIn">NetOffice.MSHTMLApi.IHTMLDOMAttribute pNodeIn</param>
		[SupportByVersion("MSHTML", 4)]
		[BaseResult]
		NetOffice.MSHTMLApi.IHTMLDOMAttribute ie8_setNamedItem(NetOffice.MSHTMLApi.IHTMLDOMAttribute pNodeIn);

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// </summary>
		/// <param name="bstrName">string bstrName</param>
		[SupportByVersion("MSHTML", 4)]
		[BaseResult]
		NetOffice.MSHTMLApi.IHTMLDOMAttribute ie8_removeNamedItem(string bstrName);

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// </summary>
		/// <param name="index">Int32 index</param>
		[SupportByVersion("MSHTML", 4)]
		[BaseResult]
		NetOffice.MSHTMLApi.IHTMLDOMAttribute ie8_item(Int32 index);

		#endregion
	}
}
