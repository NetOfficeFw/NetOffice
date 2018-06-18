using System;
using NetRuntimeSystem = System;
using System.ComponentModel;
using NetOffice.Attributes;

namespace NetOffice.MSHTMLApi
{
	/// <summary>
	/// DispatchInterface DispHTMLXMLHttpRequest 
	/// SupportByVersion MSHTML, 4
	/// </summary>
	[SupportByVersion("MSHTML", 4)]
	[EntityType(EntityType.IsDispatchInterface), BaseType]
	[TypeId("3050F596-98B5-11CF-BB82-00AA00BDCE0B")]
    [CoClassSource(typeof(NetOffice.MSHTMLApi.IHTMLXMLHttpRequest))]
    public interface DispHTMLXMLHttpRequest : ICOMObject
	{
		#region Properties

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// Get
		/// </summary>
		[SupportByVersion("MSHTML", 4)]
		Int32 readyState { get; }

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// Get
		/// </summary>
		[SupportByVersion("MSHTML", 4)]
		object responseBody { get; }

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// Get
		/// </summary>
		[SupportByVersion("MSHTML", 4)]
		string responseText { get; }

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// Get
		/// Unknown COM Proxy
		/// </summary>
		[SupportByVersion("MSHTML", 4), ProxyResult]
		object responseXML { get; }

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// Get
		/// </summary>
		[SupportByVersion("MSHTML", 4)]
		Int32 status { get; }

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// Get
		/// </summary>
		[SupportByVersion("MSHTML", 4)]
		string statusText { get; }

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// Get/Set
		/// </summary>
		[SupportByVersion("MSHTML", 4)]
		object onreadystatechange { get; set; }

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// Get/Set
		/// </summary>
		[SupportByVersion("MSHTML", 4)]
		Int32 timeout { get; set; }

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// Get/Set
		/// </summary>
		[SupportByVersion("MSHTML", 4)]
		object ontimeout { get; set; }

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
		[SupportByVersion("MSHTML", 4)]
		void abort();

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// </summary>
		/// <param name="bstrMethod">string bstrMethod</param>
		/// <param name="bstrUrl">string bstrUrl</param>
		/// <param name="varAsync">object varAsync</param>
		/// <param name="varUser">optional object varUser</param>
		/// <param name="varPassword">optional object varPassword</param>
		[SupportByVersion("MSHTML", 4)]
		void open(string bstrMethod, string bstrUrl, object varAsync, object varUser, object varPassword);

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// </summary>
		/// <param name="bstrMethod">string bstrMethod</param>
		/// <param name="bstrUrl">string bstrUrl</param>
		/// <param name="varAsync">object varAsync</param>
		[CustomMethod]
		[SupportByVersion("MSHTML", 4)]
		void open(string bstrMethod, string bstrUrl, object varAsync);

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// </summary>
		/// <param name="bstrMethod">string bstrMethod</param>
		/// <param name="bstrUrl">string bstrUrl</param>
		/// <param name="varAsync">object varAsync</param>
		/// <param name="varUser">optional object varUser</param>
		[CustomMethod]
		[SupportByVersion("MSHTML", 4)]
		void open(string bstrMethod, string bstrUrl, object varAsync, object varUser);

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// </summary>
		/// <param name="varBody">optional object varBody</param>
		[SupportByVersion("MSHTML", 4)]
		void send(object varBody);

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// </summary>
		[CustomMethod]
		[SupportByVersion("MSHTML", 4)]
		void send();

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// </summary>
		[SupportByVersion("MSHTML", 4)]
		string getAllResponseHeaders();

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// </summary>
		/// <param name="bstrHeader">string bstrHeader</param>
		[SupportByVersion("MSHTML", 4)]
		string getResponseHeader(string bstrHeader);

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// </summary>
		/// <param name="bstrHeader">string bstrHeader</param>
		/// <param name="bstrValue">string bstrValue</param>
		[SupportByVersion("MSHTML", 4)]
		void setRequestHeader(string bstrHeader, string bstrValue);

		#endregion
	}
}
