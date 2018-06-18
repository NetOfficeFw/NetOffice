using System;
using NetRuntimeSystem = System;
using System.ComponentModel;
using NetOffice.Attributes;

namespace NetOffice.MSHTMLApi
{
	/// <summary>
	/// DispatchInterface DispHTMLDOMTextNode 
	/// SupportByVersion MSHTML, 4
	/// </summary>
	[SupportByVersion("MSHTML", 4)]
	[EntityType(EntityType.IsDispatchInterface), BaseType]
	[TypeId("3050F565-98B5-11CF-BB82-00AA00BDCE0B")]
    [CoClassSource(typeof(NetOffice.MSHTMLApi.HTMLDOMTextNode))]
    public interface DispHTMLDOMTextNode : ICOMObject
	{
		#region Properties

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// Get/Set
		/// </summary>
		[SupportByVersion("MSHTML", 4)]
		string data { get; set; }

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// Get
		/// </summary>
		[SupportByVersion("MSHTML", 4)]
		Int32 length { get; }

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// Get
		/// </summary>
		[SupportByVersion("MSHTML", 4)]
		Int32 nodeType { get; }

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// Get
		/// </summary>
		[SupportByVersion("MSHTML", 4)]
		[BaseResult]
		NetOffice.MSHTMLApi.IHTMLDOMNode parentNode { get; }

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// Get
		/// Unknown COM Proxy
		/// </summary>
		[SupportByVersion("MSHTML", 4), ProxyResult]
		object childNodes { get; }

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// Get
		/// Unknown COM Proxy
		/// </summary>
		[SupportByVersion("MSHTML", 4), ProxyResult]
		object attributes { get; }

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// Get
		/// </summary>
		[SupportByVersion("MSHTML", 4)]
		string nodeName { get; }

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// Get/Set
		/// </summary>
		[SupportByVersion("MSHTML", 4)]
		object nodeValue { get; set; }

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// Get
		/// </summary>
		[SupportByVersion("MSHTML", 4)]
		[BaseResult]
		NetOffice.MSHTMLApi.IHTMLDOMNode firstChild { get; }

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// Get
		/// </summary>
		[SupportByVersion("MSHTML", 4)]
		[BaseResult]
		NetOffice.MSHTMLApi.IHTMLDOMNode lastChild { get; }

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// Get
		/// </summary>
		[SupportByVersion("MSHTML", 4)]
		[BaseResult]
		NetOffice.MSHTMLApi.IHTMLDOMNode previousSibling { get; }

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// Get
		/// </summary>
		[SupportByVersion("MSHTML", 4)]
		[BaseResult]
		NetOffice.MSHTMLApi.IHTMLDOMNode nextSibling { get; }

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// Get
		/// Unknown COM Proxy
		/// </summary>
		[SupportByVersion("MSHTML", 4), ProxyResult]
		object ownerDocument { get; }

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
		string toString();

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// </summary>
		/// <param name="offset">Int32 offset</param>
		[SupportByVersion("MSHTML", 4)]
		[BaseResult]
		NetOffice.MSHTMLApi.IHTMLDOMNode splitText(Int32 offset);

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// </summary>
		/// <param name="offset">Int32 offset</param>
		/// <param name="count">Int32 count</param>
		[SupportByVersion("MSHTML", 4)]
		string substringData(Int32 offset, Int32 count);

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// </summary>
		/// <param name="bstrstring">string bstrstring</param>
		[SupportByVersion("MSHTML", 4)]
		void appendData(string bstrstring);

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// </summary>
		/// <param name="offset">Int32 offset</param>
		/// <param name="bstrstring">string bstrstring</param>
		[SupportByVersion("MSHTML", 4)]
		void insertData(Int32 offset, string bstrstring);

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// </summary>
		/// <param name="offset">Int32 offset</param>
		/// <param name="count">Int32 count</param>
		[SupportByVersion("MSHTML", 4)]
		void deleteData(Int32 offset, Int32 count);

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// </summary>
		/// <param name="offset">Int32 offset</param>
		/// <param name="count">Int32 count</param>
		/// <param name="bstrstring">string bstrstring</param>
		[SupportByVersion("MSHTML", 4)]
		void replaceData(Int32 offset, Int32 count, string bstrstring);

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// </summary>
		[SupportByVersion("MSHTML", 4)]
		bool hasChildNodes();

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// </summary>
		/// <param name="newChild">NetOffice.MSHTMLApi.IHTMLDOMNode newChild</param>
		/// <param name="refChild">optional object refChild</param>
		[SupportByVersion("MSHTML", 4)]
		[BaseResult]
		NetOffice.MSHTMLApi.IHTMLDOMNode insertBefore(NetOffice.MSHTMLApi.IHTMLDOMNode newChild, object refChild);

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// </summary>
		/// <param name="newChild">NetOffice.MSHTMLApi.IHTMLDOMNode newChild</param>
		[CustomMethod]
		[BaseResult]
		[SupportByVersion("MSHTML", 4)]
		NetOffice.MSHTMLApi.IHTMLDOMNode insertBefore(NetOffice.MSHTMLApi.IHTMLDOMNode newChild);

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// </summary>
		/// <param name="oldChild">NetOffice.MSHTMLApi.IHTMLDOMNode oldChild</param>
		[SupportByVersion("MSHTML", 4)]
		[BaseResult]
		NetOffice.MSHTMLApi.IHTMLDOMNode removeChild(NetOffice.MSHTMLApi.IHTMLDOMNode oldChild);

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// </summary>
		/// <param name="newChild">NetOffice.MSHTMLApi.IHTMLDOMNode newChild</param>
		/// <param name="oldChild">NetOffice.MSHTMLApi.IHTMLDOMNode oldChild</param>
		[SupportByVersion("MSHTML", 4)]
		[BaseResult]
		NetOffice.MSHTMLApi.IHTMLDOMNode replaceChild(NetOffice.MSHTMLApi.IHTMLDOMNode newChild, NetOffice.MSHTMLApi.IHTMLDOMNode oldChild);

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// </summary>
		/// <param name="fDeep">bool fDeep</param>
		[SupportByVersion("MSHTML", 4)]
		[BaseResult]
		NetOffice.MSHTMLApi.IHTMLDOMNode cloneNode(bool fDeep);

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// </summary>
		/// <param name="fDeep">optional bool fDeep = false</param>
		[SupportByVersion("MSHTML", 4)]
		[BaseResult]
		NetOffice.MSHTMLApi.IHTMLDOMNode removeNode(object fDeep);

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// </summary>
		[CustomMethod]
		[BaseResult]
		[SupportByVersion("MSHTML", 4)]
		NetOffice.MSHTMLApi.IHTMLDOMNode removeNode();

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// </summary>
		/// <param name="otherNode">NetOffice.MSHTMLApi.IHTMLDOMNode otherNode</param>
		[SupportByVersion("MSHTML", 4)]
		[BaseResult]
		NetOffice.MSHTMLApi.IHTMLDOMNode swapNode(NetOffice.MSHTMLApi.IHTMLDOMNode otherNode);

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// </summary>
		/// <param name="replacement">NetOffice.MSHTMLApi.IHTMLDOMNode replacement</param>
		[SupportByVersion("MSHTML", 4)]
		[BaseResult]
		NetOffice.MSHTMLApi.IHTMLDOMNode replaceNode(NetOffice.MSHTMLApi.IHTMLDOMNode replacement);

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// </summary>
		/// <param name="newChild">NetOffice.MSHTMLApi.IHTMLDOMNode newChild</param>
		[SupportByVersion("MSHTML", 4)]
		[BaseResult]
		NetOffice.MSHTMLApi.IHTMLDOMNode appendChild(NetOffice.MSHTMLApi.IHTMLDOMNode newChild);

		#endregion
	}
}
