using System;
using NetRuntimeSystem = System;
using System.ComponentModel;
using NetOffice.Attributes;

namespace NetOffice.MSHTMLApi
{
	/// <summary>
	/// DispatchInterface DispHTMLAreasCollection 
	/// SupportByVersion MSHTML, 4
	/// </summary>
	[SupportByVersion("MSHTML", 4)]
	[EntityType(EntityType.IsDispatchInterface), BaseType]
	[TypeId("3050F56A-98B5-11CF-BB82-00AA00BDCE0B")]
    [CoClassSource(typeof(NetOffice.MSHTMLApi.HTMLAreasCollection))]
    public interface DispHTMLAreasCollection : ICOMObject
	{
		#region Properties

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// Get/Set
		/// </summary>
		[SupportByVersion("MSHTML", 4)]
		Int32 length { get; set; }

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
		/// <param name="index">optional object index</param>
		[SupportByVersion("MSHTML", 4)]
		object item(object name, object index);

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// </summary>
		[CustomMethod]
		[SupportByVersion("MSHTML", 4)]
		object item();

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// </summary>
		/// <param name="name">optional object name</param>
		[CustomMethod]
		[SupportByVersion("MSHTML", 4)]
		object item(object name);

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// </summary>
		/// <param name="tagName">object tagName</param>
		[SupportByVersion("MSHTML", 4)]
		object tags(object tagName);

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// </summary>
		/// <param name="element">NetOffice.MSHTMLApi.IHTMLElement element</param>
		/// <param name="before">optional object before</param>
		[SupportByVersion("MSHTML", 4)]
		void add(NetOffice.MSHTMLApi.IHTMLElement element, object before);

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// </summary>
		/// <param name="element">NetOffice.MSHTMLApi.IHTMLElement element</param>
		[CustomMethod]
		[SupportByVersion("MSHTML", 4)]
		void add(NetOffice.MSHTMLApi.IHTMLElement element);

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// </summary>
		/// <param name="index">optional Int32 index = -1</param>
		[SupportByVersion("MSHTML", 4)]
		void remove(object index);

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// </summary>
		[CustomMethod]
		[SupportByVersion("MSHTML", 4)]
		void remove();

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// </summary>
		/// <param name="urn">object urn</param>
		[SupportByVersion("MSHTML", 4)]
		object urns(object urn);

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// </summary>
		/// <param name="name">string name</param>
		[SupportByVersion("MSHTML", 4)]
		object namedItem(string name);

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// </summary>
		/// <param name="index">Int32 index</param>
		[SupportByVersion("MSHTML", 4)]
		[BaseResult]
		NetOffice.MSHTMLApi.IHTMLElement2 ie8_item(Int32 index);

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// </summary>
		/// <param name="name">string name</param>
		[SupportByVersion("MSHTML", 4)]
		[BaseResult]
		NetOffice.MSHTMLApi.IHTMLElement2 ie8_namedItem(string name);

		#endregion
	}
}
