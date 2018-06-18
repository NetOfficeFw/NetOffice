using System;
using NetRuntimeSystem = System;
using System.ComponentModel;
using NetOffice.Attributes;

namespace NetOffice.MSHTMLApi
{
	/// <summary>
	/// DispatchInterface IHTMLWindow6 
	/// SupportByVersion MSHTML, 4
	/// </summary>
	[SupportByVersion("MSHTML", 4)]
	[EntityType(EntityType.IsDispatchInterface)]
	[TypeId("30510453-98B5-11CF-BB82-00AA00BDCE0B")]
	public interface IHTMLWindow6 : IHTMLWindow5
	{
		#region Properties

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// Get/Set
		/// </summary>
		[SupportByVersion("MSHTML", 4)]
		object XDomainRequest { get; set; }

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// Get
		/// </summary>
		[SupportByVersion("MSHTML", 4)]
		NetOffice.MSHTMLApi.IHTMLStorage sessionStorage { get; }

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// Get
		/// </summary>
		[SupportByVersion("MSHTML", 4)]
		NetOffice.MSHTMLApi.IHTMLStorage localStorage { get; }

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// Get/Set
		/// </summary>
		[SupportByVersion("MSHTML", 4)]
		object onhashchange { get; set; }

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// Get
		/// </summary>
		[SupportByVersion("MSHTML", 4)]
		Int32 maxConnectionsPerServer { get; }

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// Get/Set
		/// </summary>
		[SupportByVersion("MSHTML", 4)]
		object onmessage { get; set; }

		#endregion

		#region Methods

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// </summary>
		/// <param name="msg">string msg</param>
		/// <param name="targetOrigin">optional object targetOrigin</param>
		[SupportByVersion("MSHTML", 4)]
		void postMessage(string msg, object targetOrigin);

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// </summary>
		/// <param name="msg">string msg</param>
		[CustomMethod]
		[SupportByVersion("MSHTML", 4)]
		void postMessage(string msg);

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// </summary>
		/// <param name="bstrHTML">string bstrHTML</param>
		[SupportByVersion("MSHTML", 4)]
		string toStaticHTML(string bstrHTML);

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// </summary>
		/// <param name="bstrProfilerMarkName">string bstrProfilerMarkName</param>
		[SupportByVersion("MSHTML", 4)]
		void msWriteProfilerMark(string bstrProfilerMarkName);

		#endregion
	}
}
