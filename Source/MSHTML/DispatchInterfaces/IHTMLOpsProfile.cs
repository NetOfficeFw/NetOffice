using System;
using NetRuntimeSystem = System;
using System.ComponentModel;
using NetOffice.Attributes;

namespace NetOffice.MSHTMLApi
{
	/// <summary>
	/// DispatchInterface IHTMLOpsProfile 
	/// SupportByVersion MSHTML, 4
	/// </summary>
	[SupportByVersion("MSHTML", 4)]
	[EntityType(EntityType.IsDispatchInterface), BaseType]
	[TypeId("3050F401-98B5-11CF-BB82-00AA00BDCE0B")]
    [CoClassSource(typeof(NetOffice.MSHTMLApi.COpsProfile))]
    public interface IHTMLOpsProfile : ICOMObject
	{
		#region Methods

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// </summary>
		/// <param name="name">string name</param>
		/// <param name="reserved">optional object reserved</param>
		[SupportByVersion("MSHTML", 4)]
		bool addRequest(string name, object reserved);

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// </summary>
		/// <param name="name">string name</param>
		[CustomMethod]
		[SupportByVersion("MSHTML", 4)]
		bool addRequest(string name);

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// </summary>
		[SupportByVersion("MSHTML", 4)]
		void clearRequest();

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// </summary>
		/// <param name="usage">object usage</param>
		/// <param name="fname">optional object fname</param>
		/// <param name="domain">optional object domain</param>
		/// <param name="path">optional object path</param>
		/// <param name="expire">optional object expire</param>
		/// <param name="reserved">optional object reserved</param>
		[SupportByVersion("MSHTML", 4)]
		void doRequest(object usage, object fname, object domain, object path, object expire, object reserved);

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// </summary>
		/// <param name="usage">object usage</param>
		[CustomMethod]
		[SupportByVersion("MSHTML", 4)]
		void doRequest(object usage);

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// </summary>
		/// <param name="usage">object usage</param>
		/// <param name="fname">optional object fname</param>
		[CustomMethod]
		[SupportByVersion("MSHTML", 4)]
		void doRequest(object usage, object fname);

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// </summary>
		/// <param name="usage">object usage</param>
		/// <param name="fname">optional object fname</param>
		/// <param name="domain">optional object domain</param>
		[CustomMethod]
		[SupportByVersion("MSHTML", 4)]
		void doRequest(object usage, object fname, object domain);

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// </summary>
		/// <param name="usage">object usage</param>
		/// <param name="fname">optional object fname</param>
		/// <param name="domain">optional object domain</param>
		/// <param name="path">optional object path</param>
		[CustomMethod]
		[SupportByVersion("MSHTML", 4)]
		void doRequest(object usage, object fname, object domain, object path);

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// </summary>
		/// <param name="usage">object usage</param>
		/// <param name="fname">optional object fname</param>
		/// <param name="domain">optional object domain</param>
		/// <param name="path">optional object path</param>
		/// <param name="expire">optional object expire</param>
		[CustomMethod]
		[SupportByVersion("MSHTML", 4)]
		void doRequest(object usage, object fname, object domain, object path, object expire);

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// </summary>
		/// <param name="name">string name</param>
		[SupportByVersion("MSHTML", 4)]
		string getAttribute(string name);

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// </summary>
		/// <param name="name">string name</param>
		/// <param name="value">string value</param>
		/// <param name="prefs">optional object prefs</param>
		[SupportByVersion("MSHTML", 4)]
		bool setAttribute(string name, string value, object prefs);

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// </summary>
		/// <param name="name">string name</param>
		/// <param name="value">string value</param>
		[CustomMethod]
		[SupportByVersion("MSHTML", 4)]
		bool setAttribute(string name, string value);

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// </summary>
		[SupportByVersion("MSHTML", 4)]
		bool commitChanges();

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// </summary>
		/// <param name="name">string name</param>
		/// <param name="reserved">optional object reserved</param>
		[SupportByVersion("MSHTML", 4)]
		bool addReadRequest(string name, object reserved);

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// </summary>
		/// <param name="name">string name</param>
		[CustomMethod]
		[SupportByVersion("MSHTML", 4)]
		bool addReadRequest(string name);

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// </summary>
		/// <param name="usage">object usage</param>
		/// <param name="fname">optional object fname</param>
		/// <param name="domain">optional object domain</param>
		/// <param name="path">optional object path</param>
		/// <param name="expire">optional object expire</param>
		/// <param name="reserved">optional object reserved</param>
		[SupportByVersion("MSHTML", 4)]
		void doReadRequest(object usage, object fname, object domain, object path, object expire, object reserved);

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// </summary>
		/// <param name="usage">object usage</param>
		[CustomMethod]
		[SupportByVersion("MSHTML", 4)]
		void doReadRequest(object usage);

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// </summary>
		/// <param name="usage">object usage</param>
		/// <param name="fname">optional object fname</param>
		[CustomMethod]
		[SupportByVersion("MSHTML", 4)]
		void doReadRequest(object usage, object fname);

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// </summary>
		/// <param name="usage">object usage</param>
		/// <param name="fname">optional object fname</param>
		/// <param name="domain">optional object domain</param>
		[CustomMethod]
		[SupportByVersion("MSHTML", 4)]
		void doReadRequest(object usage, object fname, object domain);

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// </summary>
		/// <param name="usage">object usage</param>
		/// <param name="fname">optional object fname</param>
		/// <param name="domain">optional object domain</param>
		/// <param name="path">optional object path</param>
		[CustomMethod]
		[SupportByVersion("MSHTML", 4)]
		void doReadRequest(object usage, object fname, object domain, object path);

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// </summary>
		/// <param name="usage">object usage</param>
		/// <param name="fname">optional object fname</param>
		/// <param name="domain">optional object domain</param>
		/// <param name="path">optional object path</param>
		/// <param name="expire">optional object expire</param>
		[CustomMethod]
		[SupportByVersion("MSHTML", 4)]
		void doReadRequest(object usage, object fname, object domain, object path, object expire);

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// </summary>
		[SupportByVersion("MSHTML", 4)]
		bool doWriteRequest();

		#endregion
	}
}
