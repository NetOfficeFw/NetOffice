using System;
using NetRuntimeSystem = System;
using System.ComponentModel;
using NetOffice.Attributes;

namespace NetOffice.MSHTMLApi
{
	/// <summary>
	/// DispatchInterface IHTMLDataTransfer 
	/// SupportByVersion MSHTML, 4
	/// </summary>
	[SupportByVersion("MSHTML", 4)]
	[EntityType(EntityType.IsDispatchInterface)]
	[TypeId("3050F4B3-98B5-11CF-BB82-00AA00BDCE0B")]
	public interface IHTMLDataTransfer : ICOMObject
	{
		#region Properties

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// Get/Set
		/// </summary>
		[SupportByVersion("MSHTML", 4)]
		string dropEffect { get; set; }

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// Get/Set
		/// </summary>
		[SupportByVersion("MSHTML", 4)]
		string effectAllowed { get; set; }

		#endregion

		#region Methods

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// </summary>
		/// <param name="format">string format</param>
		/// <param name="data">object data</param>
		[SupportByVersion("MSHTML", 4)]
		bool setData(string format, object data);

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// </summary>
		/// <param name="format">string format</param>
		[SupportByVersion("MSHTML", 4)]
		object getData(string format);

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// </summary>
		/// <param name="format">string format</param>
		[SupportByVersion("MSHTML", 4)]
		bool clearData(string format);

		#endregion
	}
}
