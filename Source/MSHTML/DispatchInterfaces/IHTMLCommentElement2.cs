using System;
using NetRuntimeSystem = System;
using System.ComponentModel;
using NetOffice.Attributes;

namespace NetOffice.MSHTMLApi
{
	/// <summary>
	/// DispatchInterface IHTMLCommentElement2 
	/// SupportByVersion MSHTML, 4
	/// </summary>
	[SupportByVersion("MSHTML", 4)]
	[EntityType(EntityType.IsDispatchInterface)]
	[TypeId("3050F813-98B5-11CF-BB82-00AA00BDCE0B")]
	public interface IHTMLCommentElement2 : IHTMLCommentElement
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

		#endregion

		#region Methods

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

		#endregion
	}
}
