using System;
using NetRuntimeSystem = System;
using System.ComponentModel;
using NetOffice.Attributes;

namespace NetOffice.MSHTMLApi
{
	/// <summary>
	/// DispatchInterface IHTMLTable2 
	/// SupportByVersion MSHTML, 4
	/// </summary>
	[SupportByVersion("MSHTML", 4)]
	[EntityType(EntityType.IsDispatchInterface), BaseType]
	[TypeId("3050F4AD-98B5-11CF-BB82-00AA00BDCE0B")]
	public interface IHTMLTable2 : IHTMLTable
	{
		#region Properties

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// Get
		/// </summary>
		[SupportByVersion("MSHTML", 4)]
		[BaseResult]
		NetOffice.MSHTMLApi.IHTMLElementCollection cells { get; }

		#endregion

		#region Methods

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// </summary>
		[SupportByVersion("MSHTML", 4)]
		void firstPage();

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// </summary>
		[SupportByVersion("MSHTML", 4)]
		void lastPage();

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// </summary>
		/// <param name="indexFrom">optional Int32 indexFrom = -1</param>
		/// <param name="indexTo">optional Int32 indexTo = -1</param>
		[SupportByVersion("MSHTML", 4)]
		object moveRow(object indexFrom, object indexTo);

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// </summary>
		[CustomMethod]
		[SupportByVersion("MSHTML", 4)]
		object moveRow();

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// </summary>
		/// <param name="indexFrom">optional Int32 indexFrom = -1</param>
		[CustomMethod]
		[SupportByVersion("MSHTML", 4)]
		object moveRow(object indexFrom);

		#endregion
	}
}
