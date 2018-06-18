using System;
using NetRuntimeSystem = System;
using System.ComponentModel;
using NetOffice.Attributes;

namespace NetOffice.MSHTMLApi
{
	/// <summary>
	/// Interface IHTMLSelectElementEx 
	/// SupportByVersion MSHTML, 4
	/// </summary>
	[SupportByVersion("MSHTML", 4)]
	[EntityType(EntityType.IsInterface)]
	[TypeId("3050F2D1-98B5-11CF-BB82-00AA00BDCE0B")]
	public interface IHTMLSelectElementEx : ICOMObject
	{
		#region Methods

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// </summary>
		/// <param name="fShow">Int32 fShow</param>
		[SupportByVersion("MSHTML", 4)]
		Int32 ShowDropdown(Int32 fShow);

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// </summary>
		/// <param name="lFlags">Int32 lFlags</param>
		[SupportByVersion("MSHTML", 4)]
		Int32 SetSelectExFlags(Int32 lFlags);

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// </summary>
		[SupportByVersion("MSHTML", 4)]
		Int32 GetSelectExFlags();

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// </summary>
		[SupportByVersion("MSHTML", 4)]
		Int32 GetDropdownOpen();

		#endregion
	}
}
