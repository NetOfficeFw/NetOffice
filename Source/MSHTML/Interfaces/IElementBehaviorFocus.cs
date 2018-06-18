using System;
using NetRuntimeSystem = System;
using System.ComponentModel;
using NetOffice.Attributes;

namespace NetOffice.MSHTMLApi
{
	/// <summary>
	/// Interface IElementBehaviorFocus 
	/// SupportByVersion MSHTML, 4
	/// </summary>
	[SupportByVersion("MSHTML", 4)]
	[EntityType(EntityType.IsInterface)]
	[TypeId("3050F6B6-98B5-11CF-BB82-00AA00BDCE0B")]
	public interface IElementBehaviorFocus : ICOMObject
	{
		#region Methods

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// </summary>
		/// <param name="pRect">tagRECT pRect</param>
		[SupportByVersion("MSHTML", 4)]
		Int32 GetFocusRect(tagRECT pRect);

		#endregion
	}
}
