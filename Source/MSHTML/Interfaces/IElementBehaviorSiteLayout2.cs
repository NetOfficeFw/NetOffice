using System.Reflection;
using System;
using NetRuntimeSystem = System;
using System.ComponentModel;
using NetOffice.Attributes;

namespace NetOffice.MSHTMLApi
{
	/// <summary>
	/// Interface IElementBehaviorSiteLayout2 
	/// SupportByVersion MSHTML, 4
	/// </summary>
	[SupportByVersion("MSHTML", 4)]
	[EntityType(EntityType.IsInterface)]
	[TypeId("3050F847-98B5-11CF-BB82-00AA00BDCE0B")]
	public interface IElementBehaviorSiteLayout2 : ICOMObject
	{
		#region Methods

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// </summary>
		/// <param name="plf">tagLOGFONTW plf</param>
		[SupportByVersion("MSHTML", 4)]
		Int32 GetFontInfo(out tagLOGFONTW plf);

		#endregion
	}
}
