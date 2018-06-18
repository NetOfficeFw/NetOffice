using System;
using NetRuntimeSystem = System;
using System.ComponentModel;
using NetOffice.Attributes;

namespace NetOffice.MSHTMLApi
{
	/// <summary>
	/// Interface IHTMLPainterOverlay 
	/// SupportByVersion MSHTML, 4
	/// </summary>
	[SupportByVersion("MSHTML", 4)]
	[EntityType(EntityType.IsInterface)]
	[TypeId("3050F7E3-98B5-11CF-BB82-00AA00BDCE0B")]
	public interface IHTMLPainterOverlay : ICOMObject
	{
		#region Methods

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// </summary>
		/// <param name="rcDevice">tagRECT rcDevice</param>
		[SupportByVersion("MSHTML", 4)]
		Int32 onmove(tagRECT rcDevice);

		#endregion
	}
}
