using System;
using NetRuntimeSystem = System;
using System.ComponentModel;
using NetOffice.Attributes;

namespace NetOffice.MSHTMLApi
{
	/// <summary>
	/// Interface IXMLGenericParse 
	/// SupportByVersion MSHTML, 4
	/// </summary>
	[SupportByVersion("MSHTML", 4)]
	[EntityType(EntityType.IsInterface)]
	[TypeId("E4E23071-4D07-11D2-AE76-0080C73BC199")]
	public interface IXMLGenericParse : ICOMObject
	{
		#region Methods

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// </summary>
		/// <param name="fDoGeneric">bool fDoGeneric</param>
		[SupportByVersion("MSHTML", 4)]
		Int32 SetGenericParse(bool fDoGeneric);

		#endregion
	}
}
