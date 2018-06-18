using System.Reflection;
using System;
using NetRuntimeSystem = System;
using System.ComponentModel;
using NetOffice.Attributes;

namespace NetOffice.MSHTMLApi
{
	/// <summary>
	/// Interface ISecureUrlHost 
	/// SupportByVersion MSHTML, 4
	/// </summary>
	[SupportByVersion("MSHTML", 4)]
	[EntityType(EntityType.IsInterface)]
	[TypeId("C81984C4-74C8-11D2-BAA9-00C04FC2040E")]
	public interface ISecureUrlHost : ICOMObject
	{
		#region Methods

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// </summary>
		/// <param name="pfAllow">Int32 pfAllow</param>
		/// <param name="pchUrlInQuestion">Int16 pchUrlInQuestion</param>
		/// <param name="dwFlags">Int32 dwFlags</param>
		[SupportByVersion("MSHTML", 4)]
		Int32 ValidateSecureUrl(out Int32 pfAllow, Int16 pchUrlInQuestion, Int32 dwFlags);

		#endregion
	}
}
