using System.Reflection;
using System;
using NetRuntimeSystem = System;
using System.ComponentModel;
using NetOffice.Attributes;

namespace NetOffice.MSHTMLApi
{
	/// <summary>
	/// Interface IEnumPrivacyRecords 
	/// SupportByVersion MSHTML, 4
	/// </summary>
	[SupportByVersion("MSHTML", 4)]
	[EntityType(EntityType.IsInterface)]
	[TypeId("3050F844-98B5-11CF-BB82-00AA00BDCE0B")]
	public interface IEnumPrivacyRecords : ICOMObject
	{
		#region Methods

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// </summary>
		[SupportByVersion("MSHTML", 4)]
		Int32 reset();

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// </summary>
		/// <param name="pSize">Int32 pSize</param>
		[SupportByVersion("MSHTML", 4)]
		Int32 GetSize(out Int32 pSize);

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// </summary>
		/// <param name="pState">Int32 pState</param>
		[SupportByVersion("MSHTML", 4)]
		Int32 GetPrivacyImpacted(out Int32 pState);

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// </summary>
		/// <param name="pbstrUrl">string pbstrUrl</param>
		/// <param name="pbstrPolicyRef">string pbstrPolicyRef</param>
		/// <param name="pdwReserved">Int32 pdwReserved</param>
		/// <param name="pdwPrivacyFlags">Int32 pdwPrivacyFlags</param>
		[SupportByVersion("MSHTML", 4)]
		Int32 Next(out string pbstrUrl, out string pbstrPolicyRef, out Int32 pdwReserved, out Int32 pdwPrivacyFlags);

		#endregion
	}
}
