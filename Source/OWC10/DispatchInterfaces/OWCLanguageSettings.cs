using System;
using NetRuntimeSystem = System;
using System.ComponentModel;
using NetOffice.Attributes;

namespace NetOffice.OWC10Api
{
	/// <summary>
	/// DispatchInterface OWCLanguageSettings 
	/// SupportByVersion OWC10, 1
	/// </summary>
	[SupportByVersion("OWC10", 1)]
	[EntityType(EntityType.IsDispatchInterface)]
	[TypeId("6F5A76C3-0AC7-4DED-9A6B-A3547FD7B7BB")]
	public interface OWCLanguageSettings : ICOMObject
	{
		#region Properties

		/// <summary>
		/// SupportByVersion OWC10 1
		/// Get
		/// Unknown COM Proxy
		/// </summary>
		[SupportByVersion("OWC10", 1), ProxyResult]
		object Application { get; }

		/// <summary>
		/// SupportByVersion OWC10 1
		/// Get
		/// </summary>
		/// <param name="id">NetOffice.OWC10Api.Enums.MsoAppLanguageID id</param>
		[SupportByVersion("OWC10", 1)]
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		Int32 get_LanguageID(NetOffice.OWC10Api.Enums.MsoAppLanguageID id);

		/// <summary>
		/// SupportByVersion OWC10 1
		/// Alias for get_LanguageID
		/// </summary>
		/// <param name="id">NetOffice.OWC10Api.Enums.MsoAppLanguageID id</param>
		[SupportByVersion("OWC10", 1), Redirect("get_LanguageID")]
		Int32 LanguageID(NetOffice.OWC10Api.Enums.MsoAppLanguageID id);

		/// <summary>
		/// SupportByVersion OWC10 1
		/// Get
		/// </summary>
		/// <param name="lid">NetOffice.OWC10Api.Enums.MsoLanguageID lid</param>
		[SupportByVersion("OWC10", 1)]
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		bool get_LanguagePreferredForEditing(NetOffice.OWC10Api.Enums.MsoLanguageID lid);

		/// <summary>
		/// SupportByVersion OWC10 1
		/// Alias for get_LanguagePreferredForEditing
		/// </summary>
		/// <param name="lid">NetOffice.OWC10Api.Enums.MsoLanguageID lid</param>
		[SupportByVersion("OWC10", 1), Redirect("get_LanguagePreferredForEditing")]
		bool LanguagePreferredForEditing(NetOffice.OWC10Api.Enums.MsoLanguageID lid);

		#endregion

	}
}
