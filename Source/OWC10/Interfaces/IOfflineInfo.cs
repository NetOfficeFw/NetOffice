using System;
using NetRuntimeSystem = System;
using System.ComponentModel;
using NetOffice.Attributes;

namespace NetOffice.OWC10Api
{
	/// <summary>
	/// Interface IOfflineInfo 
	/// SupportByVersion OWC10, 1
	/// </summary>
	[SupportByVersion("OWC10", 1)]
	[EntityType(EntityType.IsInterface), BaseType]
	[TypeId("E2AC0C69-7079-11D3-8D01-0050048383A8")]
    [CoClassSource(typeof(NetOffice.OWC10Api.OfflineInfo))]
    public interface IOfflineInfo : ICOMObject
	{
		#region Methods

		/// <summary>
		/// SupportByVersion OWC10 1
		/// </summary>
		/// <param name="pwzUrl">string pwzUrl</param>
		/// <param name="pwzServerFilter">string pwzServerFilter</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		[SupportByVersion("OWC10", 1)]
		Int32 PutServerFilter(string pwzUrl, string pwzServerFilter);

		/// <summary>
		/// SupportByVersion OWC10 1
		/// </summary>
		/// <param name="pwzUrl">string pwzUrl</param>
		/// <param name="pwzServerFilter">string pwzServerFilter</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		[SupportByVersion("OWC10", 1)]
		Int32 GetServerFilter(string pwzUrl, string pwzServerFilter);

		/// <summary>
		/// SupportByVersion OWC10 1
		/// </summary>
		/// <param name="pwzUrl">string pwzUrl</param>
		/// <param name="pfSubscribed">Int32 pfSubscribed</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		[SupportByVersion("OWC10", 1)]
		Int32 GetIsPageSubscribed(string pwzUrl, Int32 pfSubscribed);

		/// <summary>
		/// SupportByVersion OWC10 1
		/// </summary>
		/// <param name="pbstrPath">string pbstrPath</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		[SupportByVersion("OWC10", 1)]
		Int32 GetOfflineXMLFileLocation(string pbstrPath);

		#endregion
	}
}
