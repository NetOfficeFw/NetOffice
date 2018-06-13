using System;
using NetRuntimeSystem = System;
using System.ComponentModel;
using NetOffice.Attributes;

namespace NetOffice.OutlookApi
{
	/// <summary>
	/// DispatchInterface _FormRegionStartup 
	/// SupportByVersion Outlook, 12,14,15,16
	/// </summary>
	[SupportByVersion("Outlook", 12,14,15,16)]
	[EntityType(EntityType.IsDispatchInterface), BaseType]
	[TypeId("00063059-0000-0000-C000-000000000046")]
    [CoClassSource(typeof(NetOffice.OutlookApi.FormRegionStartup))]
    public interface _FormRegionStartup : ICOMObject
	{
		#region Methods

		/// <summary>
		/// SupportByVersion Outlook 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff866042.aspx </remarks>
		/// <param name="formRegionName">string formRegionName</param>
		/// <param name="item">object item</param>
		/// <param name="lCID">Int32 lCID</param>
		/// <param name="formRegionMode">NetOffice.OutlookApi.Enums.OlFormRegionMode formRegionMode</param>
		/// <param name="formRegionSize">NetOffice.OutlookApi.Enums.OlFormRegionSize formRegionSize</param>
		[SupportByVersion("Outlook", 12,14,15,16)]
		object GetFormRegionStorage(string formRegionName, object item, Int32 lCID, NetOffice.OutlookApi.Enums.OlFormRegionMode formRegionMode, NetOffice.OutlookApi.Enums.OlFormRegionSize formRegionSize);

		/// <summary>
		/// SupportByVersion Outlook 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff869072.aspx </remarks>
		/// <param name="formRegion">NetOffice.OutlookApi.FormRegion formRegion</param>
		[SupportByVersion("Outlook", 12,14,15,16)]
		void BeforeFormRegionShow(NetOffice.OutlookApi.FormRegion formRegion);

		/// <summary>
		/// SupportByVersion Outlook 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff869502.aspx </remarks>
		/// <param name="formRegionName">string formRegionName</param>
		/// <param name="lCID">Int32 lCID</param>
		[SupportByVersion("Outlook", 12,14,15,16)]
		object GetFormRegionManifest(string formRegionName, Int32 lCID);

		/// <summary>
		/// SupportByVersion Outlook 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff868914.aspx </remarks>
		/// <param name="formRegionName">string formRegionName</param>
		/// <param name="lCID">Int32 lCID</param>
		/// <param name="icon">NetOffice.OutlookApi.Enums.OlFormRegionIcon icon</param>
		[SupportByVersion("Outlook", 12,14,15,16)]
		object GetFormRegionIcon(string formRegionName, Int32 lCID, NetOffice.OutlookApi.Enums.OlFormRegionIcon icon);

		#endregion
	}
}
