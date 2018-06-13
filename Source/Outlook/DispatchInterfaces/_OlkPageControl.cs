using System;
using NetRuntimeSystem = System;
using System.ComponentModel;
using NetOffice.Attributes;

namespace NetOffice.OutlookApi
{
	/// <summary>
	/// DispatchInterface _OlkPageControl 
	/// SupportByVersion Outlook, 12,14,15,16
	/// </summary>
	[SupportByVersion("Outlook", 12,14,15,16)]
	[EntityType(EntityType.IsDispatchInterface), BaseType]
	[TypeId("000672F8-0000-0000-C000-000000000046")]
    [CoClassSource(typeof(NetOffice.OutlookApi.OlkPageControl))]
    public interface _OlkPageControl : ICOMObject
	{
		#region Properties

		/// <summary>
		/// SupportByVersion Outlook 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff870128.aspx </remarks>
		[SupportByVersion("Outlook", 12,14,15,16)]
		NetOffice.OutlookApi.Enums.OlPageType Page { get; set; }

		#endregion

	}
}
