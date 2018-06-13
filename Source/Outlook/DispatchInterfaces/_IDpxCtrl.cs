using System;
using NetRuntimeSystem = System;
using System.ComponentModel;
using NetOffice.Attributes;

namespace NetOffice.OutlookApi
{
	/// <summary>
	/// DispatchInterface _IDpxCtrl 
	/// SupportByVersion Outlook, 10
	/// </summary>
	[SupportByVersion("Outlook", 10)]
	[EntityType(EntityType.IsDispatchInterface), BaseType]
	[TypeId("0006F097-0000-0000-C000-000000000046")]
    [CoClassSource(typeof(NetOffice.OutlookApi._DpxCtrl))]
    public interface _IDpxCtrl : ICOMObject
	{
		#region Properties

		/// <summary>
		/// SupportByVersion Outlook 10
		/// Get/Set
		/// </summary>
		[SupportByVersion("Outlook", 10)]
		Int32 StartDate { get; set; }

		/// <summary>
		/// SupportByVersion Outlook 10
		/// Get/Set
		/// </summary>
		[SupportByVersion("Outlook", 10)]
		Int32 EndDate { get; set; }

		#endregion

	}
}
