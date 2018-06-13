using System;
using NetRuntimeSystem = System;
using System.ComponentModel;
using NetOffice.Attributes;

namespace NetOffice.OutlookApi
{
	/// <summary>
	/// Interface _IInspectorCtrl 
	/// SupportByVersion Outlook, 10
	/// </summary>
	[SupportByVersion("Outlook", 10)]
	[EntityType(EntityType.IsInterface)]
	[TypeId("E182A127-EADD-46E1-B878-482C48CD8754")]
	public interface _IInspectorCtrl : ICOMObject
	{
		#region Properties

		/// <summary>
		/// SupportByVersion Outlook 10
		/// Get/Set
		/// </summary>
		[SupportByVersion("Outlook", 10)]
		string URL { get; set; }

		/// <summary>
		/// SupportByVersion Outlook 10
		/// Get
		/// Unknown COM Proxy
		/// </summary>
		[SupportByVersion("Outlook", 10), ProxyResult]
		object Item { get; }

		#endregion

		#region Methods

		/// <summary>
		/// SupportByVersion Outlook 10
		/// </summary>
		/// <param name="pdispItem">object pdispItem</param>
		[SupportByVersion("Outlook", 10)]
		Int32 OnItemChange(object pdispItem);

		#endregion
	}
}
