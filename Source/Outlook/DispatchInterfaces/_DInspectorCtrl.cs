using System;
using NetRuntimeSystem = System;
using System.ComponentModel;
using NetOffice.Attributes;

namespace NetOffice.OutlookApi
{
	/// <summary>
	/// DispatchInterface _DInspectorCtrl 
	/// SupportByVersion Outlook, 10
	/// </summary>
	[SupportByVersion("Outlook", 10)]
	[EntityType(EntityType.IsDispatchInterface), BaseType]
	[TypeId("0006F09D-0000-0000-C000-000000000046")]
    [CoClassSource(typeof(NetOffice.OutlookApi._InspectorCtrl))]
    public interface _DInspectorCtrl : ICOMObject
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
		void OnItemChange(object pdispItem);

		#endregion
	}
}
