using System;
using NetRuntimeSystem = System;
using System.ComponentModel;
using NetOffice.Attributes;

namespace NetOffice.OutlookApi
{
	/// <summary>
	/// DispatchInterface _CategoryRuleCondition 
	/// SupportByVersion Outlook, 12,14,15,16
	/// </summary>
	[SupportByVersion("Outlook", 12,14,15,16)]
	[EntityType(EntityType.IsDispatchInterface), BaseType]
	[TypeId("000630DC-0000-0000-C000-000000000046")]
    [CoClassSource(typeof(NetOffice.OutlookApi.CategoryRuleCondition))]
    public interface _CategoryRuleCondition : ICOMObject
	{
		#region Properties

		/// <summary>
		/// SupportByVersion Outlook 12, 14, 15, 16
		/// Get
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff864750.aspx </remarks>
		[SupportByVersion("Outlook", 12,14,15,16)]
		[BaseResult]
		NetOffice.OutlookApi._Application Application { get; }

		/// <summary>
		/// SupportByVersion Outlook 12, 14, 15, 16
		/// Get
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff864455.aspx </remarks>
		[SupportByVersion("Outlook", 12,14,15,16)]
		NetOffice.OutlookApi.Enums.OlObjectClass Class { get; }

		/// <summary>
		/// SupportByVersion Outlook 12, 14, 15, 16
		/// Get
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff869796.aspx </remarks>
		[SupportByVersion("Outlook", 12,14,15,16)]
		[BaseResult]
		NetOffice.OutlookApi._NameSpace Session { get; }

		/// <summary>
		/// SupportByVersion Outlook 12, 14, 15, 16
		/// Get
		/// Unknown COM Proxy
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff868817.aspx </remarks>
		[SupportByVersion("Outlook", 12,14,15,16), ProxyResult]
		object Parent { get; }

		/// <summary>
		/// SupportByVersion Outlook 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff860359.aspx </remarks>
		[SupportByVersion("Outlook", 12,14,15,16)]
		bool Enabled { get; set; }

		/// <summary>
		/// SupportByVersion Outlook 12, 14, 15, 16
		/// Get
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff861565.aspx </remarks>
		[SupportByVersion("Outlook", 12,14,15,16)]
		NetOffice.OutlookApi.Enums.OlRuleConditionType ConditionType { get; }

		/// <summary>
		/// SupportByVersion Outlook 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff866733.aspx </remarks>
		[SupportByVersion("Outlook", 12,14,15,16)]
		object Categories { get; set; }

		#endregion

	}
}
