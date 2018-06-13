using System;
using NetRuntimeSystem = System;
using System.ComponentModel;
using NetOffice.Attributes;

namespace NetOffice.OutlookApi
{
	/// <summary>
	/// DispatchInterface _SolutionsModule 
	/// SupportByVersion Outlook, 14,15,16
	/// </summary>
	[SupportByVersion("Outlook", 14,15,16)]
	[EntityType(EntityType.IsDispatchInterface), BaseType]
	[TypeId("000630FF-0000-0000-C000-000000000046")]
    [CoClassSource(typeof(NetOffice.OutlookApi.SolutionsModule))]
    public interface _SolutionsModule : ICOMObject
	{
		#region Properties

		/// <summary>
		/// SupportByVersion Outlook 14, 15, 16
		/// Get
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff867375.aspx </remarks>
		[SupportByVersion("Outlook", 14,15,16)]
		[BaseResult]
		NetOffice.OutlookApi._Application Application { get; }

		/// <summary>
		/// SupportByVersion Outlook 14, 15, 16
		/// Get
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff869389.aspx </remarks>
		[SupportByVersion("Outlook", 14,15,16)]
		NetOffice.OutlookApi.Enums.OlObjectClass Class { get; }

		/// <summary>
		/// SupportByVersion Outlook 14, 15, 16
		/// Get
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff862689.aspx </remarks>
		[SupportByVersion("Outlook", 14,15,16)]
		[BaseResult]
		NetOffice.OutlookApi._NameSpace Session { get; }

		/// <summary>
		/// SupportByVersion Outlook 14, 15, 16
		/// Get
		/// Unknown COM Proxy
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff860715.aspx </remarks>
		[SupportByVersion("Outlook", 14,15,16), ProxyResult]
		object Parent { get; }

		/// <summary>
		/// SupportByVersion Outlook 14, 15, 16
		/// Get
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff866442.aspx </remarks>
		[SupportByVersion("Outlook", 14,15,16)]
		NetOffice.OutlookApi.Enums.OlNavigationModuleType NavigationModuleType { get; }

		/// <summary>
		/// SupportByVersion Outlook 14, 15, 16
		/// Get
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff869461.aspx </remarks>
		[SupportByVersion("Outlook", 14,15,16)]
		string Name { get; }

		/// <summary>
		/// SupportByVersion Outlook 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff869583.aspx </remarks>
		[SupportByVersion("Outlook", 14,15,16)]
		Int32 Position { get; set; }

		/// <summary>
		/// SupportByVersion Outlook 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff864776.aspx </remarks>
		[SupportByVersion("Outlook", 14,15,16)]
		bool Visible { get; set; }

		#endregion

		#region Methods

		/// <summary>
		/// SupportByVersion Outlook 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff867156.aspx </remarks>
		/// <param name="solution">NetOffice.OutlookApi.MAPIFolder solution</param>
		/// <param name="scope">NetOffice.OutlookApi.Enums.OlSolutionScope scope</param>
		[SupportByVersion("Outlook", 14,15,16)]
		void AddSolution(NetOffice.OutlookApi.MAPIFolder solution, NetOffice.OutlookApi.Enums.OlSolutionScope scope);

		#endregion
	}
}
