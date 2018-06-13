using System;
using NetRuntimeSystem = System;
using System.ComponentModel;
using NetOffice.Attributes;

namespace NetOffice.OutlookApi
{
	/// <summary>
	/// DispatchInterface _Rule 
	/// SupportByVersion Outlook, 12,14,15,16
	/// </summary>
	[SupportByVersion("Outlook", 12,14,15,16)]
	[EntityType(EntityType.IsDispatchInterface), BaseType]
	[TypeId("000630CD-0000-0000-C000-000000000046")]
    [CoClassSource(typeof(NetOffice.OutlookApi.Rule))]
    public interface _Rule : ICOMObject
	{
		#region Properties

		/// <summary>
		/// SupportByVersion Outlook 12, 14, 15, 16
		/// Get
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff867534.aspx </remarks>
		[SupportByVersion("Outlook", 12,14,15,16)]
		[BaseResult]
		NetOffice.OutlookApi._Application Application { get; }

		/// <summary>
		/// SupportByVersion Outlook 12, 14, 15, 16
		/// Get
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff868179.aspx </remarks>
		[SupportByVersion("Outlook", 12,14,15,16)]
		NetOffice.OutlookApi.Enums.OlObjectClass Class { get; }

		/// <summary>
		/// SupportByVersion Outlook 12, 14, 15, 16
		/// Get
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff866703.aspx </remarks>
		[SupportByVersion("Outlook", 12,14,15,16)]
		[BaseResult]
		NetOffice.OutlookApi._NameSpace Session { get; }

		/// <summary>
		/// SupportByVersion Outlook 12, 14, 15, 16
		/// Get
		/// Unknown COM Proxy
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff869390.aspx </remarks>
		[SupportByVersion("Outlook", 12,14,15,16), ProxyResult]
		object Parent { get; }

		/// <summary>
		/// SupportByVersion Outlook 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff866230.aspx </remarks>
		[SupportByVersion("Outlook", 12,14,15,16)]
		string Name { get; set; }

		/// <summary>
		/// SupportByVersion Outlook 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff860647.aspx </remarks>
		[SupportByVersion("Outlook", 12,14,15,16)]
		Int32 ExecutionOrder { get; set; }

		/// <summary>
		/// SupportByVersion Outlook 12, 14, 15, 16
		/// Get
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff866197.aspx </remarks>
		[SupportByVersion("Outlook", 12,14,15,16)]
		NetOffice.OutlookApi.Enums.OlRuleType RuleType { get; }

		/// <summary>
		/// SupportByVersion Outlook 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff868097.aspx </remarks>
		[SupportByVersion("Outlook", 12,14,15,16)]
		bool Enabled { get; set; }

		/// <summary>
		/// SupportByVersion Outlook 12, 14, 15, 16
		/// Get
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff864224.aspx </remarks>
		[SupportByVersion("Outlook", 12,14,15,16)]
		bool IsLocalRule { get; }

		/// <summary>
		/// SupportByVersion Outlook 12, 14, 15, 16
		/// Get
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff862765.aspx </remarks>
		[SupportByVersion("Outlook", 12,14,15,16)]
		NetOffice.OutlookApi.RuleActions Actions { get; }

		/// <summary>
		/// SupportByVersion Outlook 12, 14, 15, 16
		/// Get
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff869579.aspx </remarks>
		[SupportByVersion("Outlook", 12,14,15,16)]
		NetOffice.OutlookApi.RuleConditions Conditions { get; }

		/// <summary>
		/// SupportByVersion Outlook 12, 14, 15, 16
		/// Get
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff867239.aspx </remarks>
		[SupportByVersion("Outlook", 12,14,15,16)]
		NetOffice.OutlookApi.RuleConditions Exceptions { get; }

		#endregion

		#region Methods

		/// <summary>
		/// SupportByVersion Outlook 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff864433.aspx </remarks>
		/// <param name="showProgress">optional object showProgress</param>
		/// <param name="folder">optional object folder</param>
		/// <param name="includeSubfolders">optional object includeSubfolders</param>
		/// <param name="ruleExecuteOption">optional object ruleExecuteOption</param>
		[SupportByVersion("Outlook", 12,14,15,16)]
		void Execute(object showProgress, object folder, object includeSubfolders, object ruleExecuteOption);

		/// <summary>
		/// SupportByVersion Outlook 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff864433.aspx </remarks>
		[CustomMethod]
		[SupportByVersion("Outlook", 12,14,15,16)]
		void Execute();

		/// <summary>
		/// SupportByVersion Outlook 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff864433.aspx </remarks>
		/// <param name="showProgress">optional object showProgress</param>
		[CustomMethod]
		[SupportByVersion("Outlook", 12,14,15,16)]
		void Execute(object showProgress);

		/// <summary>
		/// SupportByVersion Outlook 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff864433.aspx </remarks>
		/// <param name="showProgress">optional object showProgress</param>
		/// <param name="folder">optional object folder</param>
		[CustomMethod]
		[SupportByVersion("Outlook", 12,14,15,16)]
		void Execute(object showProgress, object folder);

		/// <summary>
		/// SupportByVersion Outlook 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff864433.aspx </remarks>
		/// <param name="showProgress">optional object showProgress</param>
		/// <param name="folder">optional object folder</param>
		/// <param name="includeSubfolders">optional object includeSubfolders</param>
		[CustomMethod]
		[SupportByVersion("Outlook", 12,14,15,16)]
		void Execute(object showProgress, object folder, object includeSubfolders);

		#endregion
	}
}
