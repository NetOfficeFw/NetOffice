using System;
using NetRuntimeSystem = System;
using System.ComponentModel;
using NetOffice.Attributes;

namespace NetOffice.OutlookApi
{
	/// <summary>
	/// DispatchInterface _PropertyAccessor 
	/// SupportByVersion Outlook, 12,14,15,16
	/// </summary>
	[SupportByVersion("Outlook", 12,14,15,16)]
	[EntityType(EntityType.IsDispatchInterface), BaseType]
	[TypeId("0006302D-0000-0000-C000-000000000046")]
    [CoClassSource(typeof(NetOffice.OutlookApi.PropertyAccessor))]
    public interface _PropertyAccessor : ICOMObject
	{
		#region Properties

		/// <summary>
		/// SupportByVersion Outlook 12, 14, 15, 16
		/// Get
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff865030.aspx </remarks>
		[SupportByVersion("Outlook", 12,14,15,16)]
		[BaseResult]
		NetOffice.OutlookApi._Application Application { get; }

		/// <summary>
		/// SupportByVersion Outlook 12, 14, 15, 16
		/// Get
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff869821.aspx </remarks>
		[SupportByVersion("Outlook", 12,14,15,16)]
		NetOffice.OutlookApi.Enums.OlObjectClass Class { get; }

		/// <summary>
		/// SupportByVersion Outlook 12, 14, 15, 16
		/// Get
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff869435.aspx </remarks>
		[SupportByVersion("Outlook", 12,14,15,16)]
		[BaseResult]
		NetOffice.OutlookApi._NameSpace Session { get; }

		/// <summary>
		/// SupportByVersion Outlook 12, 14, 15, 16
		/// Get
		/// Unknown COM Proxy
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff866730.aspx </remarks>
		[SupportByVersion("Outlook", 12,14,15,16), ProxyResult]
		object Parent { get; }

		#endregion

		#region Methods

		/// <summary>
		/// SupportByVersion Outlook 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff868350.aspx </remarks>
		/// <param name="schemaName">string schemaName</param>
		[SupportByVersion("Outlook", 12,14,15,16)]
		object GetProperty(string schemaName);

		/// <summary>
		/// SupportByVersion Outlook 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff862751.aspx </remarks>
		/// <param name="schemaName">string schemaName</param>
		/// <param name="value">object value</param>
		[SupportByVersion("Outlook", 12,14,15,16)]
		void SetProperty(string schemaName, object value);

		/// <summary>
		/// SupportByVersion Outlook 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff869865.aspx </remarks>
		/// <param name="schemaNames">object schemaNames</param>
		[SupportByVersion("Outlook", 12,14,15,16)]
		object GetProperties(object schemaNames);

		/// <summary>
		/// SupportByVersion Outlook 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff868862.aspx </remarks>
		/// <param name="schemaNames">object schemaNames</param>
		/// <param name="values">object values</param>
		[SupportByVersion("Outlook", 12,14,15,16)]
		object SetProperties(object schemaNames, object values);

		/// <summary>
		/// SupportByVersion Outlook 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff868342.aspx </remarks>
		/// <param name="value">DateTime value</param>
		[SupportByVersion("Outlook", 12,14,15,16)]
		DateTime UTCToLocalTime(DateTime value);

		/// <summary>
		/// SupportByVersion Outlook 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff868909.aspx </remarks>
		/// <param name="value">DateTime value</param>
		[SupportByVersion("Outlook", 12,14,15,16)]
		DateTime LocalTimeToUTC(DateTime value);

		/// <summary>
		/// SupportByVersion Outlook 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff862123.aspx </remarks>
		/// <param name="value">string value</param>
		[SupportByVersion("Outlook", 12,14,15,16)]
		object StringToBinary(string value);

		/// <summary>
		/// SupportByVersion Outlook 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff864468.aspx </remarks>
		/// <param name="value">object value</param>
		[SupportByVersion("Outlook", 12,14,15,16)]
		string BinaryToString(object value);

		/// <summary>
		/// SupportByVersion Outlook 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff868076.aspx </remarks>
		/// <param name="schemaName">string schemaName</param>
		[SupportByVersion("Outlook", 12,14,15,16)]
		void DeleteProperty(string schemaName);

		/// <summary>
		/// SupportByVersion Outlook 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff869707.aspx </remarks>
		/// <param name="schemaNames">object schemaNames</param>
		[SupportByVersion("Outlook", 12,14,15,16)]
		object DeleteProperties(object schemaNames);

		#endregion
	}
}
