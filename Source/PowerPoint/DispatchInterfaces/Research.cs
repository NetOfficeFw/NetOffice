using System;
using NetRuntimeSystem = System;
using System.ComponentModel;
using NetOffice.Attributes;

namespace NetOffice.PowerPointApi
{
	/// <summary>
	/// DispatchInterface Research 
	/// SupportByVersion PowerPoint, 12,14,15,16
	/// </summary>
	/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff745646.aspx </remarks>
	[SupportByVersion("PowerPoint", 12,14,15,16)]
	[EntityType(EntityType.IsDispatchInterface)]
	[TypeId("914934F7-5A91-11CF-8700-00AA0060263B")]
	public interface Research : ICOMObject
	{
		#region Properties

		/// <summary>
		/// SupportByVersion PowerPoint 12, 14, 15, 16
		/// Get
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff746098.aspx </remarks>
		[SupportByVersion("PowerPoint", 12,14,15,16)]
		NetOffice.PowerPointApi.Application Application { get; }

		/// <summary>
		/// SupportByVersion PowerPoint 12, 14, 15, 16
		/// Get
		/// Unknown COM Proxy
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff744070.aspx </remarks>
		[SupportByVersion("PowerPoint", 12,14,15,16), ProxyResult]
		object Parent { get; }

		#endregion

		#region Methods

		/// <summary>
		/// SupportByVersion PowerPoint 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff744220.aspx </remarks>
		/// <param name="serviceID">string serviceID</param>
		/// <param name="queryString">optional object queryString</param>
		/// <param name="queryLanguage">optional object queryLanguage</param>
		/// <param name="useSelection">optional bool UseSelection = false</param>
		/// <param name="launchQuery">optional bool LaunchQuery = true</param>
		[SupportByVersion("PowerPoint", 12,14,15,16)]
		void Query(string serviceID, object queryString, object queryLanguage, object useSelection, object launchQuery);

		/// <summary>
		/// SupportByVersion PowerPoint 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff744220.aspx </remarks>
		/// <param name="serviceID">string serviceID</param>
		[CustomMethod]
		[SupportByVersion("PowerPoint", 12,14,15,16)]
		void Query(string serviceID);

		/// <summary>
		/// SupportByVersion PowerPoint 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff744220.aspx </remarks>
		/// <param name="serviceID">string serviceID</param>
		/// <param name="queryString">optional object queryString</param>
		[CustomMethod]
		[SupportByVersion("PowerPoint", 12,14,15,16)]
		void Query(string serviceID, object queryString);

		/// <summary>
		/// SupportByVersion PowerPoint 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff744220.aspx </remarks>
		/// <param name="serviceID">string serviceID</param>
		/// <param name="queryString">optional object queryString</param>
		/// <param name="queryLanguage">optional object queryLanguage</param>
		[CustomMethod]
		[SupportByVersion("PowerPoint", 12,14,15,16)]
		void Query(string serviceID, object queryString, object queryLanguage);

		/// <summary>
		/// SupportByVersion PowerPoint 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff744220.aspx </remarks>
		/// <param name="serviceID">string serviceID</param>
		/// <param name="queryString">optional object queryString</param>
		/// <param name="queryLanguage">optional object queryLanguage</param>
		/// <param name="useSelection">optional bool UseSelection = false</param>
		[CustomMethod]
		[SupportByVersion("PowerPoint", 12,14,15,16)]
		void Query(string serviceID, object queryString, object queryLanguage, object useSelection);

		/// <summary>
		/// SupportByVersion PowerPoint 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff745349.aspx </remarks>
		/// <param name="language1">object language1</param>
		/// <param name="language2">object language2</param>
		[SupportByVersion("PowerPoint", 12,14,15,16)]
		void SetLanguagePair(object language1, object language2);

		/// <summary>
		/// SupportByVersion PowerPoint 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff746351.aspx </remarks>
		/// <param name="serviceID">string serviceID</param>
		[SupportByVersion("PowerPoint", 12,14,15,16)]
		bool IsResearchService(string serviceID);

		#endregion
	}
}
