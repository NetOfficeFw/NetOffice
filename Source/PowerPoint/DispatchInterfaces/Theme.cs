using System;
using NetRuntimeSystem = System;
using System.ComponentModel;
using NetOffice.Attributes;

namespace NetOffice.PowerPointApi
{
	/// <summary>
	/// DispatchInterface Theme 
	/// SupportByVersion PowerPoint, 15, 16
	/// </summary>
	/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/jj230426.aspx </remarks>
	[SupportByVersion("PowerPoint", 15, 16)]
	[EntityType(EntityType.IsDispatchInterface)]
	[TypeId("D9D60EB3-D4B4-4991-9C16-75585B3346BB")]
	public interface Theme : ICOMObject
	{
		#region Properties

		/// <summary>
		/// SupportByVersion PowerPoint 15,16
		/// Get
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/jj717753.aspx </remarks>
		[SupportByVersion("PowerPoint", 15, 16)]
		NetOffice.PowerPointApi.Application Application { get; }

		/// <summary>
		/// SupportByVersion PowerPoint 15,16
		/// Get
		/// Unknown COM Proxy
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/jj717713.aspx </remarks>
		[SupportByVersion("PowerPoint", 15, 16), ProxyResult]
		object Parent { get; }

		/// <summary>
		/// SupportByVersion PowerPoint 15,16
		/// Get
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/jj230439.aspx </remarks>
		[SupportByVersion("PowerPoint", 15, 16)]
		NetOffice.PowerPointApi.ThemeVariants ThemeVariants { get; }

		#endregion

	}
}
