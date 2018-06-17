using System;
using NetRuntimeSystem = System;
using System.ComponentModel;
using NetOffice.Attributes;

namespace NetOffice.AccessApi
{
	/// <summary>
	/// DispatchInterface _DependencyInfo 
	/// SupportByVersion Access, 11,12,14,15,16
	/// </summary>
	[SupportByVersion("Access", 11,12,14,15,16)]
	[EntityType(EntityType.IsDispatchInterface), BaseType]
	[TypeId("D05819C6-8859-418B-A82F-18B6CB743C8E")]
    [CoClassSource(typeof(NetOffice.AccessApi.DependencyInfo))]
    public interface _DependencyInfo : ICOMObject
	{
		#region Properties

		/// <summary>
		/// SupportByVersion Access 11, 12, 14, 15, 16
		/// Get
		/// Unknown COM Proxy
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff821492.aspx </remarks>
		[SupportByVersion("Access", 11,12,14,15,16), ProxyResult]
		object Parent { get; }

		/// <summary>
		/// SupportByVersion Access 11, 12, 14, 15, 16
		/// Get
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff836006.aspx </remarks>
		[SupportByVersion("Access", 11,12,14,15,16)]
		[BaseResult]
		NetOffice.AccessApi._DependencyObjects Dependants { get; }

		/// <summary>
		/// SupportByVersion Access 11, 12, 14, 15, 16
		/// Get
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff192892.aspx </remarks>
		[SupportByVersion("Access", 11,12,14,15,16)]
		[BaseResult]
		NetOffice.AccessApi._DependencyObjects Dependencies { get; }

		/// <summary>
		/// SupportByVersion Access 11, 12, 14, 15, 16
		/// Get
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff192732.aspx </remarks>
		[SupportByVersion("Access", 11,12,14,15,16)]
		[BaseResult]
		NetOffice.AccessApi._DependencyObjects OutOfDateObjects { get; }

		/// <summary>
		/// SupportByVersion Access 11, 12, 14, 15, 16
		/// Get
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff822804.aspx </remarks>
		[SupportByVersion("Access", 11,12,14,15,16)]
		[BaseResult]
		NetOffice.AccessApi._DependencyObjects InsufficientPermissions { get; }

		/// <summary>
		/// SupportByVersion Access 11, 12, 14, 15, 16
		/// Get
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff197991.aspx </remarks>
		[SupportByVersion("Access", 11,12,14,15,16)]
		[BaseResult]
		NetOffice.AccessApi._DependencyObjects UnsupportedObjects { get; }

		#endregion

		#region Methods

		/// <summary>
		/// SupportByVersion Access 11, 12, 14, 15, 16
		/// </summary>
		/// <param name="dispid">Int32 dispid</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		[SupportByVersion("Access", 11,12,14,15,16)]
		bool IsMemberSafe(Int32 dispid);

		#endregion
	}
}
