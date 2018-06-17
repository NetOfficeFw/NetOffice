using System;
using NetRuntimeSystem = System;
using System.ComponentModel;
using NetOffice.Attributes;

namespace NetOffice.AccessApi
{
	/// <summary>
	/// DispatchInterface _SmartTag 
	/// SupportByVersion Access, 11,12,14,15,16
	/// </summary>
	[SupportByVersion("Access", 11,12,14,15,16)]
	[EntityType(EntityType.IsDispatchInterface), BaseType]
	[TypeId("9D2AB5D3-CD72-4A9A-A72E-2B3492CBD0AE")]
    [CoClassSource(typeof(NetOffice.AccessApi.SmartTag))]
	public interface _SmartTag : ICOMObject
	{
		#region Properties

		/// <summary>
		/// SupportByVersion Access 11, 12, 14, 15, 16
		/// Get
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff821717.aspx </remarks>
		[SupportByVersion("Access", 11,12,14,15,16)]
		NetOffice.AccessApi.Application Application { get; }

		/// <summary>
		/// SupportByVersion Access 11, 12, 14, 15, 16
		/// Get
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff192048.aspx </remarks>
		[SupportByVersion("Access", 11,12,14,15,16)]
		string Name { get; }

		/// <summary>
		/// SupportByVersion Access 11, 12, 14, 15, 16
		/// Get
		/// Unknown COM Proxy
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff194759.aspx </remarks>
		[SupportByVersion("Access", 11,12,14,15,16), ProxyResult]
		object Parent { get; }

		/// <summary>
		/// SupportByVersion Access 11, 12, 14, 15, 16
		/// Get
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff834783.aspx </remarks>
		[SupportByVersion("Access", 11,12,14,15,16)]
		[BaseResult]
		NetOffice.AccessApi._SmartTagProperties Properties { get; }

		/// <summary>
		/// SupportByVersion Access 11, 12, 14, 15, 16
		/// Get
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff845414.aspx </remarks>
		[SupportByVersion("Access", 11,12,14,15,16)]
		[BaseResult]
		NetOffice.AccessApi._SmartTagActions SmartTagActions { get; }

		/// <summary>
		/// SupportByVersion Access 11, 12, 14, 15, 16
		/// Get
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff197988.aspx </remarks>
		[SupportByVersion("Access", 11,12,14,15,16)]
		string XML { get; }

		/// <summary>
		/// SupportByVersion Access 11, 12, 14, 15, 16
		/// Get
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff844953.aspx </remarks>
		[SupportByVersion("Access", 11,12,14,15,16)]
		bool IsMissing { get; }

		#endregion

		#region Methods

		/// <summary>
		/// SupportByVersion Access 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff192119.aspx </remarks>
		[SupportByVersion("Access", 11,12,14,15,16)]
		void Delete();

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
