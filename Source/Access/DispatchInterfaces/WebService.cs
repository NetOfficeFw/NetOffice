using System;
using NetRuntimeSystem = System;
using System.ComponentModel;
using NetOffice.Attributes;

namespace NetOffice.AccessApi
{
	/// <summary>
	/// DispatchInterface WebService 
	/// SupportByVersion Access, 14,15,16
	/// </summary>
	/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff834759.aspx </remarks>
	[SupportByVersion("Access", 14,15,16)]
	[EntityType(EntityType.IsDispatchInterface)]
	[TypeId("7A7B2B33-A93D-4E04-BFC7-6041AAA0046D")]
	public interface WebService : ICOMObject
	{
		#region Properties

		/// <summary>
		/// SupportByVersion Access 14, 15, 16
		/// Get
		/// Unknown COM Proxy
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff823018.aspx </remarks>
		[SupportByVersion("Access", 14,15,16), ProxyResult]
		object Parent { get; }

		/// <summary>
		/// SupportByVersion Access 14, 15, 16
		/// Get
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff836020.aspx </remarks>
		[SupportByVersion("Access", 14,15,16)]
		string Name { get; }

		/// <summary>
		/// SupportByVersion Access 14, 15, 16
		/// Get
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff821721.aspx </remarks>
		[SupportByVersion("Access", 14,15,16)]
		NetOffice.AccessApi.Entities Entities { get; }

		#endregion

		#region Methods

		/// <summary>
		/// SupportByVersion Access 14, 15, 16
		/// </summary>
		/// <param name="dispid">Int32 dispid</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		[SupportByVersion("Access", 14,15,16)]
		bool IsMemberSafe(Int32 dispid);

		#endregion
	}
}
