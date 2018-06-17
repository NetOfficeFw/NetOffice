using System;
using NetRuntimeSystem = System;
using System.ComponentModel;
using NetOffice.Attributes;

namespace NetOffice.AccessApi
{
	/// <summary>
	/// DispatchInterface ImportExportSpecification 
	/// SupportByVersion Access, 12,14,15,16
	/// </summary>
	/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff820988.aspx </remarks>
	[SupportByVersion("Access", 12,14,15,16)]
	[EntityType(EntityType.IsDispatchInterface)]
	[TypeId("55B0E0C9-C75D-4F42-AD20-6939C1D05B70")]
	public interface ImportExportSpecification : ICOMObject
	{
		#region Properties

		/// <summary>
		/// SupportByVersion Access 12, 14, 15, 16
		/// Get
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff195433.aspx </remarks>
		[SupportByVersion("Access", 12,14,15,16)]
		NetOffice.AccessApi.Application Application { get; }

		/// <summary>
		/// SupportByVersion Access 12, 14, 15, 16
		/// Get
		/// Unknown COM Proxy
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff834782.aspx </remarks>
		[SupportByVersion("Access", 12,14,15,16), ProxyResult]
		object Parent { get; }

		/// <summary>
		/// SupportByVersion Access 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff192471.aspx </remarks>
		[SupportByVersion("Access", 12,14,15,16)]
		string Name { get; set; }

		/// <summary>
		/// SupportByVersion Access 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff197630.aspx </remarks>
		[SupportByVersion("Access", 12,14,15,16)]
		string XML { get; set; }

		/// <summary>
		/// SupportByVersion Access 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff837049.aspx </remarks>
		[SupportByVersion("Access", 12,14,15,16)]
		string Description { get; set; }

		#endregion

		#region Methods

		/// <summary>
		/// SupportByVersion Access 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff837213.aspx </remarks>
		/// <param name="prompt">optional object prompt</param>
		[SupportByVersion("Access", 12,14,15,16)]
		void Execute(object prompt);

		/// <summary>
		/// SupportByVersion Access 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff837213.aspx </remarks>
		[CustomMethod]
		[SupportByVersion("Access", 12,14,15,16)]
		void Execute();

		/// <summary>
		/// SupportByVersion Access 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff834427.aspx </remarks>
		[SupportByVersion("Access", 12,14,15,16)]
		void Delete();

		/// <summary>
		/// SupportByVersion Access 12, 14, 15, 16
		/// </summary>
		/// <param name="dispid">Int32 dispid</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		[SupportByVersion("Access", 12,14,15,16)]
		bool IsMemberSafe(Int32 dispid);

		#endregion
	}
}
