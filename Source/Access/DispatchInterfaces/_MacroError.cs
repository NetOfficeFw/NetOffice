using System;
using NetRuntimeSystem = System;
using System.ComponentModel;
using NetOffice.Attributes;

namespace NetOffice.AccessApi
{
	/// <summary>
	/// DispatchInterface _MacroError 
	/// SupportByVersion Access, 12,14,15,16
	/// </summary>
	[SupportByVersion("Access", 12,14,15,16)]
	[EntityType(EntityType.IsDispatchInterface), BaseType]
	[TypeId("22585BA1-7BD1-40AF-8095-E688176CDEB0")]
    [CoClassSource(typeof(NetOffice.AccessApi.MacroError))]
    public interface _MacroError : ICOMObject
	{
		#region Properties

		/// <summary>
		/// SupportByVersion Access 12, 14, 15, 16
		/// Get
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff192907.aspx </remarks>
		[SupportByVersion("Access", 12,14,15,16)]
		string Condition { get; }

		/// <summary>
		/// SupportByVersion Access 12, 14, 15, 16
		/// Get
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff845761.aspx </remarks>
		[SupportByVersion("Access", 12,14,15,16)]
		string ActionName { get; }

		/// <summary>
		/// SupportByVersion Access 12, 14, 15, 16
		/// Get
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff845160.aspx </remarks>
		[SupportByVersion("Access", 12,14,15,16)]
		string Arguments { get; }

		/// <summary>
		/// SupportByVersion Access 12, 14, 15, 16
		/// Get
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff193827.aspx </remarks>
		[SupportByVersion("Access", 12,14,15,16)]
		string Description { get; }

		/// <summary>
		/// SupportByVersion Access 12, 14, 15, 16
		/// Get
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff836073.aspx </remarks>
		[SupportByVersion("Access", 12,14,15,16)]
		Int32 Number { get; }

		/// <summary>
		/// SupportByVersion Access 12, 14, 15, 16
		/// Get
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff198364.aspx </remarks>
		[SupportByVersion("Access", 12,14,15,16)]
		string MacroName { get; }

		#endregion

		#region Methods

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
