using System;
using NetRuntimeSystem = System;
using System.ComponentModel;
using NetOffice.Attributes;

namespace NetOffice.VisioApi
{
	/// <summary>
	/// DispatchInterface IVBUndoUnit 
	/// SupportByVersion Visio, 11,12,14,15,16
	/// </summary>
	/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/ff769307(v=office.14).aspx </remarks>
	[SupportByVersion("Visio", 11,12,14,15,16)]
	[EntityType(EntityType.IsDispatchInterface)]
	[TypeId("000D1305-0000-0000-C000-000000000046")]
	public interface IVBUndoUnit : ICOMObject
	{
		#region Properties

		/// <summary>
		/// SupportByVersion Visio 11, 12, 14, 15, 16
		/// Get
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/ff765404(v=office.14).aspx </remarks>
		[SupportByVersion("Visio", 11,12,14,15,16)]
		string Description { get; }

		/// <summary>
		/// SupportByVersion Visio 11, 12, 14, 15, 16
		/// Get
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/ff767058(v=office.14).aspx </remarks>
		[SupportByVersion("Visio", 11,12,14,15,16)]
		string UnitTypeCLSID { get; }

		/// <summary>
		/// SupportByVersion Visio 11, 12, 14, 15, 16
		/// Get
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/ff766307(v=office.14).aspx </remarks>
		[SupportByVersion("Visio", 11,12,14,15,16)]
		Int32 UnitTypeLong { get; }

		/// <summary>
		/// SupportByVersion Visio 11, 12, 14, 15, 16
		/// Get
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/ff766293(v=office.14).aspx </remarks>
		[SupportByVersion("Visio", 11,12,14,15,16)]
		Int32 UnitSize { get; }

		#endregion

		#region Methods

		/// <summary>
		/// SupportByVersion Visio 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/ff766032(v=office.14).aspx </remarks>
		/// <param name="pMgr">NetOffice.VisioApi.IVBUndoManager pMgr</param>
		[SupportByVersion("Visio", 11,12,14,15,16)]
		void Do(NetOffice.VisioApi.IVBUndoManager pMgr);

		/// <summary>
		/// SupportByVersion Visio 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/ff767691(v=office.14).aspx </remarks>
		[SupportByVersion("Visio", 11,12,14,15,16)]
		void OnNextAdd();

		#endregion
	}
}
