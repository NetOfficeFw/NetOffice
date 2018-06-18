using System;
using NetRuntimeSystem = System;
using System.ComponentModel;
using NetOffice.Attributes;

namespace NetOffice.VisioApi
{
	/// <summary>
	/// DispatchInterface IVisEventProc 
	/// SupportByVersion Visio, 11,12,14,15,16
	/// </summary>
	/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/ff769310(v=office.14).aspx </remarks>
	[SupportByVersion("Visio", 11,12,14,15,16)]
	[EntityType(EntityType.IsDispatchInterface)]
	[TypeId("000D0728-0000-0000-C000-000000000046")]
	public interface IVisEventProc : ICOMObject
	{
		#region Methods

		/// <summary>
		/// SupportByVersion Visio 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/ff768483(v=office.14).aspx </remarks>
		/// <param name="nEventCode">Int16 nEventCode</param>
		/// <param name="pSourceObj">object pSourceObj</param>
		/// <param name="nEventID">Int32 nEventID</param>
		/// <param name="nEventSeqNum">Int32 nEventSeqNum</param>
		/// <param name="pSubjectObj">object pSubjectObj</param>
		/// <param name="vMoreInfo">object vMoreInfo</param>
		[SupportByVersion("Visio", 11,12,14,15,16)]
		object VisEventProc(Int16 nEventCode, object pSourceObj, Int32 nEventID, Int32 nEventSeqNum, object pSubjectObj, object vMoreInfo);

		#endregion
	}
}
