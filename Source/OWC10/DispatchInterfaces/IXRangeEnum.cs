using System.Reflection;
using System;
using NetRuntimeSystem = System;
using System.ComponentModel;
using NetOffice.Attributes;

namespace NetOffice.OWC10Api
{
	/// <summary>
	/// DispatchInterface IXRangeEnum 
	/// SupportByVersion OWC10, 1
	/// </summary>
	[SupportByVersion("OWC10", 1)]
	[EntityType(EntityType.IsDispatchInterface)]
	[TypeId("F5B39B09-1480-11D3-8549-00C04FAC67D7")]
	public interface IXRangeEnum : ICOMObject
	{
		#region Properties

		/// <summary>
		/// SupportByVersion OWC10 1
		/// Get
		/// </summary>
		[SupportByVersion("OWC10", 1)]
		UIntPtr RowCount { get; }

		/// <summary>
		/// SupportByVersion OWC10 1
		/// Get
		/// </summary>
		[SupportByVersion("OWC10", 1)]
		UIntPtr ColCount { get; }

		#endregion

		#region Methods

		/// <summary>
		/// SupportByVersion OWC10 1
		/// </summary>
		/// <param name="cElt">Int32 cElt</param>
		/// <param name="rgvar">object rgvar</param>
		/// <param name="pcEltFetched">Int32 pcEltFetched</param>
		[SupportByVersion("OWC10", 1)]
		void Next(Int32 cElt, out object rgvar, out Int32 pcEltFetched);

		/// <summary>
		/// SupportByVersion OWC10 1
		/// </summary>
		/// <param name="cElt">Int32 cElt</param>
		[SupportByVersion("OWC10", 1)]
		void Skip(Int32 cElt);

		/// <summary>
		/// SupportByVersion OWC10 1
		/// </summary>
		[SupportByVersion("OWC10", 1)]
		void Reset();

		/// <summary>
		/// SupportByVersion OWC10 1
		/// </summary>
		/// <param name="ppEnum">NetOffice.OWC10Api.IXRangeEnum ppEnum</param>
		[SupportByVersion("OWC10", 1)]
		void Clone(out NetOffice.OWC10Api.IXRangeEnum ppEnum);

		/// <summary>
		/// SupportByVersion OWC10 1
		/// </summary>
		/// <param name="nTraverseCode">UIntPtr nTraverseCode</param>
		[SupportByVersion("OWC10", 1)]
		void SetTraversal(UIntPtr nTraverseCode);

		/// <summary>
		/// SupportByVersion OWC10 1
		/// </summary>
		/// <param name="_out">object out</param>
		/// <param name="_in">object in</param>
		/// <param name="vt">Int16 vt</param>
		[SupportByVersion("OWC10", 1)]
		void ChangeType(out object _out, object _in, Int16 vt);

		/// <summary>
		/// SupportByVersion OWC10 1
		/// </summary>
		/// <param name="cElt">Int32 cElt</param>
		/// <param name="iStart">Int32 iStart</param>
		/// <param name="rvarDest">object rvarDest</param>
		/// <param name="pcFetched">Int32 pcFetched</param>
		/// <param name="vtCoerceTo">Int16 vtCoerceTo</param>
		/// <param name="vtbCoerceFrom">Int32 vtbCoerceFrom</param>
		/// <param name="fill">object fill</param>
		[SupportByVersion("OWC10", 1)]
		void GetElements(Int32 cElt, Int32 iStart, object rvarDest, out Int32 pcFetched, Int16 vtCoerceTo, Int32 vtbCoerceFrom, object fill);

		#endregion
	}
}
