using System;
using NetRuntimeSystem = System;
using System.ComponentModel;
using NetOffice.Attributes;

namespace NetOffice.OWC10Api
{
	/// <summary>
	/// DispatchInterface FieldListHierarchy 
	/// SupportByVersion OWC10, 1
	/// </summary>
	[SupportByVersion("OWC10", 1)]
	[EntityType(EntityType.IsDispatchInterface)]
	[TypeId("7BD180A4-0406-11D3-8549-00C04FAC67D7")]
	public interface FieldListHierarchy : ICOMObject
	{
		#region Properties

		/// <summary>
		/// SupportByVersion OWC10 1
		/// Get
		/// </summary>
		[SupportByVersion("OWC10", 1)]
		NetOffice.OWC10Api.FieldListNode Root { get; }

		/// <summary>
		/// SupportByVersion OWC10 1
		/// Get/Set
		/// </summary>
		[SupportByVersion("OWC10", 1)]
		bool Visible { get; set; }

		/// <summary>
		/// SupportByVersion OWC10 1
		/// Get
		/// </summary>
		[SupportByVersion("OWC10", 1)]
		NetOffice.OWC10Api.FieldListNode Selection { get; }

		/// <summary>
		/// SupportByVersion OWC10 1
		/// Get/Set
		/// </summary>
		[SupportByVersion("OWC10", 1)]
		bool ConcatenateData { get; set; }

		/// <summary>
		/// SupportByVersion OWC10 1
		/// Get/Set
		/// </summary>
		[SupportByVersion("OWC10", 1)]
		string DataSeparator { get; set; }

		#endregion

		#region Methods

		/// <summary>
		/// SupportByVersion OWC10 1
		/// </summary>
		/// <param name="pflhs">NetOffice.OWC10Api.FieldListHierarchySite pflhs</param>
		[SupportByVersion("OWC10", 1)]
		void SetHierarchySite(NetOffice.OWC10Api.FieldListHierarchySite pflhs);

		/// <summary>
		/// SupportByVersion OWC10 1
		/// </summary>
		/// <param name="pflnParent">NetOffice.OWC10Api.FieldListNode pflnParent</param>
		/// <param name="fInsertFirst">bool fInsertFirst</param>
		/// <param name="nID">Int32 nID</param>
		/// <param name="bstrName">string bstrName</param>
		/// <param name="bstrData">string bstrData</param>
		/// <param name="nType">Int32 nType</param>
		[SupportByVersion("OWC10", 1)]
		NetOffice.OWC10Api.FieldListNode AddNode(NetOffice.OWC10Api.FieldListNode pflnParent, bool fInsertFirst, Int32 nID, string bstrName, string bstrData, Int32 nType);

		/// <summary>
		/// SupportByVersion OWC10 1
		/// </summary>
		/// <param name="nID">Int32 nID</param>
		[SupportByVersion("OWC10", 1)]
		NetOffice.OWC10Api.FieldListNode GetNode(Int32 nID);

		/// <summary>
		/// SupportByVersion OWC10 1
		/// </summary>
		/// <param name="pfln">NetOffice.OWC10Api.FieldListNode pfln</param>
		[SupportByVersion("OWC10", 1)]
		void RemoveNode(NetOffice.OWC10Api.FieldListNode pfln);

		/// <summary>
		/// SupportByVersion OWC10 1
		/// </summary>
		/// <param name="nType">Int32 nType</param>
		[SupportByVersion("OWC10", 1)]
		NetOffice.OWC10Api.FieldListType AddType(Int32 nType);

		/// <summary>
		/// SupportByVersion OWC10 1
		/// </summary>
		/// <param name="nTypeId">Int32 nTypeId</param>
		[SupportByVersion("OWC10", 1)]
		NetOffice.OWC10Api.FieldListType GetType(Int32 nTypeId);

		/// <summary>
		/// SupportByVersion OWC10 1
		/// </summary>
		/// <param name="pfln">NetOffice.OWC10Api.FieldListNode pfln</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		[SupportByVersion("OWC10", 1)]
		NetOffice.OWC10Api.FieldListNode GetNextSelected(NetOffice.OWC10Api.FieldListNode pfln);

		#endregion
	}
}
