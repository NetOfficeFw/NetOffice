using System;
using NetRuntimeSystem = System;
using System.ComponentModel;
using NetOffice.Attributes;

namespace NetOffice.VisioApi
{
	/// <summary>
	/// Interface LPVISIOCONTAINERPROPERTIES 
	/// SupportByVersion Visio, 14,15,16
	/// </summary>
	[SupportByVersion("Visio", 14,15,16)]
	[EntityType(EntityType.IsInterface)]
	[TypeId("00000000-0000-0000-0000-000000000000")]
	public interface LPVISIOCONTAINERPROPERTIES : ICOMObject
	{
		#region Properties

		/// <summary>
		/// SupportByVersion Visio 14, 15, 16
		/// Get
		/// </summary>
		[SupportByVersion("Visio", 14,15,16)]
		[BaseResult]
		NetOffice.VisioApi.IVApplication Application { get; }

		/// <summary>
		/// SupportByVersion Visio 14, 15, 16
		/// Get
		/// </summary>
		[SupportByVersion("Visio", 14,15,16)]
		Int16 Stat { get; }

		/// <summary>
		/// SupportByVersion Visio 14, 15, 16
		/// Get
		/// </summary>
		[SupportByVersion("Visio", 14,15,16)]
		Int16 ObjectType { get; }

		/// <summary>
		/// SupportByVersion Visio 14, 15, 16
		/// Get
		/// </summary>
		[SupportByVersion("Visio", 14,15,16)]
		[BaseResult]
		NetOffice.VisioApi.IVDocument Document { get; }

		/// <summary>
		/// SupportByVersion Visio 14, 15, 16
		/// Get
		/// </summary>
		[SupportByVersion("Visio", 14,15,16)]
		[BaseResult]
		NetOffice.VisioApi.IVShape Shape { get; }

		/// <summary>
		/// SupportByVersion Visio 14, 15, 16
		/// Get
		/// </summary>
		[SupportByVersion("Visio", 14,15,16)]
		NetOffice.VisioApi.Enums.VisContainerTypes ContainerType { get; }

		/// <summary>
		/// SupportByVersion Visio 14, 15, 16
		/// Get/Set
		/// </summary>
		[SupportByVersion("Visio", 14,15,16)]
		NetOffice.VisioApi.Enums.VisListAlignment ListAlignment { get; set; }

		/// <summary>
		/// SupportByVersion Visio 14, 15, 16
		/// Get/Set
		/// </summary>
		[SupportByVersion("Visio", 14,15,16)]
		NetOffice.VisioApi.Enums.VisListDirection ListDirection { get; set; }

		/// <summary>
		/// SupportByVersion Visio 14, 15, 16
		/// Get/Set
		/// </summary>
		[SupportByVersion("Visio", 14,15,16)]
		bool LockMembership { get; set; }

		/// <summary>
		/// SupportByVersion Visio 14, 15, 16
		/// Get/Set
		/// </summary>
		[SupportByVersion("Visio", 14,15,16)]
		NetOffice.VisioApi.Enums.VisContainerAutoResize ResizeAsNeeded { get; set; }

		/// <summary>
		/// SupportByVersion Visio 14, 15, 16
		/// Get/Set
		/// </summary>
		[SupportByVersion("Visio", 14,15,16)]
		[BaseResult]
		NetOffice.VisioApi.IVShape OverlappedList { get; set; }

		/// <summary>
		/// SupportByVersion Visio 14, 15, 16
		/// Get/Set
		/// </summary>
		[SupportByVersion("Visio", 14,15,16)]
		Int32 ContainerStyle { get; set; }

		/// <summary>
		/// SupportByVersion Visio 14, 15, 16
		/// Get/Set
		/// </summary>
		[SupportByVersion("Visio", 14,15,16)]
		Int32 HeadingStyle { get; set; }

		#endregion

		#region Methods

		/// <summary>
		/// SupportByVersion Visio 14, 15, 16
		/// </summary>
		[SupportByVersion("Visio", 14,15,16)]
		void Disband();

		/// <summary>
		/// SupportByVersion Visio 14, 15, 16
		/// </summary>
		[SupportByVersion("Visio", 14,15,16)]
		void FitToContents();

		/// <summary>
		/// SupportByVersion Visio 14, 15, 16
		/// </summary>
		/// <param name="marginUnits">NetOffice.VisioApi.Enums.VisUnitCodes marginUnits</param>
		[SupportByVersion("Visio", 14,15,16)]
		Double GetMargin(NetOffice.VisioApi.Enums.VisUnitCodes marginUnits);

		/// <summary>
		/// SupportByVersion Visio 14, 15, 16
		/// </summary>
		/// <param name="marginUnits">NetOffice.VisioApi.Enums.VisUnitCodes marginUnits</param>
		/// <param name="marginSize">Double marginSize</param>
		[SupportByVersion("Visio", 14,15,16)]
		void SetMargin(NetOffice.VisioApi.Enums.VisUnitCodes marginUnits, Double marginSize);

		/// <summary>
		/// SupportByVersion Visio 14, 15, 16
		/// </summary>
		/// <param name="spacingUnits">NetOffice.VisioApi.Enums.VisUnitCodes spacingUnits</param>
		[SupportByVersion("Visio", 14,15,16)]
		Double GetListSpacing(NetOffice.VisioApi.Enums.VisUnitCodes spacingUnits);

		/// <summary>
		/// SupportByVersion Visio 14, 15, 16
		/// </summary>
		/// <param name="spacingUnits">NetOffice.VisioApi.Enums.VisUnitCodes spacingUnits</param>
		/// <param name="spacingSize">Double spacingSize</param>
		[SupportByVersion("Visio", 14,15,16)]
		void SetListSpacing(NetOffice.VisioApi.Enums.VisUnitCodes spacingUnits, Double spacingSize);

		/// <summary>
		/// SupportByVersion Visio 14, 15, 16
		/// </summary>
		/// <param name="objectToInsert">object objectToInsert</param>
		/// <param name="position">Int32 position</param>
		[SupportByVersion("Visio", 14,15,16)]
		void InsertListMember(object objectToInsert, Int32 position);

		/// <summary>
		/// SupportByVersion Visio 14, 15, 16
		/// </summary>
		/// <param name="shapeMember">NetOffice.VisioApi.IVShape shapeMember</param>
		[SupportByVersion("Visio", 14,15,16)]
		Int32 GetListMemberPosition(NetOffice.VisioApi.IVShape shapeMember);

		/// <summary>
		/// SupportByVersion Visio 14, 15, 16
		/// </summary>
		/// <param name="shape">NetOffice.VisioApi.IVShape shape</param>
		[SupportByVersion("Visio", 14,15,16)]
		NetOffice.VisioApi.Enums.VisContainerMemberState GetMemberState(NetOffice.VisioApi.IVShape shape);

		/// <summary>
		/// SupportByVersion Visio 14, 15, 16
		/// </summary>
		/// <param name="objectToRemove">object objectToRemove</param>
		[SupportByVersion("Visio", 14,15,16)]
		void RemoveMember(object objectToRemove);

		/// <summary>
		/// SupportByVersion Visio 14, 15, 16
		/// </summary>
		/// <param name="objectToReorder">object objectToReorder</param>
		/// <param name="position">Int32 position</param>
		[SupportByVersion("Visio", 14,15,16)]
		void ReorderListMember(object objectToReorder, Int32 position);

		/// <summary>
		/// SupportByVersion Visio 14, 15, 16
		/// </summary>
		[SupportByVersion("Visio", 14,15,16)]
		Int32[] GetListMembers();

		/// <summary>
		/// SupportByVersion Visio 14, 15, 16
		/// </summary>
		/// <param name="containerFlags">Int32 containerFlags</param>
		[SupportByVersion("Visio", 14,15,16)]
		Int32[] GetMemberShapes(Int32 containerFlags);

		/// <summary>
		/// SupportByVersion Visio 14, 15, 16
		/// </summary>
		/// <param name="pObjectToAdd">object pObjectToAdd</param>
		/// <param name="addOptions">NetOffice.VisioApi.Enums.VisMemberAddOptions addOptions</param>
		[SupportByVersion("Visio", 14,15,16)]
		void AddMember(object pObjectToAdd, NetOffice.VisioApi.Enums.VisMemberAddOptions addOptions);

		/// <summary>
		/// SupportByVersion Visio 14, 15, 16
		/// </summary>
		/// <param name="direction">NetOffice.VisioApi.Enums.VisLayoutDirection direction</param>
		[SupportByVersion("Visio", 14,15,16)]
		void RotateFlipList(NetOffice.VisioApi.Enums.VisLayoutDirection direction);

		#endregion
	}
}
