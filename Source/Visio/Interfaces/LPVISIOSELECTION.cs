using System.Collections;
using System.Collections.Generic;
using System.Reflection;
using System;
using NetRuntimeSystem = System;
using System.ComponentModel;
using NetOffice.Attributes;
using NetOffice.CollectionsGeneric;

namespace NetOffice.VisioApi
{
	/// <summary>
	/// Interface LPVISIOSELECTION 
	/// SupportByVersion Visio, 11,12,14,15,16
	/// </summary>
	[SupportByVersion("Visio", 11,12,14,15,16)]
	[EntityType(EntityType.IsInterface), Enumerator(Enumerator.Reference, EnumeratorInvoke.Property, "Visio", 11, 12, 14, 15, 16), HasIndexProperty(IndexInvoke.Property, "Item")]
	[TypeId("00000000-0000-0000-0000-000000000000")]
	public interface LPVISIOSELECTION : ICOMObject, IEnumerableProvider<NetOffice.VisioApi.IVShape>
	{
		#region Properties

		/// <summary>
		/// SupportByVersion Visio 11, 12, 14, 15, 16
		/// Get
		/// </summary>
		[SupportByVersion("Visio", 11,12,14,15,16)]
		[BaseResult]
		NetOffice.VisioApi.IVApplication Application { get; }

		/// <summary>
		/// SupportByVersion Visio 11, 12, 14, 15, 16
		/// Get
		/// </summary>
		[SupportByVersion("Visio", 11,12,14,15,16)]
		Int16 Stat { get; }

		/// <summary>
		/// SupportByVersion Visio 11, 12, 14, 15, 16
		/// Get
		/// </summary>
		[SupportByVersion("Visio", 11,12,14,15,16)]
		Int16 ObjectType { get; }

		/// <summary>
		/// SupportByVersion Visio 11, 12, 14, 15, 16
		/// Get
		/// </summary>
		/// <param name="index">Int16 index</param>
		[SupportByVersion("Visio", 11,12,14,15,16)]
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		NetOffice.VisioApi.IVShape get_Item16(Int16 index);

		/// <summary>
		/// SupportByVersion Visio 11, 12, 14, 15, 16
		/// Alias for get_Item16
		/// </summary>
		/// <param name="index">Int16 index</param>
		[SupportByVersion("Visio", 11,12,14,15,16), Redirect("get_Item16")]
		NetOffice.VisioApi.IVShape Item16(Int16 index);

		/// <summary>
		/// SupportByVersion Visio 11, 12, 14, 15, 16
		/// Get
		/// </summary>
		[SupportByVersion("Visio", 11,12,14,15,16)]
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		Int16 Count16 { get; }

		/// <summary>
		/// SupportByVersion Visio 11, 12, 14, 15, 16
		/// Get
		/// </summary>
		[SupportByVersion("Visio", 11,12,14,15,16)]
		[BaseResult]
		NetOffice.VisioApi.IVDocument Document { get; }

		/// <summary>
		/// SupportByVersion Visio 11, 12, 14, 15, 16
		/// Get
		/// </summary>
		[SupportByVersion("Visio", 11,12,14,15,16)]
		[BaseResult]
		NetOffice.VisioApi.IVPage ContainingPage { get; }

		/// <summary>
		/// SupportByVersion Visio 11, 12, 14, 15, 16
		/// Get
		/// </summary>
		[SupportByVersion("Visio", 11,12,14,15,16)]
		[BaseResult]
		NetOffice.VisioApi.IVMaster ContainingMaster { get; }

		/// <summary>
		/// SupportByVersion Visio 11, 12, 14, 15, 16
		/// Get
		/// </summary>
		[SupportByVersion("Visio", 11,12,14,15,16)]
		[BaseResult]
		NetOffice.VisioApi.IVShape ContainingShape { get; }

		/// <summary>
		/// SupportByVersion Visio 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		[SupportByVersion("Visio", 11,12,14,15,16)]
		string Style { get; set; }

		/// <summary>
		/// SupportByVersion Visio 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		[SupportByVersion("Visio", 11,12,14,15,16)]
		string StyleKeepFmt { get; set; }

		/// <summary>
		/// SupportByVersion Visio 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		[SupportByVersion("Visio", 11,12,14,15,16)]
		string LineStyle { get; set; }

		/// <summary>
		/// SupportByVersion Visio 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		[SupportByVersion("Visio", 11,12,14,15,16)]
		string LineStyleKeepFmt { get; set; }

		/// <summary>
		/// SupportByVersion Visio 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		[SupportByVersion("Visio", 11,12,14,15,16)]
		string FillStyle { get; set; }

		/// <summary>
		/// SupportByVersion Visio 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		[SupportByVersion("Visio", 11,12,14,15,16)]
		string FillStyleKeepFmt { get; set; }

		/// <summary>
		/// SupportByVersion Visio 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		[SupportByVersion("Visio", 11,12,14,15,16)]
		string TextStyle { get; set; }

		/// <summary>
		/// SupportByVersion Visio 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		[SupportByVersion("Visio", 11,12,14,15,16)]
		string TextStyleKeepFmt { get; set; }

		/// <summary>
		/// SupportByVersion Visio 11, 12, 14, 15, 16
		/// Get
		/// </summary>
		[SupportByVersion("Visio", 11,12,14,15,16)]
		[BaseResult]
		NetOffice.VisioApi.IVEventList EventList { get; }

		/// <summary>
		/// SupportByVersion Visio 11, 12, 14, 15, 16
		/// Get
		/// </summary>
		[SupportByVersion("Visio", 11,12,14,15,16)]
		Int16 PersistsEvents { get; }

		/// <summary>
		/// SupportByVersion Visio 11, 12, 14, 15, 16
		/// Get
		/// </summary>
		/// <param name="index">Int32 index</param>
		[SupportByVersion("Visio", 11,12,14,15,16)]
		[BaseResult]
		[NetRuntimeSystem.Runtime.CompilerServices.IndexerName("Item"), IndexProperty]
		NetOffice.VisioApi.IVShape this[Int32 index] { get; }

		/// <summary>
		/// SupportByVersion Visio 11, 12, 14, 15, 16
		/// Get
		/// </summary>
		[SupportByVersion("Visio", 11,12,14,15,16)]
		Int32 Count { get; }

		/// <summary>
		/// SupportByVersion Visio 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		[SupportByVersion("Visio", 11,12,14,15,16)]
		Int32 IterationMode { get; set; }

		/// <summary>
		/// SupportByVersion Visio 11, 12, 14, 15, 16
		/// Get
		/// </summary>
		/// <param name="index">Int32 index</param>
		[SupportByVersion("Visio", 11,12,14,15,16)]
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		Int16 get_ItemStatus(Int32 index);

		/// <summary>
		/// SupportByVersion Visio 11, 12, 14, 15, 16
		/// Alias for get_ItemStatus
		/// </summary>
		/// <param name="index">Int32 index</param>
		[SupportByVersion("Visio", 11,12,14,15,16), Redirect("get_ItemStatus")]
		Int16 ItemStatus(Int32 index);

		/// <summary>
		/// SupportByVersion Visio 11, 12, 14, 15, 16
		/// Get
		/// </summary>
		[SupportByVersion("Visio", 11,12,14,15,16)]
		[BaseResult]
		NetOffice.VisioApi.IVShape PrimaryItem { get; }

		/// <summary>
		/// SupportByVersion Visio 11, 12, 14, 15, 16
		/// Get
		/// </summary>
		[SupportByVersion("Visio", 11,12,14,15,16), NativeResult]
		stdole.Picture Picture { get; }

		/// <summary>
		/// SupportByVersion Visio 11, 12, 14, 15, 16
		/// Get
		/// </summary>
		[SupportByVersion("Visio", 11,12,14,15,16)]
		Int32 ContainingPageID { get; }

		/// <summary>
		/// SupportByVersion Visio 11, 12, 14, 15, 16
		/// Get
		/// </summary>
		[SupportByVersion("Visio", 11,12,14,15,16)]
		Int32 ContainingMasterID { get; }

		/// <summary>
		/// SupportByVersion Visio 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		[SupportByVersion("Visio", 12,14,15,16)]
		[BaseResult]
		NetOffice.VisioApi.IVMaster DataGraphic { get; set; }

		/// <summary>
		/// SupportByVersion Visio 14, 15, 16
		/// Get
		/// </summary>
		[SupportByVersion("Visio", 14,15,16)]
		[BaseResult]
		NetOffice.VisioApi.IVSelection SelectionForDragCopy { get; }

		#endregion

		#region Methods

		/// <summary>
		/// SupportByVersion Visio 11, 12, 14, 15, 16
		/// </summary>
		/// <param name="fileName">string fileName</param>
		[SupportByVersion("Visio", 11,12,14,15,16)]
		void Export(string fileName);

		/// <summary>
		/// SupportByVersion Visio 11, 12, 14, 15, 16
		/// </summary>
		[SupportByVersion("Visio", 11,12,14,15,16)]
		void BringForward();

		/// <summary>
		/// SupportByVersion Visio 11, 12, 14, 15, 16
		/// </summary>
		[SupportByVersion("Visio", 11,12,14,15,16)]
		void BringToFront();

		/// <summary>
		/// SupportByVersion Visio 11, 12, 14, 15, 16
		/// </summary>
		[SupportByVersion("Visio", 11,12,14,15,16)]
		void SendBackward();

		/// <summary>
		/// SupportByVersion Visio 11, 12, 14, 15, 16
		/// </summary>
		[SupportByVersion("Visio", 11,12,14,15,16)]
		void SendToBack();

		/// <summary>
		/// SupportByVersion Visio 11, 12, 14, 15, 16
		/// </summary>
		[SupportByVersion("Visio", 11,12,14,15,16)]
		void Combine();

		/// <summary>
		/// SupportByVersion Visio 11, 12, 14, 15, 16
		/// </summary>
		[SupportByVersion("Visio", 11,12,14,15,16)]
		void Fragment();

		/// <summary>
		/// SupportByVersion Visio 11, 12, 14, 15, 16
		/// </summary>
		[SupportByVersion("Visio", 11,12,14,15,16)]
		void Intersect();

		/// <summary>
		/// SupportByVersion Visio 11, 12, 14, 15, 16
		/// </summary>
		[SupportByVersion("Visio", 11,12,14,15,16)]
		void Subtract();

		/// <summary>
		/// SupportByVersion Visio 11, 12, 14, 15, 16
		/// </summary>
		[SupportByVersion("Visio", 11,12,14,15,16)]
		void Union();

		/// <summary>
		/// SupportByVersion Visio 11, 12, 14, 15, 16
		/// </summary>
		[SupportByVersion("Visio", 11,12,14,15,16)]
		void FlipHorizontal();

		/// <summary>
		/// SupportByVersion Visio 11, 12, 14, 15, 16
		/// </summary>
		[SupportByVersion("Visio", 11,12,14,15,16)]
		void FlipVertical();

		/// <summary>
		/// SupportByVersion Visio 11, 12, 14, 15, 16
		/// </summary>
		[SupportByVersion("Visio", 11,12,14,15,16)]
		void ReverseEnds();

		/// <summary>
		/// SupportByVersion Visio 11, 12, 14, 15, 16
		/// </summary>
		[SupportByVersion("Visio", 11,12,14,15,16)]
		void Rotate90();

		/// <summary>
		/// SupportByVersion Visio 11, 12, 14, 15, 16
		/// </summary>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		[SupportByVersion("Visio", 11,12,14,15,16)]
		void old_Copy();

		/// <summary>
		/// SupportByVersion Visio 11, 12, 14, 15, 16
		/// </summary>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		[SupportByVersion("Visio", 11,12,14,15,16)]
		void old_Cut();

		/// <summary>
		/// SupportByVersion Visio 11, 12, 14, 15, 16
		/// </summary>
		[SupportByVersion("Visio", 11,12,14,15,16)]
		void Delete();

		/// <summary>
		/// SupportByVersion Visio 11, 12, 14, 15, 16
		/// </summary>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		[SupportByVersion("Visio", 11,12,14,15,16)]
		void VoidDuplicate();

		/// <summary>
		/// SupportByVersion Visio 11, 12, 14, 15, 16
		/// </summary>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		[SupportByVersion("Visio", 11,12,14,15,16)]
		void VoidGroup();

		/// <summary>
		/// SupportByVersion Visio 11, 12, 14, 15, 16
		/// </summary>
		[SupportByVersion("Visio", 11,12,14,15,16)]
		void ConvertToGroup();

		/// <summary>
		/// SupportByVersion Visio 11, 12, 14, 15, 16
		/// </summary>
		[SupportByVersion("Visio", 11,12,14,15,16)]
		void Ungroup();

		/// <summary>
		/// SupportByVersion Visio 11, 12, 14, 15, 16
		/// </summary>
		[SupportByVersion("Visio", 11,12,14,15,16)]
		void SelectAll();

		/// <summary>
		/// SupportByVersion Visio 11, 12, 14, 15, 16
		/// </summary>
		[SupportByVersion("Visio", 11,12,14,15,16)]
		void DeselectAll();

		/// <summary>
		/// SupportByVersion Visio 11, 12, 14, 15, 16
		/// </summary>
		/// <param name="sheetObject">NetOffice.VisioApi.IVShape sheetObject</param>
		/// <param name="selectAction">Int16 selectAction</param>
		[SupportByVersion("Visio", 11,12,14,15,16)]
		void Select(NetOffice.VisioApi.IVShape sheetObject, Int16 selectAction);

		/// <summary>
		/// SupportByVersion Visio 11, 12, 14, 15, 16
		/// </summary>
		[SupportByVersion("Visio", 11,12,14,15,16)]
		void Trim();

		/// <summary>
		/// SupportByVersion Visio 11, 12, 14, 15, 16
		/// </summary>
		[SupportByVersion("Visio", 11,12,14,15,16)]
		void Join();

		/// <summary>
		/// SupportByVersion Visio 11, 12, 14, 15, 16
		/// </summary>
		/// <param name="tolerance">Double tolerance</param>
		/// <param name="flags">Int16 flags</param>
		[SupportByVersion("Visio", 11,12,14,15,16)]
		void FitCurve(Double tolerance, Int16 flags);

		/// <summary>
		/// SupportByVersion Visio 11, 12, 14, 15, 16
		/// </summary>
		[SupportByVersion("Visio", 11,12,14,15,16)]
		void Layout();

		/// <summary>
		/// SupportByVersion Visio 11, 12, 14, 15, 16
		/// </summary>
		/// <param name="flags">Int16 flags</param>
		/// <param name="lpr8Left">Double lpr8Left</param>
		/// <param name="lpr8Bottom">Double lpr8Bottom</param>
		/// <param name="lpr8Right">Double lpr8Right</param>
		/// <param name="lpr8Top">Double lpr8Top</param>
		[SupportByVersion("Visio", 11,12,14,15,16)]
		void BoundingBox(Int16 flags, out Double lpr8Left, out Double lpr8Bottom, out Double lpr8Right, out Double lpr8Top);

		/// <summary>
		/// SupportByVersion Visio 11, 12, 14, 15, 16
		/// </summary>
		/// <param name="tolerance">Double tolerance</param>
		/// <param name="flags">Int16 flags</param>
		/// <param name="x">optional object x</param>
		/// <param name="y">optional object y</param>
		/// <param name="resultsMaster">optional object resultsMaster</param>
		[SupportByVersion("Visio", 11,12,14,15,16)]
		[BaseResult]
		NetOffice.VisioApi.IVShape DrawRegion(Double tolerance, Int16 flags, object x, object y, object resultsMaster);

		/// <summary>
		/// SupportByVersion Visio 11, 12, 14, 15, 16
		/// </summary>
		/// <param name="tolerance">Double tolerance</param>
		/// <param name="flags">Int16 flags</param>
		[CustomMethod]
		[BaseResult]
		[SupportByVersion("Visio", 11,12,14,15,16)]
		NetOffice.VisioApi.IVShape DrawRegion(Double tolerance, Int16 flags);

		/// <summary>
		/// SupportByVersion Visio 11, 12, 14, 15, 16
		/// </summary>
		/// <param name="tolerance">Double tolerance</param>
		/// <param name="flags">Int16 flags</param>
		/// <param name="x">optional object x</param>
		[CustomMethod]
		[BaseResult]
		[SupportByVersion("Visio", 11,12,14,15,16)]
		NetOffice.VisioApi.IVShape DrawRegion(Double tolerance, Int16 flags, object x);

		/// <summary>
		/// SupportByVersion Visio 11, 12, 14, 15, 16
		/// </summary>
		/// <param name="tolerance">Double tolerance</param>
		/// <param name="flags">Int16 flags</param>
		/// <param name="x">optional object x</param>
		/// <param name="y">optional object y</param>
		[CustomMethod]
		[BaseResult]
		[SupportByVersion("Visio", 11,12,14,15,16)]
		NetOffice.VisioApi.IVShape DrawRegion(Double tolerance, Int16 flags, object x, object y);

		/// <summary>
		/// SupportByVersion Visio 11, 12, 14, 15, 16
		/// </summary>
		[SupportByVersion("Visio", 11,12,14,15,16)]
		[BaseResult]
		NetOffice.VisioApi.IVShape Group();

		/// <summary>
		/// SupportByVersion Visio 11, 12, 14, 15, 16
		/// </summary>
		[SupportByVersion("Visio", 11,12,14,15,16)]
		void SwapEnds();

		/// <summary>
		/// SupportByVersion Visio 11, 12, 14, 15, 16
		/// </summary>
		[SupportByVersion("Visio", 11,12,14,15,16)]
		void AddToGroup();

		/// <summary>
		/// SupportByVersion Visio 11, 12, 14, 15, 16
		/// </summary>
		[SupportByVersion("Visio", 11,12,14,15,16)]
		void RemoveFromGroup();

		/// <summary>
		/// SupportByVersion Visio 11, 12, 14, 15, 16
		/// </summary>
		[SupportByVersion("Visio", 11,12,14,15,16)]
		[BaseResult]
		NetOffice.VisioApi.IVSelection Duplicate();

		/// <summary>
		/// SupportByVersion Visio 11, 12, 14, 15, 16
		/// </summary>
		/// <param name="flags">optional object flags</param>
		[SupportByVersion("Visio", 11,12,14,15,16)]
		void Copy(object flags);

		/// <summary>
		/// SupportByVersion Visio 11, 12, 14, 15, 16
		/// </summary>
		[CustomMethod]
		[SupportByVersion("Visio", 11,12,14,15,16)]
		void Copy();

		/// <summary>
		/// SupportByVersion Visio 11, 12, 14, 15, 16
		/// </summary>
		/// <param name="flags">optional object flags</param>
		[SupportByVersion("Visio", 11,12,14,15,16)]
		void Cut(object flags);

		/// <summary>
		/// SupportByVersion Visio 11, 12, 14, 15, 16
		/// </summary>
		[CustomMethod]
		[SupportByVersion("Visio", 11,12,14,15,16)]
		void Cut();

		/// <summary>
		/// SupportByVersion Visio 11, 12, 14, 15, 16
		/// </summary>
		/// <param name="dx">Double dx</param>
		/// <param name="dy">Double dy</param>
		/// <param name="unitsNameOrCode">optional object unitsNameOrCode</param>
		[SupportByVersion("Visio", 11,12,14,15,16)]
		void Move(Double dx, Double dy, object unitsNameOrCode);

		/// <summary>
		/// SupportByVersion Visio 11, 12, 14, 15, 16
		/// </summary>
		/// <param name="dx">Double dx</param>
		/// <param name="dy">Double dy</param>
		[CustomMethod]
		[SupportByVersion("Visio", 11,12,14,15,16)]
		void Move(Double dx, Double dy);

		/// <summary>
		/// SupportByVersion Visio 11, 12, 14, 15, 16
		/// </summary>
		/// <param name="angle">Double angle</param>
		/// <param name="angleUnitsNameOrCode">optional object angleUnitsNameOrCode</param>
		/// <param name="blastGuards">optional bool BlastGuards = false</param>
		/// <param name="rotationType">optional NetOffice.VisioApi.Enums.VisRotationTypes RotationType = 0</param>
		/// <param name="pinX">optional Double PinX = 0</param>
		/// <param name="pinY">optional Double PinY = 0</param>
		/// <param name="pinUnitsNameOrCode">optional object pinUnitsNameOrCode</param>
		[SupportByVersion("Visio", 11,12,14,15,16)]
		void Rotate(Double angle, object angleUnitsNameOrCode, object blastGuards, object rotationType, object pinX, object pinY, object pinUnitsNameOrCode);

		/// <summary>
		/// SupportByVersion Visio 11, 12, 14, 15, 16
		/// </summary>
		/// <param name="angle">Double angle</param>
		[CustomMethod]
		[SupportByVersion("Visio", 11,12,14,15,16)]
		void Rotate(Double angle);

		/// <summary>
		/// SupportByVersion Visio 11, 12, 14, 15, 16
		/// </summary>
		/// <param name="angle">Double angle</param>
		/// <param name="angleUnitsNameOrCode">optional object angleUnitsNameOrCode</param>
		[CustomMethod]
		[SupportByVersion("Visio", 11,12,14,15,16)]
		void Rotate(Double angle, object angleUnitsNameOrCode);

		/// <summary>
		/// SupportByVersion Visio 11, 12, 14, 15, 16
		/// </summary>
		/// <param name="angle">Double angle</param>
		/// <param name="angleUnitsNameOrCode">optional object angleUnitsNameOrCode</param>
		/// <param name="blastGuards">optional bool BlastGuards = false</param>
		[CustomMethod]
		[SupportByVersion("Visio", 11,12,14,15,16)]
		void Rotate(Double angle, object angleUnitsNameOrCode, object blastGuards);

		/// <summary>
		/// SupportByVersion Visio 11, 12, 14, 15, 16
		/// </summary>
		/// <param name="angle">Double angle</param>
		/// <param name="angleUnitsNameOrCode">optional object angleUnitsNameOrCode</param>
		/// <param name="blastGuards">optional bool BlastGuards = false</param>
		/// <param name="rotationType">optional NetOffice.VisioApi.Enums.VisRotationTypes RotationType = 0</param>
		[CustomMethod]
		[SupportByVersion("Visio", 11,12,14,15,16)]
		void Rotate(Double angle, object angleUnitsNameOrCode, object blastGuards, object rotationType);

		/// <summary>
		/// SupportByVersion Visio 11, 12, 14, 15, 16
		/// </summary>
		/// <param name="angle">Double angle</param>
		/// <param name="angleUnitsNameOrCode">optional object angleUnitsNameOrCode</param>
		/// <param name="blastGuards">optional bool BlastGuards = false</param>
		/// <param name="rotationType">optional NetOffice.VisioApi.Enums.VisRotationTypes RotationType = 0</param>
		/// <param name="pinX">optional Double PinX = 0</param>
		[CustomMethod]
		[SupportByVersion("Visio", 11,12,14,15,16)]
		void Rotate(Double angle, object angleUnitsNameOrCode, object blastGuards, object rotationType, object pinX);

		/// <summary>
		/// SupportByVersion Visio 11, 12, 14, 15, 16
		/// </summary>
		/// <param name="angle">Double angle</param>
		/// <param name="angleUnitsNameOrCode">optional object angleUnitsNameOrCode</param>
		/// <param name="blastGuards">optional bool BlastGuards = false</param>
		/// <param name="rotationType">optional NetOffice.VisioApi.Enums.VisRotationTypes RotationType = 0</param>
		/// <param name="pinX">optional Double PinX = 0</param>
		/// <param name="pinY">optional Double PinY = 0</param>
		[CustomMethod]
		[SupportByVersion("Visio", 11,12,14,15,16)]
		void Rotate(Double angle, object angleUnitsNameOrCode, object blastGuards, object rotationType, object pinX, object pinY);

		/// <summary>
		/// SupportByVersion Visio 11, 12, 14, 15, 16
		/// </summary>
		/// <param name="alignHorizontal">NetOffice.VisioApi.Enums.VisHorizontalAlignTypes alignHorizontal</param>
		/// <param name="alignVertical">NetOffice.VisioApi.Enums.VisVerticalAlignTypes alignVertical</param>
		/// <param name="glueToGuide">optional bool GlueToGuide = false</param>
		[SupportByVersion("Visio", 11,12,14,15,16)]
		void Align(NetOffice.VisioApi.Enums.VisHorizontalAlignTypes alignHorizontal, NetOffice.VisioApi.Enums.VisVerticalAlignTypes alignVertical, object glueToGuide);

		/// <summary>
		/// SupportByVersion Visio 11, 12, 14, 15, 16
		/// </summary>
		/// <param name="alignHorizontal">NetOffice.VisioApi.Enums.VisHorizontalAlignTypes alignHorizontal</param>
		/// <param name="alignVertical">NetOffice.VisioApi.Enums.VisVerticalAlignTypes alignVertical</param>
		[CustomMethod]
		[SupportByVersion("Visio", 11,12,14,15,16)]
		void Align(NetOffice.VisioApi.Enums.VisHorizontalAlignTypes alignHorizontal, NetOffice.VisioApi.Enums.VisVerticalAlignTypes alignVertical);

		/// <summary>
		/// SupportByVersion Visio 11, 12, 14, 15, 16
		/// </summary>
		/// <param name="distribute">NetOffice.VisioApi.Enums.VisDistributeTypes distribute</param>
		/// <param name="glueToGuide">optional bool GlueToGuide = false</param>
		[SupportByVersion("Visio", 11,12,14,15,16)]
		void Distribute(NetOffice.VisioApi.Enums.VisDistributeTypes distribute, object glueToGuide);

		/// <summary>
		/// SupportByVersion Visio 11, 12, 14, 15, 16
		/// </summary>
		/// <param name="distribute">NetOffice.VisioApi.Enums.VisDistributeTypes distribute</param>
		[CustomMethod]
		[SupportByVersion("Visio", 11,12,14,15,16)]
		void Distribute(NetOffice.VisioApi.Enums.VisDistributeTypes distribute);

		/// <summary>
		/// SupportByVersion Visio 11, 12, 14, 15, 16
		/// </summary>
		[SupportByVersion("Visio", 11,12,14,15,16)]
		void UpdateAlignmentBox();

		/// <summary>
		/// SupportByVersion Visio 11, 12, 14, 15, 16
		/// </summary>
		/// <param name="distance">Double distance</param>
		[SupportByVersion("Visio", 11,12,14,15,16)]
		void Offset(Double distance);

		/// <summary>
		/// SupportByVersion Visio 11, 12, 14, 15, 16
		/// </summary>
		[SupportByVersion("Visio", 11,12,14,15,16)]
		void ConnectShapes();

		/// <summary>
		/// SupportByVersion Visio 11, 12, 14, 15, 16
		/// </summary>
		/// <param name="flipDirection">NetOffice.VisioApi.Enums.VisFlipDirection flipDirection</param>
		/// <param name="flipType">optional NetOffice.VisioApi.Enums.VisFlipTypes FlipType = 0</param>
		/// <param name="blastGuards">optional bool BlastGuards = false</param>
		/// <param name="pinX">optional Double PinX = 0</param>
		/// <param name="pinY">optional Double PinY = 0</param>
		/// <param name="pinUnitsNameOrCode">optional object pinUnitsNameOrCode</param>
		[SupportByVersion("Visio", 11,12,14,15,16)]
		void Flip(NetOffice.VisioApi.Enums.VisFlipDirection flipDirection, object flipType, object blastGuards, object pinX, object pinY, object pinUnitsNameOrCode);

		/// <summary>
		/// SupportByVersion Visio 11, 12, 14, 15, 16
		/// </summary>
		/// <param name="flipDirection">NetOffice.VisioApi.Enums.VisFlipDirection flipDirection</param>
		[CustomMethod]
		[SupportByVersion("Visio", 11,12,14,15,16)]
		void Flip(NetOffice.VisioApi.Enums.VisFlipDirection flipDirection);

		/// <summary>
		/// SupportByVersion Visio 11, 12, 14, 15, 16
		/// </summary>
		/// <param name="flipDirection">NetOffice.VisioApi.Enums.VisFlipDirection flipDirection</param>
		/// <param name="flipType">optional NetOffice.VisioApi.Enums.VisFlipTypes FlipType = 0</param>
		[CustomMethod]
		[SupportByVersion("Visio", 11,12,14,15,16)]
		void Flip(NetOffice.VisioApi.Enums.VisFlipDirection flipDirection, object flipType);

		/// <summary>
		/// SupportByVersion Visio 11, 12, 14, 15, 16
		/// </summary>
		/// <param name="flipDirection">NetOffice.VisioApi.Enums.VisFlipDirection flipDirection</param>
		/// <param name="flipType">optional NetOffice.VisioApi.Enums.VisFlipTypes FlipType = 0</param>
		/// <param name="blastGuards">optional bool BlastGuards = false</param>
		[CustomMethod]
		[SupportByVersion("Visio", 11,12,14,15,16)]
		void Flip(NetOffice.VisioApi.Enums.VisFlipDirection flipDirection, object flipType, object blastGuards);

		/// <summary>
		/// SupportByVersion Visio 11, 12, 14, 15, 16
		/// </summary>
		/// <param name="flipDirection">NetOffice.VisioApi.Enums.VisFlipDirection flipDirection</param>
		/// <param name="flipType">optional NetOffice.VisioApi.Enums.VisFlipTypes FlipType = 0</param>
		/// <param name="blastGuards">optional bool BlastGuards = false</param>
		/// <param name="pinX">optional Double PinX = 0</param>
		[CustomMethod]
		[SupportByVersion("Visio", 11,12,14,15,16)]
		void Flip(NetOffice.VisioApi.Enums.VisFlipDirection flipDirection, object flipType, object blastGuards, object pinX);

		/// <summary>
		/// SupportByVersion Visio 11, 12, 14, 15, 16
		/// </summary>
		/// <param name="flipDirection">NetOffice.VisioApi.Enums.VisFlipDirection flipDirection</param>
		/// <param name="flipType">optional NetOffice.VisioApi.Enums.VisFlipTypes FlipType = 0</param>
		/// <param name="blastGuards">optional bool BlastGuards = false</param>
		/// <param name="pinX">optional Double PinX = 0</param>
		/// <param name="pinY">optional Double PinY = 0</param>
		[CustomMethod]
		[SupportByVersion("Visio", 11,12,14,15,16)]
		void Flip(NetOffice.VisioApi.Enums.VisFlipDirection flipDirection, object flipType, object blastGuards, object pinX, object pinY);

		/// <summary>
		/// SupportByVersion Visio 12, 14, 15, 16
		/// </summary>
		/// <param name="dataRecordsetID">Int32 dataRecordsetID</param>
		/// <param name="dataRowID">Int32 dataRowID</param>
		/// <param name="applyDataGraphicAfterLink">optional bool ApplyDataGraphicAfterLink = true</param>
		[SupportByVersion("Visio", 12,14,15,16)]
		void LinkToData(Int32 dataRecordsetID, Int32 dataRowID, object applyDataGraphicAfterLink);

		/// <summary>
		/// SupportByVersion Visio 12, 14, 15, 16
		/// </summary>
		/// <param name="dataRecordsetID">Int32 dataRecordsetID</param>
		/// <param name="dataRowID">Int32 dataRowID</param>
		[CustomMethod]
		[SupportByVersion("Visio", 12,14,15,16)]
		void LinkToData(Int32 dataRecordsetID, Int32 dataRowID);

		/// <summary>
		/// SupportByVersion Visio 12, 14, 15, 16
		/// </summary>
		/// <param name="dataRecordsetID">Int32 dataRecordsetID</param>
		[SupportByVersion("Visio", 12,14,15,16)]
		void BreakLinkToData(Int32 dataRecordsetID);

		/// <summary>
		/// SupportByVersion Visio 12, 14, 15, 16
		/// </summary>
		/// <param name="shapeIDs">Int32[] shapeIDs</param>
		[SupportByVersion("Visio", 12,14,15,16)]
		void GetIDs(out Int32[] shapeIDs);

		/// <summary>
		/// SupportByVersion Visio 12, 14, 15, 16
		/// </summary>
		/// <param name="dataRecordsetID">Int32 dataRecordsetID</param>
		/// <param name="columnNames">String[] columnNames</param>
		/// <param name="autoLinkFieldTypes">Int32[] autoLinkFieldTypes</param>
		/// <param name="fieldNames">String[] fieldNames</param>
		/// <param name="autoLinkBehavior">Int32 autoLinkBehavior</param>
		/// <param name="shapeIDs">Int32[] shapeIDs</param>
		[SupportByVersion("Visio", 12,14,15,16)]
		Int32 AutomaticLink(Int32 dataRecordsetID, String[] columnNames, Int32[] autoLinkFieldTypes, String[] fieldNames, Int32 autoLinkBehavior, out Int32[] shapeIDs);

		/// <summary>
		/// SupportByVersion Visio 14, 15, 16
		/// </summary>
		/// <param name="alignOrSpace">NetOffice.VisioApi.Enums.VisLayoutIncrementalType alignOrSpace</param>
		/// <param name="alignHorizontal">NetOffice.VisioApi.Enums.VisLayoutHorzAlignType alignHorizontal</param>
		/// <param name="alignVertical">NetOffice.VisioApi.Enums.VisLayoutVertAlignType alignVertical</param>
		/// <param name="spaceHorizontal">Double spaceHorizontal</param>
		/// <param name="spaceVertical">Double spaceVertical</param>
		/// <param name="unitCode">NetOffice.VisioApi.Enums.VisUnitCodes unitCode</param>
		[SupportByVersion("Visio", 14,15,16)]
		void LayoutIncremental(NetOffice.VisioApi.Enums.VisLayoutIncrementalType alignOrSpace, NetOffice.VisioApi.Enums.VisLayoutHorzAlignType alignHorizontal, NetOffice.VisioApi.Enums.VisLayoutVertAlignType alignVertical, Double spaceHorizontal, Double spaceVertical, NetOffice.VisioApi.Enums.VisUnitCodes unitCode);

		/// <summary>
		/// SupportByVersion Visio 14, 15, 16
		/// </summary>
		/// <param name="direction">NetOffice.VisioApi.Enums.VisLayoutDirection direction</param>
		[SupportByVersion("Visio", 14,15,16)]
		void LayoutChangeDirection(NetOffice.VisioApi.Enums.VisLayoutDirection direction);

		/// <summary>
		/// SupportByVersion Visio 14, 15, 16
		/// </summary>
		[SupportByVersion("Visio", 14,15,16)]
		void AvoidPageBreaks();

		/// <summary>
		/// SupportByVersion Visio 14, 15, 16
		/// </summary>
		/// <param name="direction">NetOffice.VisioApi.Enums.VisResizeDirection direction</param>
		/// <param name="distance">Double distance</param>
		/// <param name="unitCode">NetOffice.VisioApi.Enums.VisUnitCodes unitCode</param>
		[SupportByVersion("Visio", 14,15,16)]
		void Resize(NetOffice.VisioApi.Enums.VisResizeDirection direction, Double distance, NetOffice.VisioApi.Enums.VisUnitCodes unitCode);

		/// <summary>
		/// SupportByVersion Visio 14, 15, 16
		/// </summary>
		[SupportByVersion("Visio", 14,15,16)]
		void AddToContainers();

		/// <summary>
		/// SupportByVersion Visio 14, 15, 16
		/// </summary>
		[SupportByVersion("Visio", 14,15,16)]
		void RemoveFromContainers();

		/// <summary>
		/// SupportByVersion Visio 14, 15, 16
		/// </summary>
		/// <param name="page">NetOffice.VisioApi.IVPage page</param>
		/// <param name="objectToDrop">object objectToDrop</param>
		/// <param name="newShape">optional NetOffice.VisioApi.IVShape NewShape = 0</param>
		[SupportByVersion("Visio", 14,15,16)]
		[BaseResult]
		NetOffice.VisioApi.IVSelection MoveToSubprocess(NetOffice.VisioApi.IVPage page, object objectToDrop, object newShape);

		/// <summary>
		/// SupportByVersion Visio 14, 15, 16
		/// </summary>
		/// <param name="page">NetOffice.VisioApi.IVPage page</param>
		/// <param name="objectToDrop">object objectToDrop</param>
		[CustomMethod]
		[BaseResult]
		[SupportByVersion("Visio", 14,15,16)]
		NetOffice.VisioApi.IVSelection MoveToSubprocess(NetOffice.VisioApi.IVPage page, object objectToDrop);

		/// <summary>
		/// SupportByVersion Visio 14, 15, 16
		/// </summary>
		/// <param name="delFlags">Int32 delFlags</param>
		[SupportByVersion("Visio", 14,15,16)]
		void DeleteEx(Int32 delFlags);

		/// <summary>
		/// SupportByVersion Visio 14, 15, 16
		/// </summary>
		/// <param name="nestedOptions">NetOffice.VisioApi.Enums.VisContainerNested nestedOptions</param>
		[SupportByVersion("Visio", 14,15,16)]
		Int32[] GetContainers(NetOffice.VisioApi.Enums.VisContainerNested nestedOptions);

		/// <summary>
		/// SupportByVersion Visio 14, 15, 16
		/// </summary>
		/// <param name="nestedOptions">NetOffice.VisioApi.Enums.VisContainerNested nestedOptions</param>
		[SupportByVersion("Visio", 14,15,16)]
		Int32[] GetCallouts(NetOffice.VisioApi.Enums.VisContainerNested nestedOptions);

		/// <summary>
		/// SupportByVersion Visio 14, 15, 16
		/// </summary>
		[SupportByVersion("Visio", 14,15,16)]
		Int32[] MemberOfContainersUnion();

		/// <summary>
		/// SupportByVersion Visio 14, 15, 16
		/// </summary>
		[SupportByVersion("Visio", 14,15,16)]
		Int32[] MemberOfContainersIntersection();

		/// <summary>
		/// SupportByVersion Visio 14, 15, 16
		/// </summary>
		/// <param name="formatType">NetOffice.VisioApi.Enums.VisContainerFormatType formatType</param>
		/// <param name="formatValue">optional object FormatValue = 0</param>
		[SupportByVersion("Visio", 14,15,16)]
		Int32[] SetContainerFormat(NetOffice.VisioApi.Enums.VisContainerFormatType formatType, object formatValue);

		/// <summary>
		/// SupportByVersion Visio 14, 15, 16
		/// </summary>
		/// <param name="formatType">NetOffice.VisioApi.Enums.VisContainerFormatType formatType</param>
		[CustomMethod]
		[SupportByVersion("Visio", 14,15,16)]
		Int32[] SetContainerFormat(NetOffice.VisioApi.Enums.VisContainerFormatType formatType);

		/// <summary>
		/// SupportByVersion Visio 15,16
		/// </summary>
		/// <param name="masterOrMasterShortcutToDrop">object masterOrMasterShortcutToDrop</param>
		/// <param name="replaceFlags">optional Int32 ReplaceFlags = 0</param>
		[SupportByVersion("Visio", 15, 16)]
		NetOffice.VisioApi.IVShape[] ReplaceShape(object masterOrMasterShortcutToDrop, object replaceFlags);

		/// <summary>
		/// SupportByVersion Visio 15,16
		/// </summary>
		/// <param name="masterOrMasterShortcutToDrop">object masterOrMasterShortcutToDrop</param>
		[CustomMethod]
		[SupportByVersion("Visio", 15, 16)]
		NetOffice.VisioApi.IVShape[] ReplaceShape(object masterOrMasterShortcutToDrop);

		/// <summary>
		/// SupportByVersion Visio 15,16
		/// </summary>
		/// <param name="lineMatrix">NetOffice.VisioApi.Enums.VisQuickStyleMatrixIndices lineMatrix</param>
		/// <param name="fillMatrix">NetOffice.VisioApi.Enums.VisQuickStyleMatrixIndices fillMatrix</param>
		/// <param name="effectsMatrix">NetOffice.VisioApi.Enums.VisQuickStyleMatrixIndices effectsMatrix</param>
		/// <param name="fontMatrix">NetOffice.VisioApi.Enums.VisQuickStyleMatrixIndices fontMatrix</param>
		/// <param name="lineColor">NetOffice.VisioApi.Enums.VisQuickStyleColors lineColor</param>
		/// <param name="fillColor">NetOffice.VisioApi.Enums.VisQuickStyleColors fillColor</param>
		/// <param name="shadowColor">NetOffice.VisioApi.Enums.VisQuickStyleColors shadowColor</param>
		/// <param name="fontColor">NetOffice.VisioApi.Enums.VisQuickStyleColors fontColor</param>
		[SupportByVersion("Visio", 15, 16)]
		void SetQuickStyle(NetOffice.VisioApi.Enums.VisQuickStyleMatrixIndices lineMatrix, NetOffice.VisioApi.Enums.VisQuickStyleMatrixIndices fillMatrix, NetOffice.VisioApi.Enums.VisQuickStyleMatrixIndices effectsMatrix, NetOffice.VisioApi.Enums.VisQuickStyleMatrixIndices fontMatrix, NetOffice.VisioApi.Enums.VisQuickStyleColors lineColor, NetOffice.VisioApi.Enums.VisQuickStyleColors fillColor, NetOffice.VisioApi.Enums.VisQuickStyleColors shadowColor, NetOffice.VisioApi.Enums.VisQuickStyleColors fontColor);

        #endregion

        #region IEnumerable<NetOffice.VisioApi.IVShape>

        /// <summary>
        /// SupportByVersion Visio, 11,12,14,15,16
        /// </summary>
        [SupportByVersion("Visio", 11, 12, 14, 15, 16)]
        new IEnumerator<NetOffice.VisioApi.IVShape> GetEnumerator();

        #endregion
    }
}
