using System.Collections;
using System.Collections.Generic;
using System;
using NetRuntimeSystem = System;
using System.ComponentModel;
using NetOffice.Attributes;
using NetOffice.CollectionsGeneric;

namespace NetOffice.PowerPointApi
{
	/// <summary>
	/// DispatchInterface PPRadioCluster 
	/// SupportByVersion PowerPoint, 9
	/// </summary>
	[SupportByVersion("PowerPoint", 9)]
	[EntityType(EntityType.IsDispatchInterface), Enumerator(Enumerator.Reference, EnumeratorInvoke.Custom, "PowerPoint", 9), HasIndexProperty(IndexInvoke.Method, "Item")]
	[TypeId("914934AB-5A91-11CF-8700-00AA0060263B")]
	public interface PPRadioCluster : PPControl, IEnumerableProvider<NetOffice.PowerPointApi.PPRadioButton>
	{
		#region Properties

		/// <summary>
		/// SupportByVersion PowerPoint 9
		/// Get
		/// </summary>
		[SupportByVersion("PowerPoint", 9)]
		Int32 Count { get; }

		/// <summary>
		/// SupportByVersion PowerPoint 9
		/// Get
		/// </summary>
		[SupportByVersion("PowerPoint", 9)]
		NetOffice.PowerPointApi.PPRadioButton Selected { get; }

		/// <summary>
		/// SupportByVersion PowerPoint 9
		/// Get/Set
		/// </summary>
		[SupportByVersion("PowerPoint", 9)]
		string OnClick { get; set; }

		#endregion

		#region Methods

		/// <summary>
		/// SupportByVersion PowerPoint 9
		/// </summary>
		/// <param name="index">object index</param>
		[SupportByVersion("PowerPoint", 9)]
		[NetRuntimeSystem.Runtime.CompilerServices.IndexerName("Item"), IndexProperty]
		NetOffice.PowerPointApi.PPRadioButton this[object index] { get; }

		/// <summary>
		/// SupportByVersion PowerPoint 9
		/// </summary>
		/// <param name="left">Single left</param>
		/// <param name="top">Single top</param>
		/// <param name="width">Single width</param>
		/// <param name="height">Single height</param>
		[SupportByVersion("PowerPoint", 9)]
		NetOffice.PowerPointApi.PPRadioButton Add(Single left, Single top, Single width, Single height);

        #endregion
        #region IEnumerable<NetOffice.PowerPointApi.PPRadioButton>

        /// <summary>
        /// SupportByVersion PowerPoint, 9
        /// This is a custom enumerator from NetOffice
        /// </summary>
        [SupportByVersion("PowerPoint", 9)]
        [CustomEnumerator]
        new IEnumerator<NetOffice.PowerPointApi.PPRadioButton> GetEnumerator();

        #endregion

    }
}
