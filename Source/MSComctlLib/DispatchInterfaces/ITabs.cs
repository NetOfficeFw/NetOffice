using System.Collections;
using System.Collections.Generic;
using System;
using NetRuntimeSystem = System;
using System.ComponentModel;
using NetOffice.Attributes;
using NetOffice.CollectionsGeneric;

namespace NetOffice.MSComctlLibApi
{
	/// <summary>
	/// DispatchInterface ITabs 
	/// SupportByVersion MSComctlLib, 6
	/// </summary>
	[SupportByVersion("MSComctlLib", 6)]
	[EntityType(EntityType.IsDispatchInterface), BaseType, Enumerator(Enumerator.Reference, EnumeratorInvoke.Method, "MSComctlLib", 6), HasIndexProperty(IndexInvoke.Property, "Item")]
	[TypeId("1EFB6597-857C-11D1-B16A-00C0F0283628")]
    [CoClassSource(typeof(NetOffice.MSComctlLibApi.Tabs))]
    public interface ITabs : ICOMObject, IEnumerableProvider<NetOffice.MSComctlLibApi.ITab>
	{
		#region Properties

		/// <summary>
		/// SupportByVersion MSComctlLib 6
		/// Get/Set
		/// </summary>
		[SupportByVersion("MSComctlLib", 6)]
		Int16 Count { get; set; }

		/// <summary>
		/// SupportByVersion MSComctlLib 6
		/// Get/Set
		/// </summary>
		/// <param name="pvIndex">object pvIndex</param>
		[SupportByVersion("MSComctlLib", 6)]
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		NetOffice.MSComctlLibApi.ITab get_ControlDefault(object pvIndex);

        /// <summary>
        /// SupportByVersion MSComctlLib 6
        /// Get/Set
        /// </summary>
        /// <param name="pvIndex">object pvIndex</param>
        /// <param name="value">NetOffice.MSComctlLibApi.ITab value</param>
        [SupportByVersion("MSComctlLib", 6)]
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		void set_ControlDefault(object pvIndex, NetOffice.MSComctlLibApi.ITab value);

		/// <summary>
		/// SupportByVersion MSComctlLib 6
		/// Alias for get_ControlDefault
		/// </summary>
		/// <param name="pvIndex">object pvIndex</param>
		[SupportByVersion("MSComctlLib", 6), Redirect("get_ControlDefault")]
		NetOffice.MSComctlLibApi.ITab ControlDefault(object pvIndex);

		/// <summary>
		/// SupportByVersion MSComctlLib 6
		/// Get/Set
		/// </summary>
		/// <param name="pvIndex">object pvIndex</param>
		[SupportByVersion("MSComctlLib", 6)]
		[BaseResult]
		[NetRuntimeSystem.Runtime.CompilerServices.IndexerName("Item"), IndexProperty]
		NetOffice.MSComctlLibApi.ITab this[object pvIndex] { get; set; }

		#endregion

		#region Methods

		/// <summary>
		/// SupportByVersion MSComctlLib 6
		/// </summary>
		/// <param name="pvIndex">object pvIndex</param>
		[SupportByVersion("MSComctlLib", 6)]
		void Remove(object pvIndex);

		/// <summary>
		/// SupportByVersion MSComctlLib 6
		/// </summary>
		[SupportByVersion("MSComctlLib", 6)]
		void Clear();

		/// <summary>
		/// SupportByVersion MSComctlLib 6
		/// </summary>
		/// <param name="pvIndex">optional object pvIndex</param>
		/// <param name="pvKey">optional object pvKey</param>
		/// <param name="pvCaption">optional object pvCaption</param>
		/// <param name="pvImage">optional object pvImage</param>
		[SupportByVersion("MSComctlLib", 6)]
		[BaseResult]
		NetOffice.MSComctlLibApi.ITab Add(object pvIndex, object pvKey, object pvCaption, object pvImage);

		/// <summary>
		/// SupportByVersion MSComctlLib 6
		/// </summary>
		[CustomMethod]
		[BaseResult]
		[SupportByVersion("MSComctlLib", 6)]
		NetOffice.MSComctlLibApi.ITab Add();

		/// <summary>
		/// SupportByVersion MSComctlLib 6
		/// </summary>
		/// <param name="pvIndex">optional object pvIndex</param>
		[CustomMethod]
		[BaseResult]
		[SupportByVersion("MSComctlLib", 6)]
		NetOffice.MSComctlLibApi.ITab Add(object pvIndex);

		/// <summary>
		/// SupportByVersion MSComctlLib 6
		/// </summary>
		/// <param name="pvIndex">optional object pvIndex</param>
		/// <param name="pvKey">optional object pvKey</param>
		[CustomMethod]
		[BaseResult]
		[SupportByVersion("MSComctlLib", 6)]
		NetOffice.MSComctlLibApi.ITab Add(object pvIndex, object pvKey);

		/// <summary>
		/// SupportByVersion MSComctlLib 6
		/// </summary>
		/// <param name="pvIndex">optional object pvIndex</param>
		/// <param name="pvKey">optional object pvKey</param>
		/// <param name="pvCaption">optional object pvCaption</param>
		[CustomMethod]
		[BaseResult]
		[SupportByVersion("MSComctlLib", 6)]
		NetOffice.MSComctlLibApi.ITab Add(object pvIndex, object pvKey, object pvCaption);

        #endregion

        #region IEnumerable<NetOffice.MSComctlLibApi.ITab>

        /// <summary>
        /// SupportByVersion MSComctlLib, 6
        /// </summary>
        [SupportByVersion("MSComctlLib", 6)]
        new IEnumerator<NetOffice.MSComctlLibApi.ITab> GetEnumerator();

        #endregion
    }
}
