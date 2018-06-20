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
	/// DispatchInterface IListItems 
	/// SupportByVersion MSComctlLib, 6
	/// </summary>
	[SupportByVersion("MSComctlLib", 6)]
	[EntityType(EntityType.IsDispatchInterface), BaseType, Enumerator(Enumerator.Reference, EnumeratorInvoke.Method, "MSComctlLib", 6), HasIndexProperty(IndexInvoke.Property, "Item")]
	[TypeId("BDD1F04C-858B-11D1-B16A-00C0F0283628")]
    [CoClassSource(typeof(NetOffice.MSComctlLibApi.ListItems))]
    public interface IListItems : ICOMObject, IEnumerableProvider<NetOffice.MSComctlLibApi.IListItem>
	{
		#region Properties

		/// <summary>
		/// SupportByVersion MSComctlLib 6
		/// Get/Set
		/// </summary>
		[SupportByVersion("MSComctlLib", 6)]
		Int32 Count { get; set; }

		/// <summary>
		/// SupportByVersion MSComctlLib 6
		/// Get
		/// </summary>
		/// <param name="index">object index</param>
		[SupportByVersion("MSComctlLib", 6)]
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		NetOffice.MSComctlLibApi.IListItem get_ControlDefault(object index);

		/// <summary>
		/// SupportByVersion MSComctlLib 6
		/// Alias for get_ControlDefault
		/// </summary>
		/// <param name="index">object index</param>
		[SupportByVersion("MSComctlLib", 6), Redirect("get_ControlDefault")]
		NetOffice.MSComctlLibApi.IListItem ControlDefault(object index);

		/// <summary>
		/// SupportByVersion MSComctlLib 6
		/// Get
		/// </summary>
		/// <param name="index">object index</param>
		[SupportByVersion("MSComctlLib", 6)]
		[BaseResult]
		[NetRuntimeSystem.Runtime.CompilerServices.IndexerName("Item"), IndexProperty]
		NetOffice.MSComctlLibApi.IListItem this[object index] { get; }

		#endregion

		#region Methods

		/// <summary>
		/// SupportByVersion MSComctlLib 6
		/// </summary>
		/// <param name="index">optional object index</param>
		/// <param name="key">optional object key</param>
		/// <param name="text">optional object text</param>
		/// <param name="icon">optional object icon</param>
		/// <param name="smallIcon">optional object smallIcon</param>
		[SupportByVersion("MSComctlLib", 6)]
		[BaseResult]
		NetOffice.MSComctlLibApi.IListItem Add(object index, object key, object text, object icon, object smallIcon);

		/// <summary>
		/// SupportByVersion MSComctlLib 6
		/// </summary>
		[CustomMethod]
		[BaseResult]
		[SupportByVersion("MSComctlLib", 6)]
		NetOffice.MSComctlLibApi.IListItem Add();

		/// <summary>
		/// SupportByVersion MSComctlLib 6
		/// </summary>
		/// <param name="index">optional object index</param>
		[CustomMethod]
		[BaseResult]
		[SupportByVersion("MSComctlLib", 6)]
		NetOffice.MSComctlLibApi.IListItem Add(object index);

		/// <summary>
		/// SupportByVersion MSComctlLib 6
		/// </summary>
		/// <param name="index">optional object index</param>
		/// <param name="key">optional object key</param>
		[CustomMethod]
		[BaseResult]
		[SupportByVersion("MSComctlLib", 6)]
		NetOffice.MSComctlLibApi.IListItem Add(object index, object key);

		/// <summary>
		/// SupportByVersion MSComctlLib 6
		/// </summary>
		/// <param name="index">optional object index</param>
		/// <param name="key">optional object key</param>
		/// <param name="text">optional object text</param>
		[CustomMethod]
		[BaseResult]
		[SupportByVersion("MSComctlLib", 6)]
		NetOffice.MSComctlLibApi.IListItem Add(object index, object key, object text);

		/// <summary>
		/// SupportByVersion MSComctlLib 6
		/// </summary>
		/// <param name="index">optional object index</param>
		/// <param name="key">optional object key</param>
		/// <param name="text">optional object text</param>
		/// <param name="icon">optional object icon</param>
		[CustomMethod]
		[BaseResult]
		[SupportByVersion("MSComctlLib", 6)]
		NetOffice.MSComctlLibApi.IListItem Add(object index, object key, object text, object icon);

		/// <summary>
		/// SupportByVersion MSComctlLib 6
		/// </summary>
		[SupportByVersion("MSComctlLib", 6)]
		void Clear();

		/// <summary>
		/// SupportByVersion MSComctlLib 6
		/// </summary>
		/// <param name="index">object index</param>
		[SupportByVersion("MSComctlLib", 6)]
		void Remove(object index);

        #endregion

        #region IEnumerable<NetOffice.MSComctlLibApi.IListItem>

        /// <summary>
        /// SupportByVersion MSComctlLib, 6
        /// </summary>
        [SupportByVersion("MSComctlLib", 6)]
        new IEnumerator<NetOffice.MSComctlLibApi.IListItem> GetEnumerator();

        #endregion
    }
}
