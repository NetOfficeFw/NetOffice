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
	/// DispatchInterface IColumnHeaders 
	/// SupportByVersion MSComctlLib, 6
	/// </summary>
	[SupportByVersion("MSComctlLib", 6)]
	[EntityType(EntityType.IsDispatchInterface), BaseType, Enumerator(Enumerator.Reference, EnumeratorInvoke.Method, "MSComctlLib", 6), HasIndexProperty(IndexInvoke.Property, "Item")]
	[TypeId("BDD1F050-858B-11D1-B16A-00C0F0283628")]
    [CoClassSource(typeof(NetOffice.MSComctlLibApi.ColumnHeaders))]
    public interface IColumnHeaders : ICOMObject, IEnumerableProvider<NetOffice.MSComctlLibApi.IColumnHeader>
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
		NetOffice.MSComctlLibApi.IColumnHeader get_ControlDefault(object index);

		/// <summary>
		/// SupportByVersion MSComctlLib 6
		/// Alias for get_ControlDefault
		/// </summary>
		/// <param name="index">object index</param>
		[SupportByVersion("MSComctlLib", 6), Redirect("get_ControlDefault")]
		NetOffice.MSComctlLibApi.IColumnHeader ControlDefault(object index);

		/// <summary>
		/// SupportByVersion MSComctlLib 6
		/// Get
		/// </summary>
		/// <param name="index">object index</param>
		[SupportByVersion("MSComctlLib", 6)]
		[BaseResult]
		[NetRuntimeSystem.Runtime.CompilerServices.IndexerName("Item"), IndexProperty]
		NetOffice.MSComctlLibApi.IColumnHeader this[object index] { get; }

		#endregion

		#region Methods

		/// <summary>
		/// SupportByVersion MSComctlLib 6
		/// </summary>
		/// <param name="index">optional object index</param>
		/// <param name="key">optional object key</param>
		/// <param name="text">optional object text</param>
		/// <param name="width">optional object width</param>
		/// <param name="alignment">optional object alignment</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		[BaseResult]
		[SupportByVersion("MSComctlLib", 6)]
		NetOffice.MSComctlLibApi.IColumnHeader Add_PreVB98(object index, object key, object text, object width, object alignment);

		/// <summary>
		/// SupportByVersion MSComctlLib 6
		/// </summary>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		[BaseResult]
		[CustomMethod]
		[SupportByVersion("MSComctlLib", 6)]
		NetOffice.MSComctlLibApi.IColumnHeader Add_PreVB98();

		/// <summary>
		/// SupportByVersion MSComctlLib 6
		/// </summary>
		/// <param name="index">optional object index</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		[BaseResult]
		[CustomMethod]
		[SupportByVersion("MSComctlLib", 6)]
		NetOffice.MSComctlLibApi.IColumnHeader Add_PreVB98(object index);

		/// <summary>
		/// SupportByVersion MSComctlLib 6
		/// </summary>
		/// <param name="index">optional object index</param>
		/// <param name="key">optional object key</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		[BaseResult]
		[CustomMethod]
		[SupportByVersion("MSComctlLib", 6)]
		NetOffice.MSComctlLibApi.IColumnHeader Add_PreVB98(object index, object key);

		/// <summary>
		/// SupportByVersion MSComctlLib 6
		/// </summary>
		/// <param name="index">optional object index</param>
		/// <param name="key">optional object key</param>
		/// <param name="text">optional object text</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		[BaseResult]
		[CustomMethod]
		[SupportByVersion("MSComctlLib", 6)]
		NetOffice.MSComctlLibApi.IColumnHeader Add_PreVB98(object index, object key, object text);

		/// <summary>
		/// SupportByVersion MSComctlLib 6
		/// </summary>
		/// <param name="index">optional object index</param>
		/// <param name="key">optional object key</param>
		/// <param name="text">optional object text</param>
		/// <param name="width">optional object width</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		[BaseResult]
		[CustomMethod]
		[SupportByVersion("MSComctlLib", 6)]
		NetOffice.MSComctlLibApi.IColumnHeader Add_PreVB98(object index, object key, object text, object width);

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

		/// <summary>
		/// SupportByVersion MSComctlLib 6
		/// </summary>
		/// <param name="index">optional object index</param>
		/// <param name="key">optional object key</param>
		/// <param name="text">optional object text</param>
		/// <param name="width">optional object width</param>
		/// <param name="alignment">optional object alignment</param>
		/// <param name="icon">optional object icon</param>
		[SupportByVersion("MSComctlLib", 6)]
		[BaseResult]
		NetOffice.MSComctlLibApi.IColumnHeader Add(object index, object key, object text, object width, object alignment, object icon);

		/// <summary>
		/// SupportByVersion MSComctlLib 6
		/// </summary>
		[CustomMethod]
		[BaseResult]
		[SupportByVersion("MSComctlLib", 6)]
		NetOffice.MSComctlLibApi.IColumnHeader Add();

		/// <summary>
		/// SupportByVersion MSComctlLib 6
		/// </summary>
		/// <param name="index">optional object index</param>
		[CustomMethod]
		[BaseResult]
		[SupportByVersion("MSComctlLib", 6)]
		NetOffice.MSComctlLibApi.IColumnHeader Add(object index);

		/// <summary>
		/// SupportByVersion MSComctlLib 6
		/// </summary>
		/// <param name="index">optional object index</param>
		/// <param name="key">optional object key</param>
		[CustomMethod]
		[BaseResult]
		[SupportByVersion("MSComctlLib", 6)]
		NetOffice.MSComctlLibApi.IColumnHeader Add(object index, object key);

		/// <summary>
		/// SupportByVersion MSComctlLib 6
		/// </summary>
		/// <param name="index">optional object index</param>
		/// <param name="key">optional object key</param>
		/// <param name="text">optional object text</param>
		[CustomMethod]
		[BaseResult]
		[SupportByVersion("MSComctlLib", 6)]
		NetOffice.MSComctlLibApi.IColumnHeader Add(object index, object key, object text);

		/// <summary>
		/// SupportByVersion MSComctlLib 6
		/// </summary>
		/// <param name="index">optional object index</param>
		/// <param name="key">optional object key</param>
		/// <param name="text">optional object text</param>
		/// <param name="width">optional object width</param>
		[CustomMethod]
		[BaseResult]
		[SupportByVersion("MSComctlLib", 6)]
		NetOffice.MSComctlLibApi.IColumnHeader Add(object index, object key, object text, object width);

		/// <summary>
		/// SupportByVersion MSComctlLib 6
		/// </summary>
		/// <param name="index">optional object index</param>
		/// <param name="key">optional object key</param>
		/// <param name="text">optional object text</param>
		/// <param name="width">optional object width</param>
		/// <param name="alignment">optional object alignment</param>
		[CustomMethod]
		[BaseResult]
		[SupportByVersion("MSComctlLib", 6)]
		NetOffice.MSComctlLibApi.IColumnHeader Add(object index, object key, object text, object width, object alignment);

        #endregion

        #region IEnumerable<NetOffice.MSComctlLibApi.IColumnHeader>

        /// <summary>
        /// SupportByVersion MSComctlLib, 6
        /// </summary>
        [SupportByVersion("MSComctlLib", 6)]
        new IEnumerator<NetOffice.MSComctlLibApi.IColumnHeader> GetEnumerator();

        #endregion
    }
}
