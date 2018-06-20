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
	/// DispatchInterface IButtonMenus 
	/// SupportByVersion MSComctlLib, 6
	/// </summary>
	[SupportByVersion("MSComctlLib", 6)]
	[EntityType(EntityType.IsDispatchInterface), BaseType, Enumerator(Enumerator.Reference, EnumeratorInvoke.Method, "MSComctlLib", 6), HasIndexProperty(IndexInvoke.Property, "Item")]
	[TypeId("66833FEB-8583-11D1-B16A-00C0F0283628")]
    [CoClassSource(typeof(NetOffice.MSComctlLibApi.ButtonMenus))]
    public interface IButtonMenus : ICOMObject, IEnumerableProvider<NetOffice.MSComctlLibApi.IButtonMenu>
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
		/// <param name="index">object index</param>
		[SupportByVersion("MSComctlLib", 6)]
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		NetOffice.MSComctlLibApi.IButtonMenu get_ControlDefault(object index);

        /// <summary>
        /// SupportByVersion MSComctlLib 6
        /// Get/Set
        /// </summary>
        /// <param name="index">object index</param>
        /// <param name="value">NetOffice.MSComctlLibApi.IButtonMenu value</param>
        [SupportByVersion("MSComctlLib", 6)]
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		void set_ControlDefault(object index, NetOffice.MSComctlLibApi.IButtonMenu value);

		/// <summary>
		/// SupportByVersion MSComctlLib 6
		/// Alias for get_ControlDefault
		/// </summary>
		/// <param name="index">object index</param>
		[SupportByVersion("MSComctlLib", 6), Redirect("get_ControlDefault")]
		NetOffice.MSComctlLibApi.IButtonMenu ControlDefault(object index);

		/// <summary>
		/// SupportByVersion MSComctlLib 6
		/// Get/Set
		/// </summary>
		/// <param name="index">object index</param>
		[SupportByVersion("MSComctlLib", 6)]
		[BaseResult]
		[NetRuntimeSystem.Runtime.CompilerServices.IndexerName("Item"), IndexProperty]
		NetOffice.MSComctlLibApi.IButtonMenu this[object index] { get; set; }

		#endregion

		#region Methods

		/// <summary>
		/// SupportByVersion MSComctlLib 6
		/// </summary>
		/// <param name="index">object index</param>
		[SupportByVersion("MSComctlLib", 6)]
		void Remove(object index);

		/// <summary>
		/// SupportByVersion MSComctlLib 6
		/// </summary>
		[SupportByVersion("MSComctlLib", 6)]
		void Clear();

		/// <summary>
		/// SupportByVersion MSComctlLib 6
		/// </summary>
		/// <param name="index">optional object index</param>
		/// <param name="key">optional object key</param>
		/// <param name="text">optional object text</param>
		[SupportByVersion("MSComctlLib", 6)]
		[BaseResult]
		NetOffice.MSComctlLibApi.IButtonMenu Add(object index, object key, object text);

		/// <summary>
		/// SupportByVersion MSComctlLib 6
		/// </summary>
		[CustomMethod]
		[BaseResult]
		[SupportByVersion("MSComctlLib", 6)]
		NetOffice.MSComctlLibApi.IButtonMenu Add();

		/// <summary>
		/// SupportByVersion MSComctlLib 6
		/// </summary>
		/// <param name="index">optional object index</param>
		[CustomMethod]
		[BaseResult]
		[SupportByVersion("MSComctlLib", 6)]
		NetOffice.MSComctlLibApi.IButtonMenu Add(object index);

		/// <summary>
		/// SupportByVersion MSComctlLib 6
		/// </summary>
		/// <param name="index">optional object index</param>
		/// <param name="key">optional object key</param>
		[CustomMethod]
		[BaseResult]
		[SupportByVersion("MSComctlLib", 6)]
		NetOffice.MSComctlLibApi.IButtonMenu Add(object index, object key);

        #endregion

        #region IEnumerable<NetOffice.MSComctlLibApi.IButtonMenu>

        /// <summary>
        /// SupportByVersion MSComctlLib, 6
        /// </summary>
        [SupportByVersion("MSComctlLib", 6)]
        new IEnumerator<NetOffice.MSComctlLibApi.IButtonMenu> GetEnumerator();

        #endregion
    }
}
