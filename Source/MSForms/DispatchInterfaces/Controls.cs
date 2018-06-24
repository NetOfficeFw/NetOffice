using System.Collections;
using System.Collections.Generic;
using System;
using NetRuntimeSystem = System;
using System.ComponentModel;
using NetOffice.Attributes;
using NetOffice.CollectionsGeneric;

namespace NetOffice.MSFormsApi
{
	/// <summary>
	/// DispatchInterface Controls 
	/// SupportByVersion MSForms, 2
	/// </summary>
	[SupportByVersion("MSForms", 2)]
	[EntityType(EntityType.IsDispatchInterface), Enumerator(Enumerator.Reference, EnumeratorInvoke.Property, "MSForms", 2), HasIndexProperty(IndexInvoke.Method, "Item")]
	[TypeId("04598FC7-866C-11CF-AB7C-00AA00C08FCF")]
	public interface Controls : ICOMObject, NetOffice.CollectionsGeneric.IEnumerableProvider<object>
	{
		#region Properties

		/// <summary>
		/// SupportByVersion MSForms 2
		/// Get
		/// </summary>
		[SupportByVersion("MSForms", 2)]
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		Int32 Count { get; }

		#endregion

		#region Methods

		/// <summary>
		/// SupportByVersion MSForms 2
		/// </summary>
		/// <param name="varg">object varg</param>
		[SupportByVersion("MSForms", 2)]
		[NetRuntimeSystem.Runtime.CompilerServices.IndexerName("Item"), IndexProperty]
		object this[object varg] { get; }

		/// <summary>
		/// SupportByVersion MSForms 2
		/// </summary>
		[SupportByVersion("MSForms", 2)]
		void Clear();

		/// <summary>
		/// SupportByVersion MSForms 2
		/// </summary>
		/// <param name="cx">Int32 cx</param>
		/// <param name="cy">Int32 cy</param>
		[SupportByVersion("MSForms", 2)]
		void _Move(Int32 cx, Int32 cy);

		/// <summary>
		/// SupportByVersion MSForms 2
		/// </summary>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		[SupportByVersion("MSForms", 2)]
		void SelectAll();

		/// <summary>
		/// SupportByVersion MSForms 2
		/// </summary>
		/// <param name="clsid">Int32 clsid</param>
		[SupportByVersion("MSForms", 2)]
		NetOffice.MSFormsApi.Control _AddByClass(Int32 clsid);

		/// <summary>
		/// SupportByVersion MSForms 2
		/// </summary>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		[SupportByVersion("MSForms", 2)]
		void AlignToGrid();

		/// <summary>
		/// SupportByVersion MSForms 2
		/// </summary>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		[SupportByVersion("MSForms", 2)]
		void BringForward();

		/// <summary>
		/// SupportByVersion MSForms 2
		/// </summary>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		[SupportByVersion("MSForms", 2)]
		void BringToFront();

		/// <summary>
		/// SupportByVersion MSForms 2
		/// </summary>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		[SupportByVersion("MSForms", 2)]
		void Copy();

		/// <summary>
		/// SupportByVersion MSForms 2
		/// </summary>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		[SupportByVersion("MSForms", 2)]
		void Cut();

		/// <summary>
		/// SupportByVersion MSForms 2
		/// </summary>
		[SupportByVersion("MSForms", 2)]
		object Enum();

		/// <summary>
		/// SupportByVersion MSForms 2
		/// </summary>
		/// <param name="lIndex">Int32 lIndex</param>
		[SupportByVersion("MSForms", 2)]
		NetOffice.MSFormsApi.Control _GetItemByIndex(Int32 lIndex);

		/// <summary>
		/// SupportByVersion MSForms 2
		/// </summary>
		/// <param name="pstr">string pstr</param>
		[SupportByVersion("MSForms", 2)]
		NetOffice.MSFormsApi.Control _GetItemByName(string pstr);

		/// <summary>
		/// SupportByVersion MSForms 2
		/// </summary>
		/// <param name="iD">Int32 iD</param>
		[SupportByVersion("MSForms", 2)]
		NetOffice.MSFormsApi.Control _GetItemByID(Int32 iD);

		/// <summary>
		/// SupportByVersion MSForms 2
		/// </summary>
		[SupportByVersion("MSForms", 2)]
		void SendBackward();

		/// <summary>
		/// SupportByVersion MSForms 2
		/// </summary>
		[SupportByVersion("MSForms", 2)]
		void SendToBack();

		/// <summary>
		/// SupportByVersion MSForms 2
		/// </summary>
		/// <param name="cx">Single cx</param>
		/// <param name="cy">Single cy</param>
		[SupportByVersion("MSForms", 2)]
		void Move(Single cx, Single cy);

		/// <summary>
		/// SupportByVersion MSForms 2
		/// </summary>
		/// <param name="bstrProgID">string bstrProgID</param>
		/// <param name="name">optional object name</param>
		/// <param name="visible">optional object visible</param>
		[SupportByVersion("MSForms", 2)]
		NetOffice.MSFormsApi.Control Add(string bstrProgID, object name, object visible);

		/// <summary>
		/// SupportByVersion MSForms 2
		/// </summary>
		/// <param name="bstrProgID">string bstrProgID</param>
		[CustomMethod]
		[SupportByVersion("MSForms", 2)]
		NetOffice.MSFormsApi.Control Add(string bstrProgID);

		/// <summary>
		/// SupportByVersion MSForms 2
		/// </summary>
		/// <param name="bstrProgID">string bstrProgID</param>
		/// <param name="name">optional object name</param>
		[CustomMethod]
		[SupportByVersion("MSForms", 2)]
		NetOffice.MSFormsApi.Control Add(string bstrProgID, object name);

		/// <summary>
		/// SupportByVersion MSForms 2
		/// </summary>
		/// <param name="varg">object varg</param>
		[SupportByVersion("MSForms", 2)]
		void Remove(object varg);

        #endregion

        #region IEnumerable<object>

        /// <summary>
        /// SupportByVersion MSForms, 2
        /// </summary>
        [SupportByVersion("MSForms", 2)]
        new IEnumerator<object> GetEnumerator();

        #endregion
    }
}
