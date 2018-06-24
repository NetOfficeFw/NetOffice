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
	/// DispatchInterface Pages 
	/// SupportByVersion MSForms, 2
	/// </summary>
	[SupportByVersion("MSForms", 2)]
	[EntityType(EntityType.IsDispatchInterface), Enumerator(Enumerator.Reference, EnumeratorInvoke.Property, "MSForms", 2), HasIndexProperty(IndexInvoke.Method, "Item")]
	[TypeId("92E11A03-7358-11CE-80CB-00AA00611080")]
	public interface Pages : ICOMObject, NetOffice.CollectionsGeneric.IEnumerableProvider<object>
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
		object Enum();

		/// <summary>
		/// SupportByVersion MSForms 2
		/// </summary>
		/// <param name="bstrName">optional object bstrName</param>
		/// <param name="bstrCaption">optional object bstrCaption</param>
		/// <param name="lIndex">optional object lIndex</param>
		[SupportByVersion("MSForms", 2)]
		NetOffice.MSFormsApi.Page Add(object bstrName, object bstrCaption, object lIndex);

		/// <summary>
		/// SupportByVersion MSForms 2
		/// </summary>
		[CustomMethod]
		[SupportByVersion("MSForms", 2)]
		NetOffice.MSFormsApi.Page Add();

		/// <summary>
		/// SupportByVersion MSForms 2
		/// </summary>
		/// <param name="bstrName">optional object bstrName</param>
		[CustomMethod]
		[SupportByVersion("MSForms", 2)]
		NetOffice.MSFormsApi.Page Add(object bstrName);

		/// <summary>
		/// SupportByVersion MSForms 2
		/// </summary>
		/// <param name="bstrName">optional object bstrName</param>
		/// <param name="bstrCaption">optional object bstrCaption</param>
		[CustomMethod]
		[SupportByVersion("MSForms", 2)]
		NetOffice.MSFormsApi.Page Add(object bstrName, object bstrCaption);

		/// <summary>
		/// SupportByVersion MSForms 2
		/// </summary>
		/// <param name="clsid">Int32 clsid</param>
		/// <param name="bstrName">string bstrName</param>
		/// <param name="bstrCaption">string bstrCaption</param>
		[SupportByVersion("MSForms", 2)]
		NetOffice.MSFormsApi.Page _AddCtrl(Int32 clsid, string bstrName, string bstrCaption);

		/// <summary>
		/// SupportByVersion MSForms 2
		/// </summary>
		/// <param name="clsid">Int32 clsid</param>
		/// <param name="bstrName">string bstrName</param>
		/// <param name="bstrCaption">string bstrCaption</param>
		/// <param name="lIndex">Int32 lIndex</param>
		[SupportByVersion("MSForms", 2)]
		NetOffice.MSFormsApi.Page _InsertCtrl(Int32 clsid, string bstrName, string bstrCaption, Int32 lIndex);

		/// <summary>
		/// SupportByVersion MSForms 2
		/// </summary>
		/// <param name="lIndex">Int32 lIndex</param>
		[SupportByVersion("MSForms", 2)]
		NetOffice.MSFormsApi.Control _GetItemByIndex(Int32 lIndex);

		/// <summary>
		/// SupportByVersion MSForms 2
		/// </summary>
		/// <param name="pstrName">string pstrName</param>
		[SupportByVersion("MSForms", 2)]
		NetOffice.MSFormsApi.Control _GetItemByName(string pstrName);

		/// <summary>
		/// SupportByVersion MSForms 2
		/// </summary>
		/// <param name="varg">object varg</param>
		[SupportByVersion("MSForms", 2)]
		void Remove(object varg);

		/// <summary>
		/// SupportByVersion MSForms 2
		/// </summary>
		[SupportByVersion("MSForms", 2)]
		void Clear();

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
