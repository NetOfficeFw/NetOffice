using System.Collections;
using System.Collections.Generic;
using System;
using NetRuntimeSystem = System;
using System.ComponentModel;
using NetOffice.Attributes;
using NetOffice.CollectionsGeneric;

namespace NetOffice.ADODBApi
{
	/// <summary>
	/// DispatchInterface Fields 
	/// SupportByVersion ADODB, 2.1,2.5
	/// </summary>
	[SupportByVersion("ADODB", 2.1,2.5)]
	[EntityType(EntityType.IsDispatchInterface), Enumerator(Enumerator.Reference, EnumeratorInvoke.Method, "ADODB", 2.1, 2.5), HasIndexProperty(IndexInvoke.Property, "Item")]
	[TypeId("0000054D-0000-0010-8000-00AA006D2EA4")]
	public interface Fields : Fields20, IEnumerableProvider<NetOffice.ADODBApi.Field>
	{
		#region Properties

		/// <summary>
		/// SupportByVersion ADODB 2.1, 2.5
		/// Get
		/// </summary>
		[SupportByVersion("ADODB", 2.1,2.5)]
        new Int32 Count { get; }

		/// <summary>
		/// SupportByVersion ADODB 2.5
		/// Get
		/// </summary>
		/// <param name="index">object index</param>
		[SupportByVersion("ADODB", 2.5)]
		[NetRuntimeSystem.Runtime.CompilerServices.IndexerName("Item"), IndexProperty]
        new NetOffice.ADODBApi.Field this[object index] { get; }

		#endregion

		#region Methods

		/// <summary>
		/// SupportByVersion ADODB 2.1, 2.5
		/// </summary>
		[SupportByVersion("ADODB", 2.1,2.5)]
        new void Refresh();

		/// <summary>
		/// SupportByVersion ADODB 2.1, 2.5
		/// </summary>
		/// <param name="name">string name</param>
		/// <param name="type">NetOffice.ADODBApi.Enums.DataTypeEnum type</param>
		/// <param name="definedSize">optional Int32 DefinedSize = 0</param>
		/// <param name="attrib">optional NetOffice.ADODBApi.Enums.FieldAttributeEnum Attrib = -1</param>
		[SupportByVersion("ADODB", 2.1,2.5)]
		void Append(string name, NetOffice.ADODBApi.Enums.DataTypeEnum type, object definedSize, object attrib);

		/// <summary>
		/// SupportByVersion ADODB 2.5
		/// </summary>
		/// <param name="name">string name</param>
		/// <param name="type">NetOffice.ADODBApi.Enums.DataTypeEnum type</param>
		/// <param name="definedSize">optional Int32 DefinedSize = 0</param>
		/// <param name="attrib">optional NetOffice.ADODBApi.Enums.FieldAttributeEnum Attrib = -1</param>
		/// <param name="fieldValue">optional object fieldValue</param>
		[SupportByVersion("ADODB", 2.5)]
		void Append(string name, NetOffice.ADODBApi.Enums.DataTypeEnum type, object definedSize, object attrib, object fieldValue);

		/// <summary>
		/// SupportByVersion ADODB 2.1, 2.5
		/// </summary>
		/// <param name="name">string name</param>
		/// <param name="type">NetOffice.ADODBApi.Enums.DataTypeEnum type</param>
		[CustomMethod]
		[SupportByVersion("ADODB", 2.1,2.5)]
		void Append(string name, NetOffice.ADODBApi.Enums.DataTypeEnum type);

		/// <summary>
		/// SupportByVersion ADODB 2.1, 2.5
		/// </summary>
		/// <param name="name">string name</param>
		/// <param name="type">NetOffice.ADODBApi.Enums.DataTypeEnum type</param>
		/// <param name="definedSize">optional Int32 DefinedSize = 0</param>
		[CustomMethod]
		[SupportByVersion("ADODB", 2.1,2.5)]
		void Append(string name, NetOffice.ADODBApi.Enums.DataTypeEnum type, object definedSize);

		/// <summary>
		/// SupportByVersion ADODB 2.1
		/// </summary>
		/// <param name="index">object index</param>
		[SupportByVersion("ADODB", 2.1)]
        new void Delete(object index);

		/// <summary>
		/// SupportByVersion ADODB 2.5
		/// </summary>
		[SupportByVersion("ADODB", 2.5)]
		void Update();

		/// <summary>
		/// SupportByVersion ADODB 2.5
		/// </summary>
		/// <param name="resyncValues">optional NetOffice.ADODBApi.Enums.ResyncEnum ResyncValues = 2</param>
		[SupportByVersion("ADODB", 2.5)]
		void Resync(object resyncValues);

		/// <summary>
		/// SupportByVersion ADODB 2.5
		/// </summary>
		[CustomMethod]
		[SupportByVersion("ADODB", 2.5)]
		void Resync();

		/// <summary>
		/// SupportByVersion ADODB 2.5
		/// </summary>
		[SupportByVersion("ADODB", 2.5)]
		void CancelUpdate();

        #endregion

        #region IEnumerable<NetOffice.ADODBApi.Field>

        /// <summary>
        /// SupportByVersion ADODB, 2.1,2.5
        /// </summary>
        [SupportByVersion("ADODB", 2.1, 2.5)]
        new IEnumerator<NetOffice.ADODBApi.Field> GetEnumerator();

        #endregion
    }
}
