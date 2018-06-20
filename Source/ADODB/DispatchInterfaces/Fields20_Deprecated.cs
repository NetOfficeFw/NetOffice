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
	/// DispatchInterface Fields20_Deprecated 
	/// SupportByVersion ADODB, 2.5
	/// </summary>
	[SupportByVersion("ADODB", 2.5)]
	[EntityType(EntityType.IsDispatchInterface), BaseType, Enumerator(Enumerator.Reference, EnumeratorInvoke.Method, "ADODB", 2.1, 2.5)]
	[TypeId("0000054D-0000-0010-8000-00AA006D2EA4")]
	public interface Fields20_Deprecated : Fields15_Deprecated, IEnumerableProvider<object>
	{
		#region Properties

		/// <summary>
		/// SupportByVersion ADODB 2.5
		/// Get
		/// </summary>
		[SupportByVersion("ADODB", 2.5)]
        new Int32 Count { get; }

		#endregion

		#region Methods

		/// <summary>
		/// SupportByVersion ADODB 2.5
		/// </summary>
		[SupportByVersion("ADODB", 2.5)]
        new void Refresh();

		/// <summary>
		/// SupportByVersion ADODB 2.5
		/// </summary>
		/// <param name="name">string name</param>
		/// <param name="type">NetOffice.ADODBApi.Enums.DataTypeEnum type</param>
		/// <param name="definedSize">optional Int32 DefinedSize = 0</param>
		/// <param name="attrib">optional NetOffice.ADODBApi.Enums.FieldAttributeEnum Attrib = -1</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		[SupportByVersion("ADODB", 2.5)]
		void _Append(string name, NetOffice.ADODBApi.Enums.DataTypeEnum type, object definedSize, object attrib);

		/// <summary>
		/// SupportByVersion ADODB 2.5
		/// </summary>
		/// <param name="name">string name</param>
		/// <param name="type">NetOffice.ADODBApi.Enums.DataTypeEnum type</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		[CustomMethod]
		[SupportByVersion("ADODB", 2.5)]
		void _Append(string name, NetOffice.ADODBApi.Enums.DataTypeEnum type);

		/// <summary>
		/// SupportByVersion ADODB 2.5
		/// </summary>
		/// <param name="name">string name</param>
		/// <param name="type">NetOffice.ADODBApi.Enums.DataTypeEnum type</param>
		/// <param name="definedSize">optional Int32 DefinedSize = 0</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		[CustomMethod]
		[SupportByVersion("ADODB", 2.5)]
		void _Append(string name, NetOffice.ADODBApi.Enums.DataTypeEnum type, object definedSize);

		/// <summary>
		/// SupportByVersion ADODB 2.5
		/// </summary>
		/// <param name="index">object index</param>
		[SupportByVersion("ADODB", 2.5)]
		void Delete(object index);

		#endregion
	}
}
