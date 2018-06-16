using System;
using NetRuntimeSystem = System;
using System.ComponentModel;
using NetOffice.Attributes;

namespace NetOffice.ADODBApi
{
	/// <summary>
	/// DispatchInterface Field15 
	/// SupportByVersion ADODB, 2.1,2.5
	/// </summary>
	[SupportByVersion("ADODB", 2.1,2.5)]
	[EntityType(EntityType.IsDispatchInterface)]
	[TypeId("00000505-0000-0010-8000-00AA006D2EA4")]
	public interface Field15 : _ADO
	{
		#region Properties

		/// <summary>
		/// SupportByVersion ADODB 2.1, 2.5
		/// Get
		/// </summary>
		[SupportByVersion("ADODB", 2.1,2.5)]
		Int32 ActualSize { get; }

		/// <summary>
		/// SupportByVersion ADODB 2.1, 2.5
		/// Get
		/// </summary>
		[SupportByVersion("ADODB", 2.1,2.5)]
		Int32 Attributes { get; }

		/// <summary>
		/// SupportByVersion ADODB 2.1, 2.5
		/// Get
		/// </summary>
		[SupportByVersion("ADODB", 2.1,2.5)]
		Int32 DefinedSize { get; }

		/// <summary>
		/// SupportByVersion ADODB 2.1, 2.5
		/// Get
		/// </summary>
		[SupportByVersion("ADODB", 2.1,2.5)]
		string Name { get; }

		/// <summary>
		/// SupportByVersion ADODB 2.1, 2.5
		/// Get
		/// </summary>
		[SupportByVersion("ADODB", 2.1,2.5)]
		NetOffice.ADODBApi.Enums.DataTypeEnum Type { get; }

		/// <summary>
		/// SupportByVersion ADODB 2.1, 2.5
		/// Get/Set
		/// </summary>
		[SupportByVersion("ADODB", 2.1,2.5)]
		object Value { get; set; }

		/// <summary>
		/// SupportByVersion ADODB 2.1, 2.5
		/// Get
		/// </summary>
		[SupportByVersion("ADODB", 2.1,2.5)]
		byte Precision { get; }

		/// <summary>
		/// SupportByVersion ADODB 2.1, 2.5
		/// Get
		/// </summary>
		[SupportByVersion("ADODB", 2.1,2.5)]
		byte NumericScale { get; }

		/// <summary>
		/// SupportByVersion ADODB 2.1, 2.5
		/// Get
		/// </summary>
		[SupportByVersion("ADODB", 2.1,2.5)]
		object OriginalValue { get; }

		/// <summary>
		/// SupportByVersion ADODB 2.1, 2.5
		/// Get
		/// </summary>
		[SupportByVersion("ADODB", 2.1,2.5)]
		object UnderlyingValue { get; }

		#endregion

		#region Methods

		/// <summary>
		/// SupportByVersion ADODB 2.1, 2.5
		/// </summary>
		/// <param name="data">object data</param>
		[SupportByVersion("ADODB", 2.1,2.5)]
		void AppendChunk(object data);

		/// <summary>
		/// SupportByVersion ADODB 2.1, 2.5
		/// </summary>
		/// <param name="length">Int32 length</param>
		[SupportByVersion("ADODB", 2.1,2.5)]
		object GetChunk(Int32 length);

		#endregion
	}
}
