using System;
using NetRuntimeSystem = System;
using System.ComponentModel;
using NetOffice.Attributes;

namespace NetOffice.ADODBApi
{
	/// <summary>
	/// DispatchInterface Field 
	/// SupportByVersion ADODB, 2.1,2.5
	/// </summary>
	[SupportByVersion("ADODB", 2.1,2.5)]
	[EntityType(EntityType.IsDispatchInterface)]
	[TypeId("0000054C-0000-0010-8000-00AA006D2EA4")]
	public interface Field : Field20
	{
		#region Properties

		/// <summary>
		/// SupportByVersion ADODB 2.1
		/// Get
		/// </summary>
		[SupportByVersion("ADODB", 2.1)]
        new Int32 ActualSize { get; }

		/// <summary>
		/// SupportByVersion ADODB 2.1
		/// Get/Set
		/// </summary>
		[SupportByVersion("ADODB", 2.1)]
        new Int32 Attributes { get; set; }

		/// <summary>
		/// SupportByVersion ADODB 2.1
		/// Get/Set
		/// </summary>
		[SupportByVersion("ADODB", 2.1)]
        new Int32 DefinedSize { get; set; }

		/// <summary>
		/// SupportByVersion ADODB 2.1
		/// Get
		/// </summary>
		[SupportByVersion("ADODB", 2.1)]
        new string Name { get; }

		/// <summary>
		/// SupportByVersion ADODB 2.1
		/// Get/Set
		/// </summary>
		[SupportByVersion("ADODB", 2.1)]
        new NetOffice.ADODBApi.Enums.DataTypeEnum Type { get; set; }

		/// <summary>
		/// SupportByVersion ADODB 2.1
		/// Get/Set
		/// </summary>
		[SupportByVersion("ADODB", 2.1)]
        new object Value { get; set; }

		/// <summary>
		/// SupportByVersion ADODB 2.1
		/// Get/Set
		/// </summary>
		[SupportByVersion("ADODB", 2.1)]
        new byte Precision { get; set; }

		/// <summary>
		/// SupportByVersion ADODB 2.1
		/// Get/Set
		/// </summary>
		[SupportByVersion("ADODB", 2.1)]
        new byte NumericScale { get; set; }

		/// <summary>
		/// SupportByVersion ADODB 2.1
		/// Get
		/// </summary>
		[SupportByVersion("ADODB", 2.1)]
        new object OriginalValue { get; }

		/// <summary>
		/// SupportByVersion ADODB 2.1
		/// Get
		/// </summary>
		[SupportByVersion("ADODB", 2.1)]
        new object UnderlyingValue { get; }

		/// <summary>
		/// SupportByVersion ADODB 2.1
		/// Get/Set
		/// Unknown COM Proxy
		/// </summary>
		[SupportByVersion("ADODB", 2.1), ProxyResult]
        new object DataFormat { get; set; }

		/// <summary>
		/// SupportByVersion ADODB 2.5
		/// Get
		/// </summary>
		[SupportByVersion("ADODB", 2.5)]
        new NetOffice.ADODBApi.Properties Properties { get; }

		/// <summary>
		/// SupportByVersion ADODB 2.5
		/// Get
		/// </summary>
		[SupportByVersion("ADODB", 2.5)]
		Int32 Status { get; }

		#endregion

		#region Methods

		/// <summary>
		/// SupportByVersion ADODB 2.1
		/// </summary>
		/// <param name="data">object data</param>
		[SupportByVersion("ADODB", 2.1)]
        new void AppendChunk(object data);

		/// <summary>
		/// SupportByVersion ADODB 2.1
		/// </summary>
		/// <param name="length">Int32 length</param>
		[SupportByVersion("ADODB", 2.1)]
        new object GetChunk(Int32 length);

		#endregion
	}
}
