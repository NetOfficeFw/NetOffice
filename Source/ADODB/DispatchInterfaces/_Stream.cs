using System;
using NetRuntimeSystem = System;
using System.ComponentModel;
using NetOffice.Attributes;

namespace NetOffice.ADODBApi
{
	/// <summary>
	/// DispatchInterface _Stream 
	/// SupportByVersion ADODB, 2.5
	/// </summary>
	[SupportByVersion("ADODB", 2.5)]
	[EntityType(EntityType.IsDispatchInterface), BaseType]
	[TypeId("00000565-0000-0010-8000-00AA006D2EA4")]
    [CoClassSource(typeof(NetOffice.ADODBApi.Stream))]
    public interface _Stream : ICOMObject
	{
		#region Properties

		/// <summary>
		/// SupportByVersion ADODB 2.5
		/// Get
		/// </summary>
		[SupportByVersion("ADODB", 2.5)]
		Int32 Size { get; }

		/// <summary>
		/// SupportByVersion ADODB 2.5
		/// Get
		/// </summary>
		[SupportByVersion("ADODB", 2.5)]
		bool EOS { get; }

		/// <summary>
		/// SupportByVersion ADODB 2.5
		/// Get/Set
		/// </summary>
		[SupportByVersion("ADODB", 2.5)]
		Int32 Position { get; set; }

		/// <summary>
		/// SupportByVersion ADODB 2.5
		/// Get/Set
		/// </summary>
		[SupportByVersion("ADODB", 2.5)]
		NetOffice.ADODBApi.Enums.StreamTypeEnum Type { get; set; }

		/// <summary>
		/// SupportByVersion ADODB 2.5
		/// Get/Set
		/// </summary>
		[SupportByVersion("ADODB", 2.5)]
		NetOffice.ADODBApi.Enums.LineSeparatorEnum LineSeparator { get; set; }

		/// <summary>
		/// SupportByVersion ADODB 2.5
		/// Get
		/// </summary>
		[SupportByVersion("ADODB", 2.5)]
		NetOffice.ADODBApi.Enums.ObjectStateEnum State { get; }

		/// <summary>
		/// SupportByVersion ADODB 2.5
		/// Get/Set
		/// </summary>
		[SupportByVersion("ADODB", 2.5)]
		NetOffice.ADODBApi.Enums.ConnectModeEnum Mode { get; set; }

		/// <summary>
		/// SupportByVersion ADODB 2.5
		/// Get/Set
		/// </summary>
		[SupportByVersion("ADODB", 2.5)]
		string Charset { get; set; }

		#endregion

		#region Methods

		/// <summary>
		/// SupportByVersion ADODB 2.5
		/// </summary>
		/// <param name="numBytes">optional Int32 NumBytes = -1</param>
		[SupportByVersion("ADODB", 2.5)]
		object Read(object numBytes);

		/// <summary>
		/// SupportByVersion ADODB 2.5
		/// </summary>
		[CustomMethod]
		[SupportByVersion("ADODB", 2.5)]
		object Read();

		/// <summary>
		/// SupportByVersion ADODB 2.5
		/// </summary>
		/// <param name="source">optional object source</param>
		/// <param name="mode">optional NetOffice.ADODBApi.Enums.ConnectModeEnum Mode = 0</param>
		/// <param name="options">optional NetOffice.ADODBApi.Enums.StreamOpenOptionsEnum Options = -1</param>
		/// <param name="userName">optional string userName</param>
		/// <param name="password">optional string password</param>
		[SupportByVersion("ADODB", 2.5)]
		void Open(object source, object mode, object options, object userName, object password);

		/// <summary>
		/// SupportByVersion ADODB 2.5
		/// </summary>
		[CustomMethod]
		[SupportByVersion("ADODB", 2.5)]
		void Open();

		/// <summary>
		/// SupportByVersion ADODB 2.5
		/// </summary>
		/// <param name="source">optional object source</param>
		[CustomMethod]
		[SupportByVersion("ADODB", 2.5)]
		void Open(object source);

		/// <summary>
		/// SupportByVersion ADODB 2.5
		/// </summary>
		/// <param name="source">optional object source</param>
		/// <param name="mode">optional NetOffice.ADODBApi.Enums.ConnectModeEnum Mode = 0</param>
		[CustomMethod]
		[SupportByVersion("ADODB", 2.5)]
		void Open(object source, object mode);

		/// <summary>
		/// SupportByVersion ADODB 2.5
		/// </summary>
		/// <param name="source">optional object source</param>
		/// <param name="mode">optional NetOffice.ADODBApi.Enums.ConnectModeEnum Mode = 0</param>
		/// <param name="options">optional NetOffice.ADODBApi.Enums.StreamOpenOptionsEnum Options = -1</param>
		[CustomMethod]
		[SupportByVersion("ADODB", 2.5)]
		void Open(object source, object mode, object options);

		/// <summary>
		/// SupportByVersion ADODB 2.5
		/// </summary>
		/// <param name="source">optional object source</param>
		/// <param name="mode">optional NetOffice.ADODBApi.Enums.ConnectModeEnum Mode = 0</param>
		/// <param name="options">optional NetOffice.ADODBApi.Enums.StreamOpenOptionsEnum Options = -1</param>
		/// <param name="userName">optional string userName</param>
		[CustomMethod]
		[SupportByVersion("ADODB", 2.5)]
		void Open(object source, object mode, object options, object userName);

		/// <summary>
		/// SupportByVersion ADODB 2.5
		/// </summary>
		[SupportByVersion("ADODB", 2.5)]
		void Close();

		/// <summary>
		/// SupportByVersion ADODB 2.5
		/// </summary>
		[SupportByVersion("ADODB", 2.5)]
		void SkipLine();

		/// <summary>
		/// SupportByVersion ADODB 2.5
		/// </summary>
		/// <param name="buffer">object buffer</param>
		[SupportByVersion("ADODB", 2.5)]
		void Write(object buffer);

		/// <summary>
		/// SupportByVersion ADODB 2.5
		/// </summary>
		[SupportByVersion("ADODB", 2.5)]
		void SetEOS();

		/// <summary>
		/// SupportByVersion ADODB 2.5
		/// </summary>
		/// <param name="destStream">NetOffice.ADODBApi._Stream destStream</param>
		/// <param name="charNumber">optional Int32 CharNumber = -1</param>
		[SupportByVersion("ADODB", 2.5)]
		void CopyTo(NetOffice.ADODBApi._Stream destStream, object charNumber);

		/// <summary>
		/// SupportByVersion ADODB 2.5
		/// </summary>
		/// <param name="destStream">NetOffice.ADODBApi._Stream destStream</param>
		[CustomMethod]
		[SupportByVersion("ADODB", 2.5)]
		void CopyTo(NetOffice.ADODBApi._Stream destStream);

		/// <summary>
		/// SupportByVersion ADODB 2.5
		/// </summary>
		[SupportByVersion("ADODB", 2.5)]
		void Flush();

		/// <summary>
		/// SupportByVersion ADODB 2.5
		/// </summary>
		/// <param name="fileName">string fileName</param>
		/// <param name="options">optional NetOffice.ADODBApi.Enums.SaveOptionsEnum Options = 1</param>
		[SupportByVersion("ADODB", 2.5)]
		void SaveToFile(string fileName, object options);

		/// <summary>
		/// SupportByVersion ADODB 2.5
		/// </summary>
		/// <param name="fileName">string fileName</param>
		[CustomMethod]
		[SupportByVersion("ADODB", 2.5)]
		void SaveToFile(string fileName);

		/// <summary>
		/// SupportByVersion ADODB 2.5
		/// </summary>
		/// <param name="fileName">string fileName</param>
		[SupportByVersion("ADODB", 2.5)]
		void LoadFromFile(string fileName);

		/// <summary>
		/// SupportByVersion ADODB 2.5
		/// </summary>
		/// <param name="numChars">optional Int32 NumChars = -1</param>
		[SupportByVersion("ADODB", 2.5)]
		string ReadText(object numChars);

		/// <summary>
		/// SupportByVersion ADODB 2.5
		/// </summary>
		[CustomMethod]
		[SupportByVersion("ADODB", 2.5)]
		string ReadText();

		/// <summary>
		/// SupportByVersion ADODB 2.5
		/// </summary>
		/// <param name="data">string data</param>
		/// <param name="options">optional NetOffice.ADODBApi.Enums.StreamWriteEnum Options = 0</param>
		[SupportByVersion("ADODB", 2.5)]
		void WriteText(string data, object options);

		/// <summary>
		/// SupportByVersion ADODB 2.5
		/// </summary>
		/// <param name="data">string data</param>
		[CustomMethod]
		[SupportByVersion("ADODB", 2.5)]
		void WriteText(string data);

		/// <summary>
		/// SupportByVersion ADODB 2.5
		/// </summary>
		[SupportByVersion("ADODB", 2.5)]
		void Cancel();

		#endregion
	}
}
