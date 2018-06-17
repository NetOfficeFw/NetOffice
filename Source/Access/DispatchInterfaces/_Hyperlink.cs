using System;
using NetRuntimeSystem = System;
using System.ComponentModel;
using NetOffice.Attributes;

namespace NetOffice.AccessApi
{
	/// <summary>
	/// DispatchInterface _Hyperlink 
	/// SupportByVersion Access, 9,10,11,12,14,15,16
	/// </summary>
	[SupportByVersion("Access", 9,10,11,12,14,15,16)]
	[EntityType(EntityType.IsDispatchInterface), BaseType]
	[TypeId("50D56611-60AC-11CF-82C9-00AA004B9FE6")]
    [CoClassSource(typeof(NetOffice.AccessApi.Hyperlink))]
    public interface _Hyperlink : ICOMObject
	{
		#region Properties

		/// <summary>
		/// SupportByVersion Access 9, 10, 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff822025.aspx </remarks>
		[SupportByVersion("Access", 9,10,11,12,14,15,16)]
		string SubAddress { get; set; }

		/// <summary>
		/// SupportByVersion Access 9, 10, 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff195258.aspx </remarks>
		[SupportByVersion("Access", 9,10,11,12,14,15,16)]
		string Address { get; set; }

		/// <summary>
		/// SupportByVersion Access 9, 10, 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff835717.aspx </remarks>
		[SupportByVersion("Access", 9,10,11,12,14,15,16)]
		string EmailSubject { get; set; }

		/// <summary>
		/// SupportByVersion Access 9, 10, 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff822461.aspx </remarks>
		[SupportByVersion("Access", 9,10,11,12,14,15,16)]
		string ScreenTip { get; set; }

		/// <summary>
		/// SupportByVersion Access 9, 10, 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff194871.aspx </remarks>
		[SupportByVersion("Access", 9,10,11,12,14,15,16)]
		string TextToDisplay { get; set; }

		#endregion

		#region Methods

		/// <summary>
		/// SupportByVersion Access 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff192921.aspx </remarks>
		[SupportByVersion("Access", 9,10,11,12,14,15,16)]
		void AddToFavorites();

		/// <summary>
		/// SupportByVersion Access 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff196761.aspx </remarks>
		/// <param name="newWindow">optional bool NewWindow = false</param>
		/// <param name="addHistory">optional bool AddHistory = true</param>
		/// <param name="extraInfo">optional object extraInfo</param>
		/// <param name="method">optional NetOffice.OfficeApi.Enums.MsoExtraInfoMethod Method = 0</param>
		/// <param name="headerInfo">optional string HeaderInfo = </param>
		[SupportByVersion("Access", 9,10,11,12,14,15,16)]
		void Follow(object newWindow, object addHistory, object extraInfo, object method, object headerInfo);

		/// <summary>
		/// SupportByVersion Access 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff196761.aspx </remarks>
		[CustomMethod]
		[SupportByVersion("Access", 9,10,11,12,14,15,16)]
		void Follow();

		/// <summary>
		/// SupportByVersion Access 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff196761.aspx </remarks>
		/// <param name="newWindow">optional bool NewWindow = false</param>
		[CustomMethod]
		[SupportByVersion("Access", 9,10,11,12,14,15,16)]
		void Follow(object newWindow);

		/// <summary>
		/// SupportByVersion Access 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff196761.aspx </remarks>
		/// <param name="newWindow">optional bool NewWindow = false</param>
		/// <param name="addHistory">optional bool AddHistory = true</param>
		[CustomMethod]
		[SupportByVersion("Access", 9,10,11,12,14,15,16)]
		void Follow(object newWindow, object addHistory);

		/// <summary>
		/// SupportByVersion Access 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff196761.aspx </remarks>
		/// <param name="newWindow">optional bool NewWindow = false</param>
		/// <param name="addHistory">optional bool AddHistory = true</param>
		/// <param name="extraInfo">optional object extraInfo</param>
		[CustomMethod]
		[SupportByVersion("Access", 9,10,11,12,14,15,16)]
		void Follow(object newWindow, object addHistory, object extraInfo);

		/// <summary>
		/// SupportByVersion Access 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff196761.aspx </remarks>
		/// <param name="newWindow">optional bool NewWindow = false</param>
		/// <param name="addHistory">optional bool AddHistory = true</param>
		/// <param name="extraInfo">optional object extraInfo</param>
		/// <param name="method">optional NetOffice.OfficeApi.Enums.MsoExtraInfoMethod Method = 0</param>
		[CustomMethod]
		[SupportByVersion("Access", 9,10,11,12,14,15,16)]
		void Follow(object newWindow, object addHistory, object extraInfo, object method);

		/// <summary>
		/// SupportByVersion Access 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff822725.aspx </remarks>
		/// <param name="fileName">string fileName</param>
		/// <param name="editNow">bool editNow</param>
		/// <param name="overwrite">bool overwrite</param>
		[SupportByVersion("Access", 9,10,11,12,14,15,16)]
		void CreateNewDocument(string fileName, bool editNow, bool overwrite);

		/// <summary>
		/// SupportByVersion Access 11, 12, 14, 15, 16
		/// </summary>
		/// <param name="dispid">Int32 dispid</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		[SupportByVersion("Access", 11,12,14,15,16)]
		bool IsMemberSafe(Int32 dispid);

		#endregion
	}
}
