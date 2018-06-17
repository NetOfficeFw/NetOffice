using System;
using NetRuntimeSystem = System;
using System.ComponentModel;
using NetOffice.Attributes;

namespace NetOffice.AccessApi
{
	/// <summary>
	/// DispatchInterface DoCmd 
	/// SupportByVersion Access, 9,10,11,12,14,15,16
	/// </summary>
	/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff192694.aspx </remarks>
	[SupportByVersion("Access", 9,10,11,12,14,15,16)]
	[EntityType(EntityType.IsDispatchInterface)]
	[TypeId("C547E760-9658-101B-81EE-00AA004750E2")]
	public interface DoCmd : ICOMObject
	{
		#region Methods

		/// <summary>
		/// SupportByVersion Access 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff834781.aspx </remarks>
		/// <param name="menuName">object menuName</param>
		/// <param name="menuMacroName">object menuMacroName</param>
		/// <param name="statusBarText">object statusBarText</param>
		[SupportByVersion("Access", 9,10,11,12,14,15,16)]
		void AddMenu(object menuName, object menuMacroName, object statusBarText);

		/// <summary>
		/// SupportByVersion Access 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff197651.aspx </remarks>
		/// <param name="filterName">optional object filterName</param>
		/// <param name="whereCondition">optional object whereCondition</param>
		[SupportByVersion("Access", 9,10,11,12,14,15,16)]
		void ApplyFilter(object filterName, object whereCondition);

		/// <summary>
		/// SupportByVersion Access 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff197651.aspx </remarks>
		/// <param name="filterName">optional object filterName</param>
		/// <param name="whereCondition">optional object whereCondition</param>
		/// <param name="controlName">optional object controlName</param>
		[SupportByVersion("Access", 12,14,15,16)]
		void ApplyFilter(object filterName, object whereCondition, object controlName);

		/// <summary>
		/// SupportByVersion Access 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff197651.aspx </remarks>
		[CustomMethod]
		[SupportByVersion("Access", 9,10,11,12,14,15,16)]
		void ApplyFilter();

		/// <summary>
		/// SupportByVersion Access 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff197651.aspx </remarks>
		/// <param name="filterName">optional object filterName</param>
		[CustomMethod]
		[SupportByVersion("Access", 9,10,11,12,14,15,16)]
		void ApplyFilter(object filterName);

		/// <summary>
		/// SupportByVersion Access 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff196680.aspx </remarks>
		[SupportByVersion("Access", 9,10,11,12,14,15,16)]
		void Beep();

		/// <summary>
		/// SupportByVersion Access 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff836964.aspx </remarks>
		[SupportByVersion("Access", 9,10,11,12,14,15,16)]
		void CancelEvent();

		/// <summary>
		/// SupportByVersion Access 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff192860.aspx </remarks>
		/// <param name="objectType">optional NetOffice.AccessApi.Enums.AcObjectType ObjectType = -1</param>
		/// <param name="objectName">optional object objectName</param>
		/// <param name="save">optional NetOffice.AccessApi.Enums.AcCloseSave Save = 0</param>
		[SupportByVersion("Access", 9,10,11,12,14,15,16)]
		void Close(object objectType, object objectName, object save);

		/// <summary>
		/// SupportByVersion Access 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff192860.aspx </remarks>
		[CustomMethod]
		[SupportByVersion("Access", 9,10,11,12,14,15,16)]
		void Close();

		/// <summary>
		/// SupportByVersion Access 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff192860.aspx </remarks>
		/// <param name="objectType">optional NetOffice.AccessApi.Enums.AcObjectType ObjectType = -1</param>
		[CustomMethod]
		[SupportByVersion("Access", 9,10,11,12,14,15,16)]
		void Close(object objectType);

		/// <summary>
		/// SupportByVersion Access 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff192860.aspx </remarks>
		/// <param name="objectType">optional NetOffice.AccessApi.Enums.AcObjectType ObjectType = -1</param>
		/// <param name="objectName">optional object objectName</param>
		[CustomMethod]
		[SupportByVersion("Access", 9,10,11,12,14,15,16)]
		void Close(object objectType, object objectName);

		/// <summary>
		/// SupportByVersion Access 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff844724.aspx </remarks>
		/// <param name="destinationDatabase">optional object destinationDatabase</param>
		/// <param name="newName">optional object newName</param>
		/// <param name="sourceObjectType">optional NetOffice.AccessApi.Enums.AcObjectType SourceObjectType = -1</param>
		/// <param name="sourceObjectName">optional object sourceObjectName</param>
		[SupportByVersion("Access", 9,10,11,12,14,15,16)]
		void CopyObject(object destinationDatabase, object newName, object sourceObjectType, object sourceObjectName);

		/// <summary>
		/// SupportByVersion Access 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff844724.aspx </remarks>
		[CustomMethod]
		[SupportByVersion("Access", 9,10,11,12,14,15,16)]
		void CopyObject();

		/// <summary>
		/// SupportByVersion Access 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff844724.aspx </remarks>
		/// <param name="destinationDatabase">optional object destinationDatabase</param>
		[CustomMethod]
		[SupportByVersion("Access", 9,10,11,12,14,15,16)]
		void CopyObject(object destinationDatabase);

		/// <summary>
		/// SupportByVersion Access 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff844724.aspx </remarks>
		/// <param name="destinationDatabase">optional object destinationDatabase</param>
		/// <param name="newName">optional object newName</param>
		[CustomMethod]
		[SupportByVersion("Access", 9,10,11,12,14,15,16)]
		void CopyObject(object destinationDatabase, object newName);

		/// <summary>
		/// SupportByVersion Access 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff844724.aspx </remarks>
		/// <param name="destinationDatabase">optional object destinationDatabase</param>
		/// <param name="newName">optional object newName</param>
		/// <param name="sourceObjectType">optional NetOffice.AccessApi.Enums.AcObjectType SourceObjectType = -1</param>
		[CustomMethod]
		[SupportByVersion("Access", 9,10,11,12,14,15,16)]
		void CopyObject(object destinationDatabase, object newName, object sourceObjectType);

		/// <summary>
		/// SupportByVersion Access 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff822447.aspx </remarks>
		/// <param name="menuBar">object menuBar</param>
		/// <param name="menuName">object menuName</param>
		/// <param name="command">object command</param>
		/// <param name="subcommand">optional object subcommand</param>
		/// <param name="version">optional object version</param>
		[SupportByVersion("Access", 9,10,11,12,14,15,16)]
		void DoMenuItem(object menuBar, object menuName, object command, object subcommand, object version);

		/// <summary>
		/// SupportByVersion Access 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff822447.aspx </remarks>
		/// <param name="menuBar">object menuBar</param>
		/// <param name="menuName">object menuName</param>
		/// <param name="command">object command</param>
		[CustomMethod]
		[SupportByVersion("Access", 9,10,11,12,14,15,16)]
		void DoMenuItem(object menuBar, object menuName, object command);

		/// <summary>
		/// SupportByVersion Access 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff822447.aspx </remarks>
		/// <param name="menuBar">object menuBar</param>
		/// <param name="menuName">object menuName</param>
		/// <param name="command">object command</param>
		/// <param name="subcommand">optional object subcommand</param>
		[CustomMethod]
		[SupportByVersion("Access", 9,10,11,12,14,15,16)]
		void DoMenuItem(object menuBar, object menuName, object command, object subcommand);

		/// <summary>
		/// SupportByVersion Access 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff193863.aspx </remarks>
		/// <param name="echoOn">object echoOn</param>
		/// <param name="statusBarText">optional object statusBarText</param>
		[SupportByVersion("Access", 9,10,11,12,14,15,16)]
		void Echo(object echoOn, object statusBarText);

		/// <summary>
		/// SupportByVersion Access 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff193863.aspx </remarks>
		/// <param name="echoOn">object echoOn</param>
		[CustomMethod]
		[SupportByVersion("Access", 9,10,11,12,14,15,16)]
		void Echo(object echoOn);

		/// <summary>
		/// SupportByVersion Access 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff196453.aspx </remarks>
		[SupportByVersion("Access", 9,10,11,12,14,15,16)]
		void FindNext();

		/// <summary>
		/// SupportByVersion Access 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff835361.aspx </remarks>
		/// <param name="findWhat">object findWhat</param>
		/// <param name="match">optional NetOffice.AccessApi.Enums.AcFindMatch Match = 1</param>
		/// <param name="matchCase">optional object matchCase</param>
		/// <param name="search">optional NetOffice.AccessApi.Enums.AcSearchDirection Search = 2</param>
		/// <param name="searchAsFormatted">optional object searchAsFormatted</param>
		/// <param name="onlyCurrentField">optional NetOffice.AccessApi.Enums.AcFindField OnlyCurrentField = -1</param>
		/// <param name="findFirst">optional object findFirst</param>
		[SupportByVersion("Access", 9,10,11,12,14,15,16)]
		void FindRecord(object findWhat, object match, object matchCase, object search, object searchAsFormatted, object onlyCurrentField, object findFirst);

		/// <summary>
		/// SupportByVersion Access 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff835361.aspx </remarks>
		/// <param name="findWhat">object findWhat</param>
		[CustomMethod]
		[SupportByVersion("Access", 9,10,11,12,14,15,16)]
		void FindRecord(object findWhat);

		/// <summary>
		/// SupportByVersion Access 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff835361.aspx </remarks>
		/// <param name="findWhat">object findWhat</param>
		/// <param name="match">optional NetOffice.AccessApi.Enums.AcFindMatch Match = 1</param>
		[CustomMethod]
		[SupportByVersion("Access", 9,10,11,12,14,15,16)]
		void FindRecord(object findWhat, object match);

		/// <summary>
		/// SupportByVersion Access 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff835361.aspx </remarks>
		/// <param name="findWhat">object findWhat</param>
		/// <param name="match">optional NetOffice.AccessApi.Enums.AcFindMatch Match = 1</param>
		/// <param name="matchCase">optional object matchCase</param>
		[CustomMethod]
		[SupportByVersion("Access", 9,10,11,12,14,15,16)]
		void FindRecord(object findWhat, object match, object matchCase);

		/// <summary>
		/// SupportByVersion Access 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff835361.aspx </remarks>
		/// <param name="findWhat">object findWhat</param>
		/// <param name="match">optional NetOffice.AccessApi.Enums.AcFindMatch Match = 1</param>
		/// <param name="matchCase">optional object matchCase</param>
		/// <param name="search">optional NetOffice.AccessApi.Enums.AcSearchDirection Search = 2</param>
		[CustomMethod]
		[SupportByVersion("Access", 9,10,11,12,14,15,16)]
		void FindRecord(object findWhat, object match, object matchCase, object search);

		/// <summary>
		/// SupportByVersion Access 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff835361.aspx </remarks>
		/// <param name="findWhat">object findWhat</param>
		/// <param name="match">optional NetOffice.AccessApi.Enums.AcFindMatch Match = 1</param>
		/// <param name="matchCase">optional object matchCase</param>
		/// <param name="search">optional NetOffice.AccessApi.Enums.AcSearchDirection Search = 2</param>
		/// <param name="searchAsFormatted">optional object searchAsFormatted</param>
		[CustomMethod]
		[SupportByVersion("Access", 9,10,11,12,14,15,16)]
		void FindRecord(object findWhat, object match, object matchCase, object search, object searchAsFormatted);

		/// <summary>
		/// SupportByVersion Access 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff835361.aspx </remarks>
		/// <param name="findWhat">object findWhat</param>
		/// <param name="match">optional NetOffice.AccessApi.Enums.AcFindMatch Match = 1</param>
		/// <param name="matchCase">optional object matchCase</param>
		/// <param name="search">optional NetOffice.AccessApi.Enums.AcSearchDirection Search = 2</param>
		/// <param name="searchAsFormatted">optional object searchAsFormatted</param>
		/// <param name="onlyCurrentField">optional NetOffice.AccessApi.Enums.AcFindField OnlyCurrentField = -1</param>
		[CustomMethod]
		[SupportByVersion("Access", 9,10,11,12,14,15,16)]
		void FindRecord(object findWhat, object match, object matchCase, object search, object searchAsFormatted, object onlyCurrentField);

		/// <summary>
		/// SupportByVersion Access 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff192079.aspx </remarks>
		/// <param name="controlName">object controlName</param>
		[SupportByVersion("Access", 9,10,11,12,14,15,16)]
		void GoToControl(object controlName);

		/// <summary>
		/// SupportByVersion Access 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff192504.aspx </remarks>
		/// <param name="pageNumber">object pageNumber</param>
		/// <param name="right">optional object right</param>
		/// <param name="down">optional object down</param>
		[SupportByVersion("Access", 9,10,11,12,14,15,16)]
		void GoToPage(object pageNumber, object right, object down);

		/// <summary>
		/// SupportByVersion Access 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff192504.aspx </remarks>
		/// <param name="pageNumber">object pageNumber</param>
		[CustomMethod]
		[SupportByVersion("Access", 9,10,11,12,14,15,16)]
		void GoToPage(object pageNumber);

		/// <summary>
		/// SupportByVersion Access 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff192504.aspx </remarks>
		/// <param name="pageNumber">object pageNumber</param>
		/// <param name="right">optional object right</param>
		[CustomMethod]
		[SupportByVersion("Access", 9,10,11,12,14,15,16)]
		void GoToPage(object pageNumber, object right);

		/// <summary>
		/// SupportByVersion Access 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff194117.aspx </remarks>
		/// <param name="objectType">optional NetOffice.AccessApi.Enums.AcDataObjectType ObjectType = -1</param>
		/// <param name="objectName">optional object objectName</param>
		/// <param name="record">optional NetOffice.AccessApi.Enums.AcRecord Record = 1</param>
		/// <param name="offset">optional object offset</param>
		[SupportByVersion("Access", 9,10,11,12,14,15,16)]
		void GoToRecord(object objectType, object objectName, object record, object offset);

		/// <summary>
		/// SupportByVersion Access 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff194117.aspx </remarks>
		[CustomMethod]
		[SupportByVersion("Access", 9,10,11,12,14,15,16)]
		void GoToRecord();

		/// <summary>
		/// SupportByVersion Access 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff194117.aspx </remarks>
		/// <param name="objectType">optional NetOffice.AccessApi.Enums.AcDataObjectType ObjectType = -1</param>
		[CustomMethod]
		[SupportByVersion("Access", 9,10,11,12,14,15,16)]
		void GoToRecord(object objectType);

		/// <summary>
		/// SupportByVersion Access 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff194117.aspx </remarks>
		/// <param name="objectType">optional NetOffice.AccessApi.Enums.AcDataObjectType ObjectType = -1</param>
		/// <param name="objectName">optional object objectName</param>
		[CustomMethod]
		[SupportByVersion("Access", 9,10,11,12,14,15,16)]
		void GoToRecord(object objectType, object objectName);

		/// <summary>
		/// SupportByVersion Access 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff194117.aspx </remarks>
		/// <param name="objectType">optional NetOffice.AccessApi.Enums.AcDataObjectType ObjectType = -1</param>
		/// <param name="objectName">optional object objectName</param>
		/// <param name="record">optional NetOffice.AccessApi.Enums.AcRecord Record = 1</param>
		[CustomMethod]
		[SupportByVersion("Access", 9,10,11,12,14,15,16)]
		void GoToRecord(object objectType, object objectName, object record);

		/// <summary>
		/// SupportByVersion Access 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff835648.aspx </remarks>
		/// <param name="hourglassOn">object hourglassOn</param>
		[SupportByVersion("Access", 9,10,11,12,14,15,16)]
		void Hourglass(object hourglassOn);

		/// <summary>
		/// SupportByVersion Access 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff195449.aspx </remarks>
		[SupportByVersion("Access", 9,10,11,12,14,15,16)]
		void Maximize();

		/// <summary>
		/// SupportByVersion Access 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff837032.aspx </remarks>
		[SupportByVersion("Access", 9,10,11,12,14,15,16)]
		void Minimize();

		/// <summary>
		/// SupportByVersion Access 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff197394.aspx </remarks>
		/// <param name="right">optional object right</param>
		/// <param name="down">optional object down</param>
		/// <param name="width">optional object width</param>
		/// <param name="height">optional object height</param>
		[SupportByVersion("Access", 9,10,11,12,14,15,16)]
		void MoveSize(object right, object down, object width, object height);

		/// <summary>
		/// SupportByVersion Access 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff197394.aspx </remarks>
		[CustomMethod]
		[SupportByVersion("Access", 9,10,11,12,14,15,16)]
		void MoveSize();

		/// <summary>
		/// SupportByVersion Access 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff197394.aspx </remarks>
		/// <param name="right">optional object right</param>
		[CustomMethod]
		[SupportByVersion("Access", 9,10,11,12,14,15,16)]
		void MoveSize(object right);

		/// <summary>
		/// SupportByVersion Access 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff197394.aspx </remarks>
		/// <param name="right">optional object right</param>
		/// <param name="down">optional object down</param>
		[CustomMethod]
		[SupportByVersion("Access", 9,10,11,12,14,15,16)]
		void MoveSize(object right, object down);

		/// <summary>
		/// SupportByVersion Access 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff197394.aspx </remarks>
		/// <param name="right">optional object right</param>
		/// <param name="down">optional object down</param>
		/// <param name="width">optional object width</param>
		[CustomMethod]
		[SupportByVersion("Access", 9,10,11,12,14,15,16)]
		void MoveSize(object right, object down, object width);

		/// <summary>
		/// SupportByVersion Access 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff820845.aspx </remarks>
		/// <param name="formName">object formName</param>
		/// <param name="view">optional NetOffice.AccessApi.Enums.AcFormView View = 0</param>
		/// <param name="filterName">optional object filterName</param>
		/// <param name="whereCondition">optional object whereCondition</param>
		/// <param name="dataMode">optional NetOffice.AccessApi.Enums.AcFormOpenDataMode DataMode = -1</param>
		/// <param name="windowMode">optional NetOffice.AccessApi.Enums.AcWindowMode WindowMode = 0</param>
		/// <param name="openArgs">optional object openArgs</param>
		[SupportByVersion("Access", 9,10,11,12,14,15,16)]
		void OpenForm(object formName, object view, object filterName, object whereCondition, object dataMode, object windowMode, object openArgs);

		/// <summary>
		/// SupportByVersion Access 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff820845.aspx </remarks>
		/// <param name="formName">object formName</param>
		[CustomMethod]
		[SupportByVersion("Access", 9,10,11,12,14,15,16)]
		void OpenForm(object formName);

		/// <summary>
		/// SupportByVersion Access 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff820845.aspx </remarks>
		/// <param name="formName">object formName</param>
		/// <param name="view">optional NetOffice.AccessApi.Enums.AcFormView View = 0</param>
		[CustomMethod]
		[SupportByVersion("Access", 9,10,11,12,14,15,16)]
		void OpenForm(object formName, object view);

		/// <summary>
		/// SupportByVersion Access 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff820845.aspx </remarks>
		/// <param name="formName">object formName</param>
		/// <param name="view">optional NetOffice.AccessApi.Enums.AcFormView View = 0</param>
		/// <param name="filterName">optional object filterName</param>
		[CustomMethod]
		[SupportByVersion("Access", 9,10,11,12,14,15,16)]
		void OpenForm(object formName, object view, object filterName);

		/// <summary>
		/// SupportByVersion Access 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff820845.aspx </remarks>
		/// <param name="formName">object formName</param>
		/// <param name="view">optional NetOffice.AccessApi.Enums.AcFormView View = 0</param>
		/// <param name="filterName">optional object filterName</param>
		/// <param name="whereCondition">optional object whereCondition</param>
		[CustomMethod]
		[SupportByVersion("Access", 9,10,11,12,14,15,16)]
		void OpenForm(object formName, object view, object filterName, object whereCondition);

		/// <summary>
		/// SupportByVersion Access 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff820845.aspx </remarks>
		/// <param name="formName">object formName</param>
		/// <param name="view">optional NetOffice.AccessApi.Enums.AcFormView View = 0</param>
		/// <param name="filterName">optional object filterName</param>
		/// <param name="whereCondition">optional object whereCondition</param>
		/// <param name="dataMode">optional NetOffice.AccessApi.Enums.AcFormOpenDataMode DataMode = -1</param>
		[CustomMethod]
		[SupportByVersion("Access", 9,10,11,12,14,15,16)]
		void OpenForm(object formName, object view, object filterName, object whereCondition, object dataMode);

		/// <summary>
		/// SupportByVersion Access 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff820845.aspx </remarks>
		/// <param name="formName">object formName</param>
		/// <param name="view">optional NetOffice.AccessApi.Enums.AcFormView View = 0</param>
		/// <param name="filterName">optional object filterName</param>
		/// <param name="whereCondition">optional object whereCondition</param>
		/// <param name="dataMode">optional NetOffice.AccessApi.Enums.AcFormOpenDataMode DataMode = -1</param>
		/// <param name="windowMode">optional NetOffice.AccessApi.Enums.AcWindowMode WindowMode = 0</param>
		[CustomMethod]
		[SupportByVersion("Access", 9,10,11,12,14,15,16)]
		void OpenForm(object formName, object view, object filterName, object whereCondition, object dataMode, object windowMode);

		/// <summary>
		/// SupportByVersion Access 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff192746.aspx </remarks>
		/// <param name="queryName">object queryName</param>
		/// <param name="view">optional NetOffice.AccessApi.Enums.AcView View = 0</param>
		/// <param name="dataMode">optional NetOffice.AccessApi.Enums.AcOpenDataMode DataMode = 1</param>
		[SupportByVersion("Access", 9,10,11,12,14,15,16)]
		void OpenQuery(object queryName, object view, object dataMode);

		/// <summary>
		/// SupportByVersion Access 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff192746.aspx </remarks>
		/// <param name="queryName">object queryName</param>
		[CustomMethod]
		[SupportByVersion("Access", 9,10,11,12,14,15,16)]
		void OpenQuery(object queryName);

		/// <summary>
		/// SupportByVersion Access 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff192746.aspx </remarks>
		/// <param name="queryName">object queryName</param>
		/// <param name="view">optional NetOffice.AccessApi.Enums.AcView View = 0</param>
		[CustomMethod]
		[SupportByVersion("Access", 9,10,11,12,14,15,16)]
		void OpenQuery(object queryName, object view);

		/// <summary>
		/// SupportByVersion Access 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff194975.aspx </remarks>
		/// <param name="tableName">object tableName</param>
		/// <param name="view">optional NetOffice.AccessApi.Enums.AcView View = 0</param>
		/// <param name="dataMode">optional NetOffice.AccessApi.Enums.AcOpenDataMode DataMode = 1</param>
		[SupportByVersion("Access", 9,10,11,12,14,15,16)]
		void OpenTable(object tableName, object view, object dataMode);

		/// <summary>
		/// SupportByVersion Access 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff194975.aspx </remarks>
		/// <param name="tableName">object tableName</param>
		[CustomMethod]
		[SupportByVersion("Access", 9,10,11,12,14,15,16)]
		void OpenTable(object tableName);

		/// <summary>
		/// SupportByVersion Access 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff194975.aspx </remarks>
		/// <param name="tableName">object tableName</param>
		/// <param name="view">optional NetOffice.AccessApi.Enums.AcView View = 0</param>
		[CustomMethod]
		[SupportByVersion("Access", 9,10,11,12,14,15,16)]
		void OpenTable(object tableName, object view);

		/// <summary>
		/// SupportByVersion Access 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff192667.aspx </remarks>
		/// <param name="printRange">optional NetOffice.AccessApi.Enums.AcPrintRange PrintRange = 0</param>
		/// <param name="pageFrom">optional object pageFrom</param>
		/// <param name="pageTo">optional object pageTo</param>
		/// <param name="printQuality">optional NetOffice.AccessApi.Enums.AcPrintQuality PrintQuality = 0</param>
		/// <param name="copies">optional object copies</param>
		/// <param name="collateCopies">optional object collateCopies</param>
		[SupportByVersion("Access", 9,10,11,12,14,15,16)]
		void PrintOut(object printRange, object pageFrom, object pageTo, object printQuality, object copies, object collateCopies);

		/// <summary>
		/// SupportByVersion Access 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff192667.aspx </remarks>
		[CustomMethod]
		[SupportByVersion("Access", 9,10,11,12,14,15,16)]
		void PrintOut();

		/// <summary>
		/// SupportByVersion Access 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff192667.aspx </remarks>
		/// <param name="printRange">optional NetOffice.AccessApi.Enums.AcPrintRange PrintRange = 0</param>
		[CustomMethod]
		[SupportByVersion("Access", 9,10,11,12,14,15,16)]
		void PrintOut(object printRange);

		/// <summary>
		/// SupportByVersion Access 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff192667.aspx </remarks>
		/// <param name="printRange">optional NetOffice.AccessApi.Enums.AcPrintRange PrintRange = 0</param>
		/// <param name="pageFrom">optional object pageFrom</param>
		[CustomMethod]
		[SupportByVersion("Access", 9,10,11,12,14,15,16)]
		void PrintOut(object printRange, object pageFrom);

		/// <summary>
		/// SupportByVersion Access 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff192667.aspx </remarks>
		/// <param name="printRange">optional NetOffice.AccessApi.Enums.AcPrintRange PrintRange = 0</param>
		/// <param name="pageFrom">optional object pageFrom</param>
		/// <param name="pageTo">optional object pageTo</param>
		[CustomMethod]
		[SupportByVersion("Access", 9,10,11,12,14,15,16)]
		void PrintOut(object printRange, object pageFrom, object pageTo);

		/// <summary>
		/// SupportByVersion Access 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff192667.aspx </remarks>
		/// <param name="printRange">optional NetOffice.AccessApi.Enums.AcPrintRange PrintRange = 0</param>
		/// <param name="pageFrom">optional object pageFrom</param>
		/// <param name="pageTo">optional object pageTo</param>
		/// <param name="printQuality">optional NetOffice.AccessApi.Enums.AcPrintQuality PrintQuality = 0</param>
		[CustomMethod]
		[SupportByVersion("Access", 9,10,11,12,14,15,16)]
		void PrintOut(object printRange, object pageFrom, object pageTo, object printQuality);

		/// <summary>
		/// SupportByVersion Access 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff192667.aspx </remarks>
		/// <param name="printRange">optional NetOffice.AccessApi.Enums.AcPrintRange PrintRange = 0</param>
		/// <param name="pageFrom">optional object pageFrom</param>
		/// <param name="pageTo">optional object pageTo</param>
		/// <param name="printQuality">optional NetOffice.AccessApi.Enums.AcPrintQuality PrintQuality = 0</param>
		/// <param name="copies">optional object copies</param>
		[CustomMethod]
		[SupportByVersion("Access", 9,10,11,12,14,15,16)]
		void PrintOut(object printRange, object pageFrom, object pageTo, object printQuality, object copies);

		/// <summary>
		/// SupportByVersion Access 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff191887.aspx </remarks>
		/// <param name="options">optional NetOffice.AccessApi.Enums.AcQuitOption Options = 1</param>
		[SupportByVersion("Access", 9,10,11,12,14,15,16)]
		void Quit(object options);

		/// <summary>
		/// SupportByVersion Access 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff191887.aspx </remarks>
		[CustomMethod]
		[SupportByVersion("Access", 9,10,11,12,14,15,16)]
		void Quit();

		/// <summary>
		/// SupportByVersion Access 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff195253.aspx </remarks>
		/// <param name="controlName">optional object controlName</param>
		[SupportByVersion("Access", 9,10,11,12,14,15,16)]
		void Requery(object controlName);

		/// <summary>
		/// SupportByVersion Access 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff195253.aspx </remarks>
		[CustomMethod]
		[SupportByVersion("Access", 9,10,11,12,14,15,16)]
		void Requery();

		/// <summary>
		/// SupportByVersion Access 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff195560.aspx </remarks>
		/// <param name="objectType">optional NetOffice.AccessApi.Enums.AcObjectType ObjectType = -1</param>
		/// <param name="objectName">optional object objectName</param>
		[SupportByVersion("Access", 9,10,11,12,14,15,16)]
		void RepaintObject(object objectType, object objectName);

		/// <summary>
		/// SupportByVersion Access 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff195560.aspx </remarks>
		[CustomMethod]
		[SupportByVersion("Access", 9,10,11,12,14,15,16)]
		void RepaintObject();

		/// <summary>
		/// SupportByVersion Access 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff195560.aspx </remarks>
		/// <param name="objectType">optional NetOffice.AccessApi.Enums.AcObjectType ObjectType = -1</param>
		[CustomMethod]
		[SupportByVersion("Access", 9,10,11,12,14,15,16)]
		void RepaintObject(object objectType);

		/// <summary>
		/// SupportByVersion Access 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff823209.aspx </remarks>
		/// <param name="newName">object newName</param>
		/// <param name="objectType">optional NetOffice.AccessApi.Enums.AcObjectType ObjectType = -1</param>
		/// <param name="oldName">optional object oldName</param>
		[SupportByVersion("Access", 9,10,11,12,14,15,16)]
		void Rename(object newName, object objectType, object oldName);

		/// <summary>
		/// SupportByVersion Access 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff823209.aspx </remarks>
		/// <param name="newName">object newName</param>
		[CustomMethod]
		[SupportByVersion("Access", 9,10,11,12,14,15,16)]
		void Rename(object newName);

		/// <summary>
		/// SupportByVersion Access 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff823209.aspx </remarks>
		/// <param name="newName">object newName</param>
		/// <param name="objectType">optional NetOffice.AccessApi.Enums.AcObjectType ObjectType = -1</param>
		[CustomMethod]
		[SupportByVersion("Access", 9,10,11,12,14,15,16)]
		void Rename(object newName, object objectType);

		/// <summary>
		/// SupportByVersion Access 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff193174.aspx </remarks>
		[SupportByVersion("Access", 9,10,11,12,14,15,16)]
		void Restore();

		/// <summary>
		/// SupportByVersion Access 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff192075.aspx </remarks>
		/// <param name="macroName">object macroName</param>
		/// <param name="repeatCount">optional object repeatCount</param>
		/// <param name="repeatExpression">optional object repeatExpression</param>
		[SupportByVersion("Access", 9,10,11,12,14,15,16)]
		void RunMacro(object macroName, object repeatCount, object repeatExpression);

		/// <summary>
		/// SupportByVersion Access 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff192075.aspx </remarks>
		/// <param name="macroName">object macroName</param>
		[CustomMethod]
		[SupportByVersion("Access", 9,10,11,12,14,15,16)]
		void RunMacro(object macroName);

		/// <summary>
		/// SupportByVersion Access 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff192075.aspx </remarks>
		/// <param name="macroName">object macroName</param>
		/// <param name="repeatCount">optional object repeatCount</param>
		[CustomMethod]
		[SupportByVersion("Access", 9,10,11,12,14,15,16)]
		void RunMacro(object macroName, object repeatCount);

		/// <summary>
		/// SupportByVersion Access 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff194626.aspx </remarks>
		/// <param name="sQLStatement">object sQLStatement</param>
		/// <param name="useTransaction">optional object useTransaction</param>
		[SupportByVersion("Access", 9,10,11,12,14,15,16)]
		void RunSQL(object sQLStatement, object useTransaction);

		/// <summary>
		/// SupportByVersion Access 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff194626.aspx </remarks>
		/// <param name="sQLStatement">object sQLStatement</param>
		[CustomMethod]
		[SupportByVersion("Access", 9,10,11,12,14,15,16)]
		void RunSQL(object sQLStatement);

		/// <summary>
		/// SupportByVersion Access 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff835629.aspx </remarks>
		/// <param name="objectType">NetOffice.AccessApi.Enums.AcObjectType objectType</param>
		/// <param name="objectName">optional object objectName</param>
		/// <param name="inDatabaseWindow">optional object inDatabaseWindow</param>
		[SupportByVersion("Access", 9,10,11,12,14,15,16)]
		void SelectObject(NetOffice.AccessApi.Enums.AcObjectType objectType, object objectName, object inDatabaseWindow);

		/// <summary>
		/// SupportByVersion Access 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff835629.aspx </remarks>
		/// <param name="objectType">NetOffice.AccessApi.Enums.AcObjectType objectType</param>
		[CustomMethod]
		[SupportByVersion("Access", 9,10,11,12,14,15,16)]
		void SelectObject(NetOffice.AccessApi.Enums.AcObjectType objectType);

		/// <summary>
		/// SupportByVersion Access 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff835629.aspx </remarks>
		/// <param name="objectType">NetOffice.AccessApi.Enums.AcObjectType objectType</param>
		/// <param name="objectName">optional object objectName</param>
		[CustomMethod]
		[SupportByVersion("Access", 9,10,11,12,14,15,16)]
		void SelectObject(NetOffice.AccessApi.Enums.AcObjectType objectType, object objectName);

		/// <summary>
		/// SupportByVersion Access 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff837275.aspx </remarks>
		/// <param name="warningsOn">object warningsOn</param>
		[SupportByVersion("Access", 9,10,11,12,14,15,16)]
		void SetWarnings(object warningsOn);

		/// <summary>
		/// SupportByVersion Access 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff195994.aspx </remarks>
		[SupportByVersion("Access", 9,10,11,12,14,15,16)]
		void ShowAllRecords();

		/// <summary>
		/// SupportByVersion Access 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff192676.aspx </remarks>
		/// <param name="reportName">object reportName</param>
		/// <param name="view">optional NetOffice.AccessApi.Enums.AcView View = 0</param>
		/// <param name="filterName">optional object filterName</param>
		/// <param name="whereCondition">optional object whereCondition</param>
		[SupportByVersion("Access", 9,10,11,12,14,15,16)]
		void OpenReport(object reportName, object view, object filterName, object whereCondition);

		/// <summary>
		/// SupportByVersion Access 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff192676.aspx </remarks>
		/// <param name="reportName">object reportName</param>
		/// <param name="view">optional NetOffice.AccessApi.Enums.AcView View = 0</param>
		/// <param name="filterName">optional object filterName</param>
		/// <param name="whereCondition">optional object whereCondition</param>
		/// <param name="windowMode">optional NetOffice.AccessApi.Enums.AcWindowMode WindowMode = 0</param>
		/// <param name="openArgs">optional object openArgs</param>
		[SupportByVersion("Access", 10,11,12,14,15,16)]
		void OpenReport(object reportName, object view, object filterName, object whereCondition, object windowMode, object openArgs);

		/// <summary>
		/// SupportByVersion Access 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff192676.aspx </remarks>
		/// <param name="reportName">object reportName</param>
		[CustomMethod]
		[SupportByVersion("Access", 9,10,11,12,14,15,16)]
		void OpenReport(object reportName);

		/// <summary>
		/// SupportByVersion Access 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff192676.aspx </remarks>
		/// <param name="reportName">object reportName</param>
		/// <param name="view">optional NetOffice.AccessApi.Enums.AcView View = 0</param>
		[CustomMethod]
		[SupportByVersion("Access", 9,10,11,12,14,15,16)]
		void OpenReport(object reportName, object view);

		/// <summary>
		/// SupportByVersion Access 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff192676.aspx </remarks>
		/// <param name="reportName">object reportName</param>
		/// <param name="view">optional NetOffice.AccessApi.Enums.AcView View = 0</param>
		/// <param name="filterName">optional object filterName</param>
		[CustomMethod]
		[SupportByVersion("Access", 9,10,11,12,14,15,16)]
		void OpenReport(object reportName, object view, object filterName);

		/// <summary>
		/// SupportByVersion Access 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff192676.aspx </remarks>
		/// <param name="reportName">object reportName</param>
		/// <param name="view">optional NetOffice.AccessApi.Enums.AcView View = 0</param>
		/// <param name="filterName">optional object filterName</param>
		/// <param name="whereCondition">optional object whereCondition</param>
		/// <param name="windowMode">optional NetOffice.AccessApi.Enums.AcWindowMode WindowMode = 0</param>
		[CustomMethod]
		[SupportByVersion("Access", 10,11,12,14,15,16)]
		void OpenReport(object reportName, object view, object filterName, object whereCondition, object windowMode);

		/// <summary>
		/// SupportByVersion Access 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff196455.aspx </remarks>
		/// <param name="transferType">optional NetOffice.AccessApi.Enums.AcDataTransferType TransferType = 0</param>
		/// <param name="databaseType">optional object databaseType</param>
		/// <param name="databaseName">optional object databaseName</param>
		/// <param name="objectType">optional NetOffice.AccessApi.Enums.AcObjectType ObjectType = 0</param>
		/// <param name="source">optional object source</param>
		/// <param name="destination">optional object destination</param>
		/// <param name="structureOnly">optional object structureOnly</param>
		/// <param name="storeLogin">optional object storeLogin</param>
		[SupportByVersion("Access", 9,10,11,12,14,15,16)]
		void TransferDatabase(object transferType, object databaseType, object databaseName, object objectType, object source, object destination, object structureOnly, object storeLogin);

		/// <summary>
		/// SupportByVersion Access 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff196455.aspx </remarks>
		[CustomMethod]
		[SupportByVersion("Access", 9,10,11,12,14,15,16)]
		void TransferDatabase();

		/// <summary>
		/// SupportByVersion Access 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff196455.aspx </remarks>
		/// <param name="transferType">optional NetOffice.AccessApi.Enums.AcDataTransferType TransferType = 0</param>
		[CustomMethod]
		[SupportByVersion("Access", 9,10,11,12,14,15,16)]
		void TransferDatabase(object transferType);

		/// <summary>
		/// SupportByVersion Access 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff196455.aspx </remarks>
		/// <param name="transferType">optional NetOffice.AccessApi.Enums.AcDataTransferType TransferType = 0</param>
		/// <param name="databaseType">optional object databaseType</param>
		[CustomMethod]
		[SupportByVersion("Access", 9,10,11,12,14,15,16)]
		void TransferDatabase(object transferType, object databaseType);

		/// <summary>
		/// SupportByVersion Access 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff196455.aspx </remarks>
		/// <param name="transferType">optional NetOffice.AccessApi.Enums.AcDataTransferType TransferType = 0</param>
		/// <param name="databaseType">optional object databaseType</param>
		/// <param name="databaseName">optional object databaseName</param>
		[CustomMethod]
		[SupportByVersion("Access", 9,10,11,12,14,15,16)]
		void TransferDatabase(object transferType, object databaseType, object databaseName);

		/// <summary>
		/// SupportByVersion Access 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff196455.aspx </remarks>
		/// <param name="transferType">optional NetOffice.AccessApi.Enums.AcDataTransferType TransferType = 0</param>
		/// <param name="databaseType">optional object databaseType</param>
		/// <param name="databaseName">optional object databaseName</param>
		/// <param name="objectType">optional NetOffice.AccessApi.Enums.AcObjectType ObjectType = 0</param>
		[CustomMethod]
		[SupportByVersion("Access", 9,10,11,12,14,15,16)]
		void TransferDatabase(object transferType, object databaseType, object databaseName, object objectType);

		/// <summary>
		/// SupportByVersion Access 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff196455.aspx </remarks>
		/// <param name="transferType">optional NetOffice.AccessApi.Enums.AcDataTransferType TransferType = 0</param>
		/// <param name="databaseType">optional object databaseType</param>
		/// <param name="databaseName">optional object databaseName</param>
		/// <param name="objectType">optional NetOffice.AccessApi.Enums.AcObjectType ObjectType = 0</param>
		/// <param name="source">optional object source</param>
		[CustomMethod]
		[SupportByVersion("Access", 9,10,11,12,14,15,16)]
		void TransferDatabase(object transferType, object databaseType, object databaseName, object objectType, object source);

		/// <summary>
		/// SupportByVersion Access 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff196455.aspx </remarks>
		/// <param name="transferType">optional NetOffice.AccessApi.Enums.AcDataTransferType TransferType = 0</param>
		/// <param name="databaseType">optional object databaseType</param>
		/// <param name="databaseName">optional object databaseName</param>
		/// <param name="objectType">optional NetOffice.AccessApi.Enums.AcObjectType ObjectType = 0</param>
		/// <param name="source">optional object source</param>
		/// <param name="destination">optional object destination</param>
		[CustomMethod]
		[SupportByVersion("Access", 9,10,11,12,14,15,16)]
		void TransferDatabase(object transferType, object databaseType, object databaseName, object objectType, object source, object destination);

		/// <summary>
		/// SupportByVersion Access 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff196455.aspx </remarks>
		/// <param name="transferType">optional NetOffice.AccessApi.Enums.AcDataTransferType TransferType = 0</param>
		/// <param name="databaseType">optional object databaseType</param>
		/// <param name="databaseName">optional object databaseName</param>
		/// <param name="objectType">optional NetOffice.AccessApi.Enums.AcObjectType ObjectType = 0</param>
		/// <param name="source">optional object source</param>
		/// <param name="destination">optional object destination</param>
		/// <param name="structureOnly">optional object structureOnly</param>
		[CustomMethod]
		[SupportByVersion("Access", 9,10,11,12,14,15,16)]
		void TransferDatabase(object transferType, object databaseType, object databaseName, object objectType, object source, object destination, object structureOnly);

		/// <summary>
		/// SupportByVersion Access 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff844793.aspx </remarks>
		/// <param name="transferType">optional NetOffice.AccessApi.Enums.AcDataTransferType TransferType = 0</param>
		/// <param name="spreadsheetType">optional NetOffice.AccessApi.Enums.AcSpreadSheetType SpreadsheetType = 8</param>
		/// <param name="tableName">optional object tableName</param>
		/// <param name="fileName">optional object fileName</param>
		/// <param name="hasFieldNames">optional object hasFieldNames</param>
		/// <param name="range">optional object range</param>
		/// <param name="useOA">optional object useOA</param>
		[SupportByVersion("Access", 9,10,11,12,14,15,16)]
		void TransferSpreadsheet(object transferType, object spreadsheetType, object tableName, object fileName, object hasFieldNames, object range, object useOA);

		/// <summary>
		/// SupportByVersion Access 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff844793.aspx </remarks>
		[CustomMethod]
		[SupportByVersion("Access", 9,10,11,12,14,15,16)]
		void TransferSpreadsheet();

		/// <summary>
		/// SupportByVersion Access 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff844793.aspx </remarks>
		/// <param name="transferType">optional NetOffice.AccessApi.Enums.AcDataTransferType TransferType = 0</param>
		[CustomMethod]
		[SupportByVersion("Access", 9,10,11,12,14,15,16)]
		void TransferSpreadsheet(object transferType);

		/// <summary>
		/// SupportByVersion Access 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff844793.aspx </remarks>
		/// <param name="transferType">optional NetOffice.AccessApi.Enums.AcDataTransferType TransferType = 0</param>
		/// <param name="spreadsheetType">optional NetOffice.AccessApi.Enums.AcSpreadSheetType SpreadsheetType = 8</param>
		[CustomMethod]
		[SupportByVersion("Access", 9,10,11,12,14,15,16)]
		void TransferSpreadsheet(object transferType, object spreadsheetType);

		/// <summary>
		/// SupportByVersion Access 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff844793.aspx </remarks>
		/// <param name="transferType">optional NetOffice.AccessApi.Enums.AcDataTransferType TransferType = 0</param>
		/// <param name="spreadsheetType">optional NetOffice.AccessApi.Enums.AcSpreadSheetType SpreadsheetType = 8</param>
		/// <param name="tableName">optional object tableName</param>
		[CustomMethod]
		[SupportByVersion("Access", 9,10,11,12,14,15,16)]
		void TransferSpreadsheet(object transferType, object spreadsheetType, object tableName);

		/// <summary>
		/// SupportByVersion Access 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff844793.aspx </remarks>
		/// <param name="transferType">optional NetOffice.AccessApi.Enums.AcDataTransferType TransferType = 0</param>
		/// <param name="spreadsheetType">optional NetOffice.AccessApi.Enums.AcSpreadSheetType SpreadsheetType = 8</param>
		/// <param name="tableName">optional object tableName</param>
		/// <param name="fileName">optional object fileName</param>
		[CustomMethod]
		[SupportByVersion("Access", 9,10,11,12,14,15,16)]
		void TransferSpreadsheet(object transferType, object spreadsheetType, object tableName, object fileName);

		/// <summary>
		/// SupportByVersion Access 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff844793.aspx </remarks>
		/// <param name="transferType">optional NetOffice.AccessApi.Enums.AcDataTransferType TransferType = 0</param>
		/// <param name="spreadsheetType">optional NetOffice.AccessApi.Enums.AcSpreadSheetType SpreadsheetType = 8</param>
		/// <param name="tableName">optional object tableName</param>
		/// <param name="fileName">optional object fileName</param>
		/// <param name="hasFieldNames">optional object hasFieldNames</param>
		[CustomMethod]
		[SupportByVersion("Access", 9,10,11,12,14,15,16)]
		void TransferSpreadsheet(object transferType, object spreadsheetType, object tableName, object fileName, object hasFieldNames);

		/// <summary>
		/// SupportByVersion Access 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff844793.aspx </remarks>
		/// <param name="transferType">optional NetOffice.AccessApi.Enums.AcDataTransferType TransferType = 0</param>
		/// <param name="spreadsheetType">optional NetOffice.AccessApi.Enums.AcSpreadSheetType SpreadsheetType = 8</param>
		/// <param name="tableName">optional object tableName</param>
		/// <param name="fileName">optional object fileName</param>
		/// <param name="hasFieldNames">optional object hasFieldNames</param>
		/// <param name="range">optional object range</param>
		[CustomMethod]
		[SupportByVersion("Access", 9,10,11,12,14,15,16)]
		void TransferSpreadsheet(object transferType, object spreadsheetType, object tableName, object fileName, object hasFieldNames, object range);

		/// <summary>
		/// SupportByVersion Access 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff835958.aspx </remarks>
		/// <param name="transferType">optional NetOffice.AccessApi.Enums.AcTextTransferType TransferType = 0</param>
		/// <param name="specificationName">optional object specificationName</param>
		/// <param name="tableName">optional object tableName</param>
		/// <param name="fileName">optional object fileName</param>
		/// <param name="hasFieldNames">optional object hasFieldNames</param>
		/// <param name="hTMLTableName">optional object hTMLTableName</param>
		/// <param name="codePage">optional object codePage</param>
		[SupportByVersion("Access", 9,10,11,12,14,15,16)]
		void TransferText(object transferType, object specificationName, object tableName, object fileName, object hasFieldNames, object hTMLTableName, object codePage);

		/// <summary>
		/// SupportByVersion Access 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff835958.aspx </remarks>
		[CustomMethod]
		[SupportByVersion("Access", 9,10,11,12,14,15,16)]
		void TransferText();

		/// <summary>
		/// SupportByVersion Access 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff835958.aspx </remarks>
		/// <param name="transferType">optional NetOffice.AccessApi.Enums.AcTextTransferType TransferType = 0</param>
		[CustomMethod]
		[SupportByVersion("Access", 9,10,11,12,14,15,16)]
		void TransferText(object transferType);

		/// <summary>
		/// SupportByVersion Access 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff835958.aspx </remarks>
		/// <param name="transferType">optional NetOffice.AccessApi.Enums.AcTextTransferType TransferType = 0</param>
		/// <param name="specificationName">optional object specificationName</param>
		[CustomMethod]
		[SupportByVersion("Access", 9,10,11,12,14,15,16)]
		void TransferText(object transferType, object specificationName);

		/// <summary>
		/// SupportByVersion Access 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff835958.aspx </remarks>
		/// <param name="transferType">optional NetOffice.AccessApi.Enums.AcTextTransferType TransferType = 0</param>
		/// <param name="specificationName">optional object specificationName</param>
		/// <param name="tableName">optional object tableName</param>
		[CustomMethod]
		[SupportByVersion("Access", 9,10,11,12,14,15,16)]
		void TransferText(object transferType, object specificationName, object tableName);

		/// <summary>
		/// SupportByVersion Access 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff835958.aspx </remarks>
		/// <param name="transferType">optional NetOffice.AccessApi.Enums.AcTextTransferType TransferType = 0</param>
		/// <param name="specificationName">optional object specificationName</param>
		/// <param name="tableName">optional object tableName</param>
		/// <param name="fileName">optional object fileName</param>
		[CustomMethod]
		[SupportByVersion("Access", 9,10,11,12,14,15,16)]
		void TransferText(object transferType, object specificationName, object tableName, object fileName);

		/// <summary>
		/// SupportByVersion Access 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff835958.aspx </remarks>
		/// <param name="transferType">optional NetOffice.AccessApi.Enums.AcTextTransferType TransferType = 0</param>
		/// <param name="specificationName">optional object specificationName</param>
		/// <param name="tableName">optional object tableName</param>
		/// <param name="fileName">optional object fileName</param>
		/// <param name="hasFieldNames">optional object hasFieldNames</param>
		[CustomMethod]
		[SupportByVersion("Access", 9,10,11,12,14,15,16)]
		void TransferText(object transferType, object specificationName, object tableName, object fileName, object hasFieldNames);

		/// <summary>
		/// SupportByVersion Access 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff835958.aspx </remarks>
		/// <param name="transferType">optional NetOffice.AccessApi.Enums.AcTextTransferType TransferType = 0</param>
		/// <param name="specificationName">optional object specificationName</param>
		/// <param name="tableName">optional object tableName</param>
		/// <param name="fileName">optional object fileName</param>
		/// <param name="hasFieldNames">optional object hasFieldNames</param>
		/// <param name="hTMLTableName">optional object hTMLTableName</param>
		[CustomMethod]
		[SupportByVersion("Access", 9,10,11,12,14,15,16)]
		void TransferText(object transferType, object specificationName, object tableName, object fileName, object hasFieldNames, object hTMLTableName);

		/// <summary>
		/// SupportByVersion Access 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff192065.aspx </remarks>
		/// <param name="objectType">NetOffice.AccessApi.Enums.AcOutputObjectType objectType</param>
		/// <param name="objectName">optional object objectName</param>
		/// <param name="outputFormat">optional object outputFormat</param>
		/// <param name="outputFile">optional object outputFile</param>
		/// <param name="autoStart">optional object autoStart</param>
		/// <param name="templateFile">optional object templateFile</param>
		[SupportByVersion("Access", 9,10,11,12,14,15,16)]
		void OutputTo(NetOffice.AccessApi.Enums.AcOutputObjectType objectType, object objectName, object outputFormat, object outputFile, object autoStart, object templateFile);

		/// <summary>
		/// SupportByVersion Access 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff192065.aspx </remarks>
		/// <param name="objectType">NetOffice.AccessApi.Enums.AcOutputObjectType objectType</param>
		/// <param name="objectName">optional object objectName</param>
		/// <param name="outputFormat">optional object outputFormat</param>
		/// <param name="outputFile">optional object outputFile</param>
		/// <param name="autoStart">optional object autoStart</param>
		/// <param name="templateFile">optional object templateFile</param>
		/// <param name="encoding">optional object encoding</param>
		[SupportByVersion("Access", 10,11,12,14,15,16)]
		void OutputTo(NetOffice.AccessApi.Enums.AcOutputObjectType objectType, object objectName, object outputFormat, object outputFile, object autoStart, object templateFile, object encoding);

		/// <summary>
		/// SupportByVersion Access 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff192065.aspx </remarks>
		/// <param name="objectType">NetOffice.AccessApi.Enums.AcOutputObjectType objectType</param>
		/// <param name="objectName">optional object objectName</param>
		/// <param name="outputFormat">optional object outputFormat</param>
		/// <param name="outputFile">optional object outputFile</param>
		/// <param name="autoStart">optional object autoStart</param>
		/// <param name="templateFile">optional object templateFile</param>
		/// <param name="encoding">optional object encoding</param>
		/// <param name="outputQuality">optional NetOffice.AccessApi.Enums.AcExportQuality OutputQuality = 0</param>
		[SupportByVersion("Access", 12,14,15,16)]
		void OutputTo(NetOffice.AccessApi.Enums.AcOutputObjectType objectType, object objectName, object outputFormat, object outputFile, object autoStart, object templateFile, object encoding, object outputQuality);

		/// <summary>
		/// SupportByVersion Access 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff192065.aspx </remarks>
		/// <param name="objectType">NetOffice.AccessApi.Enums.AcOutputObjectType objectType</param>
		[CustomMethod]
		[SupportByVersion("Access", 9,10,11,12,14,15,16)]
		void OutputTo(NetOffice.AccessApi.Enums.AcOutputObjectType objectType);

		/// <summary>
		/// SupportByVersion Access 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff192065.aspx </remarks>
		/// <param name="objectType">NetOffice.AccessApi.Enums.AcOutputObjectType objectType</param>
		/// <param name="objectName">optional object objectName</param>
		[CustomMethod]
		[SupportByVersion("Access", 9,10,11,12,14,15,16)]
		void OutputTo(NetOffice.AccessApi.Enums.AcOutputObjectType objectType, object objectName);

		/// <summary>
		/// SupportByVersion Access 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff192065.aspx </remarks>
		/// <param name="objectType">NetOffice.AccessApi.Enums.AcOutputObjectType objectType</param>
		/// <param name="objectName">optional object objectName</param>
		/// <param name="outputFormat">optional object outputFormat</param>
		[CustomMethod]
		[SupportByVersion("Access", 9,10,11,12,14,15,16)]
		void OutputTo(NetOffice.AccessApi.Enums.AcOutputObjectType objectType, object objectName, object outputFormat);

		/// <summary>
		/// SupportByVersion Access 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff192065.aspx </remarks>
		/// <param name="objectType">NetOffice.AccessApi.Enums.AcOutputObjectType objectType</param>
		/// <param name="objectName">optional object objectName</param>
		/// <param name="outputFormat">optional object outputFormat</param>
		/// <param name="outputFile">optional object outputFile</param>
		[CustomMethod]
		[SupportByVersion("Access", 9,10,11,12,14,15,16)]
		void OutputTo(NetOffice.AccessApi.Enums.AcOutputObjectType objectType, object objectName, object outputFormat, object outputFile);

		/// <summary>
		/// SupportByVersion Access 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff192065.aspx </remarks>
		/// <param name="objectType">NetOffice.AccessApi.Enums.AcOutputObjectType objectType</param>
		/// <param name="objectName">optional object objectName</param>
		/// <param name="outputFormat">optional object outputFormat</param>
		/// <param name="outputFile">optional object outputFile</param>
		/// <param name="autoStart">optional object autoStart</param>
		[CustomMethod]
		[SupportByVersion("Access", 9,10,11,12,14,15,16)]
		void OutputTo(NetOffice.AccessApi.Enums.AcOutputObjectType objectType, object objectName, object outputFormat, object outputFile, object autoStart);

		/// <summary>
		/// SupportByVersion Access 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff197376.aspx </remarks>
		/// <param name="objectType">optional NetOffice.AccessApi.Enums.AcObjectType ObjectType = -1</param>
		/// <param name="objectName">optional object objectName</param>
		[SupportByVersion("Access", 9,10,11,12,14,15,16)]
		void DeleteObject(object objectType, object objectName);

		/// <summary>
		/// SupportByVersion Access 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff197376.aspx </remarks>
		[CustomMethod]
		[SupportByVersion("Access", 9,10,11,12,14,15,16)]
		void DeleteObject();

		/// <summary>
		/// SupportByVersion Access 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff197376.aspx </remarks>
		/// <param name="objectType">optional NetOffice.AccessApi.Enums.AcObjectType ObjectType = -1</param>
		[CustomMethod]
		[SupportByVersion("Access", 9,10,11,12,14,15,16)]
		void DeleteObject(object objectType);

		/// <summary>
		/// SupportByVersion Access 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff192698.aspx </remarks>
		/// <param name="moduleName">optional object moduleName</param>
		/// <param name="procedureName">optional object procedureName</param>
		[SupportByVersion("Access", 9,10,11,12,14,15,16)]
		void OpenModule(object moduleName, object procedureName);

		/// <summary>
		/// SupportByVersion Access 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff192698.aspx </remarks>
		[CustomMethod]
		[SupportByVersion("Access", 9,10,11,12,14,15,16)]
		void OpenModule();

		/// <summary>
		/// SupportByVersion Access 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff192698.aspx </remarks>
		/// <param name="moduleName">optional object moduleName</param>
		[CustomMethod]
		[SupportByVersion("Access", 9,10,11,12,14,15,16)]
		void OpenModule(object moduleName);

		/// <summary>
		/// SupportByVersion Access 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff197046.aspx </remarks>
		/// <param name="objectType">optional NetOffice.AccessApi.Enums.AcSendObjectType ObjectType = -1</param>
		/// <param name="objectName">optional object objectName</param>
		/// <param name="outputFormat">optional object outputFormat</param>
		/// <param name="to">optional object to</param>
		/// <param name="cc">optional object cc</param>
		/// <param name="bcc">optional object bcc</param>
		/// <param name="subject">optional object subject</param>
		/// <param name="messageText">optional object messageText</param>
		/// <param name="editMessage">optional object editMessage</param>
		/// <param name="templateFile">optional object templateFile</param>
		[SupportByVersion("Access", 9,10,11,12,14,15,16)]
		void SendObject(object objectType, object objectName, object outputFormat, object to, object cc, object bcc, object subject, object messageText, object editMessage, object templateFile);

		/// <summary>
		/// SupportByVersion Access 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff197046.aspx </remarks>
		[CustomMethod]
		[SupportByVersion("Access", 9,10,11,12,14,15,16)]
		void SendObject();

		/// <summary>
		/// SupportByVersion Access 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff197046.aspx </remarks>
		/// <param name="objectType">optional NetOffice.AccessApi.Enums.AcSendObjectType ObjectType = -1</param>
		[CustomMethod]
		[SupportByVersion("Access", 9,10,11,12,14,15,16)]
		void SendObject(object objectType);

		/// <summary>
		/// SupportByVersion Access 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff197046.aspx </remarks>
		/// <param name="objectType">optional NetOffice.AccessApi.Enums.AcSendObjectType ObjectType = -1</param>
		/// <param name="objectName">optional object objectName</param>
		[CustomMethod]
		[SupportByVersion("Access", 9,10,11,12,14,15,16)]
		void SendObject(object objectType, object objectName);

		/// <summary>
		/// SupportByVersion Access 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff197046.aspx </remarks>
		/// <param name="objectType">optional NetOffice.AccessApi.Enums.AcSendObjectType ObjectType = -1</param>
		/// <param name="objectName">optional object objectName</param>
		/// <param name="outputFormat">optional object outputFormat</param>
		[CustomMethod]
		[SupportByVersion("Access", 9,10,11,12,14,15,16)]
		void SendObject(object objectType, object objectName, object outputFormat);

		/// <summary>
		/// SupportByVersion Access 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff197046.aspx </remarks>
		/// <param name="objectType">optional NetOffice.AccessApi.Enums.AcSendObjectType ObjectType = -1</param>
		/// <param name="objectName">optional object objectName</param>
		/// <param name="outputFormat">optional object outputFormat</param>
		/// <param name="to">optional object to</param>
		[CustomMethod]
		[SupportByVersion("Access", 9,10,11,12,14,15,16)]
		void SendObject(object objectType, object objectName, object outputFormat, object to);

		/// <summary>
		/// SupportByVersion Access 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff197046.aspx </remarks>
		/// <param name="objectType">optional NetOffice.AccessApi.Enums.AcSendObjectType ObjectType = -1</param>
		/// <param name="objectName">optional object objectName</param>
		/// <param name="outputFormat">optional object outputFormat</param>
		/// <param name="to">optional object to</param>
		/// <param name="cc">optional object cc</param>
		[CustomMethod]
		[SupportByVersion("Access", 9,10,11,12,14,15,16)]
		void SendObject(object objectType, object objectName, object outputFormat, object to, object cc);

		/// <summary>
		/// SupportByVersion Access 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff197046.aspx </remarks>
		/// <param name="objectType">optional NetOffice.AccessApi.Enums.AcSendObjectType ObjectType = -1</param>
		/// <param name="objectName">optional object objectName</param>
		/// <param name="outputFormat">optional object outputFormat</param>
		/// <param name="to">optional object to</param>
		/// <param name="cc">optional object cc</param>
		/// <param name="bcc">optional object bcc</param>
		[CustomMethod]
		[SupportByVersion("Access", 9,10,11,12,14,15,16)]
		void SendObject(object objectType, object objectName, object outputFormat, object to, object cc, object bcc);

		/// <summary>
		/// SupportByVersion Access 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff197046.aspx </remarks>
		/// <param name="objectType">optional NetOffice.AccessApi.Enums.AcSendObjectType ObjectType = -1</param>
		/// <param name="objectName">optional object objectName</param>
		/// <param name="outputFormat">optional object outputFormat</param>
		/// <param name="to">optional object to</param>
		/// <param name="cc">optional object cc</param>
		/// <param name="bcc">optional object bcc</param>
		/// <param name="subject">optional object subject</param>
		[CustomMethod]
		[SupportByVersion("Access", 9,10,11,12,14,15,16)]
		void SendObject(object objectType, object objectName, object outputFormat, object to, object cc, object bcc, object subject);

		/// <summary>
		/// SupportByVersion Access 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff197046.aspx </remarks>
		/// <param name="objectType">optional NetOffice.AccessApi.Enums.AcSendObjectType ObjectType = -1</param>
		/// <param name="objectName">optional object objectName</param>
		/// <param name="outputFormat">optional object outputFormat</param>
		/// <param name="to">optional object to</param>
		/// <param name="cc">optional object cc</param>
		/// <param name="bcc">optional object bcc</param>
		/// <param name="subject">optional object subject</param>
		/// <param name="messageText">optional object messageText</param>
		[CustomMethod]
		[SupportByVersion("Access", 9,10,11,12,14,15,16)]
		void SendObject(object objectType, object objectName, object outputFormat, object to, object cc, object bcc, object subject, object messageText);

		/// <summary>
		/// SupportByVersion Access 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff197046.aspx </remarks>
		/// <param name="objectType">optional NetOffice.AccessApi.Enums.AcSendObjectType ObjectType = -1</param>
		/// <param name="objectName">optional object objectName</param>
		/// <param name="outputFormat">optional object outputFormat</param>
		/// <param name="to">optional object to</param>
		/// <param name="cc">optional object cc</param>
		/// <param name="bcc">optional object bcc</param>
		/// <param name="subject">optional object subject</param>
		/// <param name="messageText">optional object messageText</param>
		/// <param name="editMessage">optional object editMessage</param>
		[CustomMethod]
		[SupportByVersion("Access", 9,10,11,12,14,15,16)]
		void SendObject(object objectType, object objectName, object outputFormat, object to, object cc, object bcc, object subject, object messageText, object editMessage);

		/// <summary>
		/// SupportByVersion Access 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff194957.aspx </remarks>
		/// <param name="toolbarName">object toolbarName</param>
		/// <param name="show">optional NetOffice.AccessApi.Enums.AcShowToolbar Show = 0</param>
		[SupportByVersion("Access", 9,10,11,12,14,15,16)]
		void ShowToolbar(object toolbarName, object show);

		/// <summary>
		/// SupportByVersion Access 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff194957.aspx </remarks>
		/// <param name="toolbarName">object toolbarName</param>
		[CustomMethod]
		[SupportByVersion("Access", 9,10,11,12,14,15,16)]
		void ShowToolbar(object toolbarName);

		/// <summary>
		/// SupportByVersion Access 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff196435.aspx </remarks>
		/// <param name="objectType">optional NetOffice.AccessApi.Enums.AcObjectType ObjectType = -1</param>
		/// <param name="objectName">optional object objectName</param>
		[SupportByVersion("Access", 9,10,11,12,14,15,16)]
		void Save(object objectType, object objectName);

		/// <summary>
		/// SupportByVersion Access 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff196435.aspx </remarks>
		[CustomMethod]
		[SupportByVersion("Access", 9,10,11,12,14,15,16)]
		void Save();

		/// <summary>
		/// SupportByVersion Access 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff196435.aspx </remarks>
		/// <param name="objectType">optional NetOffice.AccessApi.Enums.AcObjectType ObjectType = -1</param>
		[CustomMethod]
		[SupportByVersion("Access", 9,10,11,12,14,15,16)]
		void Save(object objectType);

		/// <summary>
		/// SupportByVersion Access 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff195265.aspx </remarks>
		/// <param name="menuIndex">object menuIndex</param>
		/// <param name="commandIndex">optional object commandIndex</param>
		/// <param name="subcommandIndex">optional object subcommandIndex</param>
		/// <param name="flag">optional object flag</param>
		[SupportByVersion("Access", 9,10,11,12,14,15,16)]
		void SetMenuItem(object menuIndex, object commandIndex, object subcommandIndex, object flag);

		/// <summary>
		/// SupportByVersion Access 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff195265.aspx </remarks>
		/// <param name="menuIndex">object menuIndex</param>
		[CustomMethod]
		[SupportByVersion("Access", 9,10,11,12,14,15,16)]
		void SetMenuItem(object menuIndex);

		/// <summary>
		/// SupportByVersion Access 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff195265.aspx </remarks>
		/// <param name="menuIndex">object menuIndex</param>
		/// <param name="commandIndex">optional object commandIndex</param>
		[CustomMethod]
		[SupportByVersion("Access", 9,10,11,12,14,15,16)]
		void SetMenuItem(object menuIndex, object commandIndex);

		/// <summary>
		/// SupportByVersion Access 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff195265.aspx </remarks>
		/// <param name="menuIndex">object menuIndex</param>
		/// <param name="commandIndex">optional object commandIndex</param>
		/// <param name="subcommandIndex">optional object subcommandIndex</param>
		[CustomMethod]
		[SupportByVersion("Access", 9,10,11,12,14,15,16)]
		void SetMenuItem(object menuIndex, object commandIndex, object subcommandIndex);

		/// <summary>
		/// SupportByVersion Access 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff194612.aspx </remarks>
		/// <param name="command">NetOffice.AccessApi.Enums.AcCommand command</param>
		[SupportByVersion("Access", 9,10,11,12,14,15,16)]
		void RunCommand(NetOffice.AccessApi.Enums.AcCommand command);

		/// <summary>
		/// SupportByVersion Access 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff845421.aspx </remarks>
		/// <param name="dataAccessPageName">object dataAccessPageName</param>
		/// <param name="view">optional NetOffice.AccessApi.Enums.AcDataAccessPageView View = 0</param>
		[SupportByVersion("Access", 9,10,11,12,14,15,16)]
		void OpenDataAccessPage(object dataAccessPageName, object view);

		/// <summary>
		/// SupportByVersion Access 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff845421.aspx </remarks>
		/// <param name="dataAccessPageName">object dataAccessPageName</param>
		[CustomMethod]
		[SupportByVersion("Access", 9,10,11,12,14,15,16)]
		void OpenDataAccessPage(object dataAccessPageName);

		/// <summary>
		/// SupportByVersion Access 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff197347.aspx </remarks>
		/// <param name="viewName">object viewName</param>
		/// <param name="view">optional NetOffice.AccessApi.Enums.AcView View = 0</param>
		/// <param name="dataMode">optional NetOffice.AccessApi.Enums.AcOpenDataMode DataMode = 1</param>
		[SupportByVersion("Access", 9,10,11,12,14,15,16)]
		void OpenView(object viewName, object view, object dataMode);

		/// <summary>
		/// SupportByVersion Access 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff197347.aspx </remarks>
		/// <param name="viewName">object viewName</param>
		[CustomMethod]
		[SupportByVersion("Access", 9,10,11,12,14,15,16)]
		void OpenView(object viewName);

		/// <summary>
		/// SupportByVersion Access 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff197347.aspx </remarks>
		/// <param name="viewName">object viewName</param>
		/// <param name="view">optional NetOffice.AccessApi.Enums.AcView View = 0</param>
		[CustomMethod]
		[SupportByVersion("Access", 9,10,11,12,14,15,16)]
		void OpenView(object viewName, object view);

		/// <summary>
		/// SupportByVersion Access 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff821439.aspx </remarks>
		/// <param name="diagramName">object diagramName</param>
		[SupportByVersion("Access", 9,10,11,12,14,15,16)]
		void OpenDiagram(object diagramName);

		/// <summary>
		/// SupportByVersion Access 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff197412.aspx </remarks>
		/// <param name="procedureName">object procedureName</param>
		/// <param name="view">optional NetOffice.AccessApi.Enums.AcView View = 0</param>
		/// <param name="dataMode">optional NetOffice.AccessApi.Enums.AcOpenDataMode DataMode = 1</param>
		[SupportByVersion("Access", 9,10,11,12,14,15,16)]
		void OpenStoredProcedure(object procedureName, object view, object dataMode);

		/// <summary>
		/// SupportByVersion Access 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff197412.aspx </remarks>
		/// <param name="procedureName">object procedureName</param>
		[CustomMethod]
		[SupportByVersion("Access", 9,10,11,12,14,15,16)]
		void OpenStoredProcedure(object procedureName);

		/// <summary>
		/// SupportByVersion Access 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff197412.aspx </remarks>
		/// <param name="procedureName">object procedureName</param>
		/// <param name="view">optional NetOffice.AccessApi.Enums.AcView View = 0</param>
		[CustomMethod]
		[SupportByVersion("Access", 9,10,11,12,14,15,16)]
		void OpenStoredProcedure(object procedureName, object view);

		/// <summary>
		/// SupportByVersion Access 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <param name="reportName">object reportName</param>
		/// <param name="view">optional NetOffice.AccessApi.Enums.AcView View = 0</param>
		/// <param name="filterName">optional object filterName</param>
		/// <param name="whereCondition">optional object whereCondition</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		[SupportByVersion("Access", 10,11,12,14,15,16)]
		void OpenReportOld0(object reportName, object view, object filterName, object whereCondition);

		/// <summary>
		/// SupportByVersion Access 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <param name="reportName">object reportName</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		[CustomMethod]
		[SupportByVersion("Access", 10,11,12,14,15,16)]
		void OpenReportOld0(object reportName);

		/// <summary>
		/// SupportByVersion Access 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <param name="reportName">object reportName</param>
		/// <param name="view">optional NetOffice.AccessApi.Enums.AcView View = 0</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		[CustomMethod]
		[SupportByVersion("Access", 10,11,12,14,15,16)]
		void OpenReportOld0(object reportName, object view);

		/// <summary>
		/// SupportByVersion Access 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <param name="reportName">object reportName</param>
		/// <param name="view">optional NetOffice.AccessApi.Enums.AcView View = 0</param>
		/// <param name="filterName">optional object filterName</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		[CustomMethod]
		[SupportByVersion("Access", 10,11,12,14,15,16)]
		void OpenReportOld0(object reportName, object view, object filterName);

		/// <summary>
		/// SupportByVersion Access 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <param name="objectType">NetOffice.AccessApi.Enums.AcOutputObjectType objectType</param>
		/// <param name="objectName">optional object objectName</param>
		/// <param name="outputFormat">optional object outputFormat</param>
		/// <param name="outputFile">optional object outputFile</param>
		/// <param name="autoStart">optional object autoStart</param>
		/// <param name="templateFile">optional object templateFile</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		[SupportByVersion("Access", 10,11,12,14,15,16)]
		void OutputToOld0(NetOffice.AccessApi.Enums.AcOutputObjectType objectType, object objectName, object outputFormat, object outputFile, object autoStart, object templateFile);

		/// <summary>
		/// SupportByVersion Access 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <param name="objectType">NetOffice.AccessApi.Enums.AcOutputObjectType objectType</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		[CustomMethod]
		[SupportByVersion("Access", 10,11,12,14,15,16)]
		void OutputToOld0(NetOffice.AccessApi.Enums.AcOutputObjectType objectType);

		/// <summary>
		/// SupportByVersion Access 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <param name="objectType">NetOffice.AccessApi.Enums.AcOutputObjectType objectType</param>
		/// <param name="objectName">optional object objectName</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		[CustomMethod]
		[SupportByVersion("Access", 10,11,12,14,15,16)]
		void OutputToOld0(NetOffice.AccessApi.Enums.AcOutputObjectType objectType, object objectName);

		/// <summary>
		/// SupportByVersion Access 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <param name="objectType">NetOffice.AccessApi.Enums.AcOutputObjectType objectType</param>
		/// <param name="objectName">optional object objectName</param>
		/// <param name="outputFormat">optional object outputFormat</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		[CustomMethod]
		[SupportByVersion("Access", 10,11,12,14,15,16)]
		void OutputToOld0(NetOffice.AccessApi.Enums.AcOutputObjectType objectType, object objectName, object outputFormat);

		/// <summary>
		/// SupportByVersion Access 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <param name="objectType">NetOffice.AccessApi.Enums.AcOutputObjectType objectType</param>
		/// <param name="objectName">optional object objectName</param>
		/// <param name="outputFormat">optional object outputFormat</param>
		/// <param name="outputFile">optional object outputFile</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		[CustomMethod]
		[SupportByVersion("Access", 10,11,12,14,15,16)]
		void OutputToOld0(NetOffice.AccessApi.Enums.AcOutputObjectType objectType, object objectName, object outputFormat, object outputFile);

		/// <summary>
		/// SupportByVersion Access 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <param name="objectType">NetOffice.AccessApi.Enums.AcOutputObjectType objectType</param>
		/// <param name="objectName">optional object objectName</param>
		/// <param name="outputFormat">optional object outputFormat</param>
		/// <param name="outputFile">optional object outputFile</param>
		/// <param name="autoStart">optional object autoStart</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		[CustomMethod]
		[SupportByVersion("Access", 10,11,12,14,15,16)]
		void OutputToOld0(NetOffice.AccessApi.Enums.AcOutputObjectType objectType, object objectName, object outputFormat, object outputFile, object autoStart);

		/// <summary>
		/// SupportByVersion Access 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff835048.aspx </remarks>
		/// <param name="server">object server</param>
		/// <param name="database">object database</param>
		/// <param name="useTrustedConnection">optional object useTrustedConnection</param>
		/// <param name="login">optional object login</param>
		/// <param name="password">optional object password</param>
		/// <param name="transferCopyData">optional object transferCopyData</param>
		[SupportByVersion("Access", 10,11,12,14,15,16)]
		void TransferSQLDatabase(object server, object database, object useTrustedConnection, object login, object password, object transferCopyData);

		/// <summary>
		/// SupportByVersion Access 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff835048.aspx </remarks>
		/// <param name="server">object server</param>
		/// <param name="database">object database</param>
		[CustomMethod]
		[SupportByVersion("Access", 10,11,12,14,15,16)]
		void TransferSQLDatabase(object server, object database);

		/// <summary>
		/// SupportByVersion Access 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff835048.aspx </remarks>
		/// <param name="server">object server</param>
		/// <param name="database">object database</param>
		/// <param name="useTrustedConnection">optional object useTrustedConnection</param>
		[CustomMethod]
		[SupportByVersion("Access", 10,11,12,14,15,16)]
		void TransferSQLDatabase(object server, object database, object useTrustedConnection);

		/// <summary>
		/// SupportByVersion Access 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff835048.aspx </remarks>
		/// <param name="server">object server</param>
		/// <param name="database">object database</param>
		/// <param name="useTrustedConnection">optional object useTrustedConnection</param>
		/// <param name="login">optional object login</param>
		[CustomMethod]
		[SupportByVersion("Access", 10,11,12,14,15,16)]
		void TransferSQLDatabase(object server, object database, object useTrustedConnection, object login);

		/// <summary>
		/// SupportByVersion Access 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff835048.aspx </remarks>
		/// <param name="server">object server</param>
		/// <param name="database">object database</param>
		/// <param name="useTrustedConnection">optional object useTrustedConnection</param>
		/// <param name="login">optional object login</param>
		/// <param name="password">optional object password</param>
		[CustomMethod]
		[SupportByVersion("Access", 10,11,12,14,15,16)]
		void TransferSQLDatabase(object server, object database, object useTrustedConnection, object login, object password);

		/// <summary>
		/// SupportByVersion Access 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff845497.aspx </remarks>
		/// <param name="databaseFileName">object databaseFileName</param>
		/// <param name="overwriteExistingFile">optional object overwriteExistingFile</param>
		/// <param name="disconnectAllUsers">optional object disconnectAllUsers</param>
		[SupportByVersion("Access", 10,11,12,14,15,16)]
		void CopyDatabaseFile(object databaseFileName, object overwriteExistingFile, object disconnectAllUsers);

		/// <summary>
		/// SupportByVersion Access 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff845497.aspx </remarks>
		/// <param name="databaseFileName">object databaseFileName</param>
		[CustomMethod]
		[SupportByVersion("Access", 10,11,12,14,15,16)]
		void CopyDatabaseFile(object databaseFileName);

		/// <summary>
		/// SupportByVersion Access 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff845497.aspx </remarks>
		/// <param name="databaseFileName">object databaseFileName</param>
		/// <param name="overwriteExistingFile">optional object overwriteExistingFile</param>
		[CustomMethod]
		[SupportByVersion("Access", 10,11,12,14,15,16)]
		void CopyDatabaseFile(object databaseFileName, object overwriteExistingFile);

		/// <summary>
		/// SupportByVersion Access 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff194192.aspx </remarks>
		/// <param name="functionName">object functionName</param>
		/// <param name="view">optional NetOffice.AccessApi.Enums.AcView View = 0</param>
		/// <param name="dataMode">optional NetOffice.AccessApi.Enums.AcOpenDataMode DataMode = 1</param>
		[SupportByVersion("Access", 10,11,12,14,15,16)]
		void OpenFunction(object functionName, object view, object dataMode);

		/// <summary>
		/// SupportByVersion Access 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff194192.aspx </remarks>
		/// <param name="functionName">object functionName</param>
		[CustomMethod]
		[SupportByVersion("Access", 10,11,12,14,15,16)]
		void OpenFunction(object functionName);

		/// <summary>
		/// SupportByVersion Access 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff194192.aspx </remarks>
		/// <param name="functionName">object functionName</param>
		/// <param name="view">optional NetOffice.AccessApi.Enums.AcView View = 0</param>
		[CustomMethod]
		[SupportByVersion("Access", 10,11,12,14,15,16)]
		void OpenFunction(object functionName, object view);

		/// <summary>
		/// SupportByVersion Access 12, 14, 15, 16
		/// </summary>
		/// <param name="filterName">optional object filterName</param>
		/// <param name="whereCondition">optional object whereCondition</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		[SupportByVersion("Access", 12,14,15,16)]
		void ApplyFilterOld0(object filterName, object whereCondition);

		/// <summary>
		/// SupportByVersion Access 12, 14, 15, 16
		/// </summary>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		[CustomMethod]
		[SupportByVersion("Access", 12,14,15,16)]
		void ApplyFilterOld0();

		/// <summary>
		/// SupportByVersion Access 12, 14, 15, 16
		/// </summary>
		/// <param name="filterName">optional object filterName</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		[CustomMethod]
		[SupportByVersion("Access", 12,14,15,16)]
		void ApplyFilterOld0(object filterName);

		/// <summary>
		/// SupportByVersion Access 12, 14, 15, 16
		/// </summary>
		/// <param name="objectType">NetOffice.AccessApi.Enums.AcOutputObjectType objectType</param>
		/// <param name="objectName">optional object objectName</param>
		/// <param name="outputFormat">optional object outputFormat</param>
		/// <param name="outputFile">optional object outputFile</param>
		/// <param name="autoStart">optional object autoStart</param>
		/// <param name="templateFile">optional object templateFile</param>
		/// <param name="encoding">optional object encoding</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		[SupportByVersion("Access", 12,14,15,16)]
		void OutputToOld1(NetOffice.AccessApi.Enums.AcOutputObjectType objectType, object objectName, object outputFormat, object outputFile, object autoStart, object templateFile, object encoding);

		/// <summary>
		/// SupportByVersion Access 12, 14, 15, 16
		/// </summary>
		/// <param name="objectType">NetOffice.AccessApi.Enums.AcOutputObjectType objectType</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		[CustomMethod]
		[SupportByVersion("Access", 12,14,15,16)]
		void OutputToOld1(NetOffice.AccessApi.Enums.AcOutputObjectType objectType);

		/// <summary>
		/// SupportByVersion Access 12, 14, 15, 16
		/// </summary>
		/// <param name="objectType">NetOffice.AccessApi.Enums.AcOutputObjectType objectType</param>
		/// <param name="objectName">optional object objectName</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		[CustomMethod]
		[SupportByVersion("Access", 12,14,15,16)]
		void OutputToOld1(NetOffice.AccessApi.Enums.AcOutputObjectType objectType, object objectName);

		/// <summary>
		/// SupportByVersion Access 12, 14, 15, 16
		/// </summary>
		/// <param name="objectType">NetOffice.AccessApi.Enums.AcOutputObjectType objectType</param>
		/// <param name="objectName">optional object objectName</param>
		/// <param name="outputFormat">optional object outputFormat</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		[CustomMethod]
		[SupportByVersion("Access", 12,14,15,16)]
		void OutputToOld1(NetOffice.AccessApi.Enums.AcOutputObjectType objectType, object objectName, object outputFormat);

		/// <summary>
		/// SupportByVersion Access 12, 14, 15, 16
		/// </summary>
		/// <param name="objectType">NetOffice.AccessApi.Enums.AcOutputObjectType objectType</param>
		/// <param name="objectName">optional object objectName</param>
		/// <param name="outputFormat">optional object outputFormat</param>
		/// <param name="outputFile">optional object outputFile</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		[CustomMethod]
		[SupportByVersion("Access", 12,14,15,16)]
		void OutputToOld1(NetOffice.AccessApi.Enums.AcOutputObjectType objectType, object objectName, object outputFormat, object outputFile);

		/// <summary>
		/// SupportByVersion Access 12, 14, 15, 16
		/// </summary>
		/// <param name="objectType">NetOffice.AccessApi.Enums.AcOutputObjectType objectType</param>
		/// <param name="objectName">optional object objectName</param>
		/// <param name="outputFormat">optional object outputFormat</param>
		/// <param name="outputFile">optional object outputFile</param>
		/// <param name="autoStart">optional object autoStart</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		[CustomMethod]
		[SupportByVersion("Access", 12,14,15,16)]
		void OutputToOld1(NetOffice.AccessApi.Enums.AcOutputObjectType objectType, object objectName, object outputFormat, object outputFile, object autoStart);

		/// <summary>
		/// SupportByVersion Access 12, 14, 15, 16
		/// </summary>
		/// <param name="objectType">NetOffice.AccessApi.Enums.AcOutputObjectType objectType</param>
		/// <param name="objectName">optional object objectName</param>
		/// <param name="outputFormat">optional object outputFormat</param>
		/// <param name="outputFile">optional object outputFile</param>
		/// <param name="autoStart">optional object autoStart</param>
		/// <param name="templateFile">optional object templateFile</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		[CustomMethod]
		[SupportByVersion("Access", 12,14,15,16)]
		void OutputToOld1(NetOffice.AccessApi.Enums.AcOutputObjectType objectType, object objectName, object outputFormat, object outputFile, object autoStart, object templateFile);

		/// <summary>
		/// SupportByVersion Access 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff198137.aspx </remarks>
		/// <param name="transferType">NetOffice.AccessApi.Enums.AcSharePointListTransferType transferType</param>
		/// <param name="siteAddress">object siteAddress</param>
		/// <param name="listID">object listID</param>
		/// <param name="viewID">optional object viewID</param>
		/// <param name="tableName">optional object tableName</param>
		/// <param name="getLookupDisplayValues">optional object getLookupDisplayValues</param>
		[SupportByVersion("Access", 12,14,15,16)]
		void TransferSharePointList(NetOffice.AccessApi.Enums.AcSharePointListTransferType transferType, object siteAddress, object listID, object viewID, object tableName, object getLookupDisplayValues);

		/// <summary>
		/// SupportByVersion Access 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff198137.aspx </remarks>
		/// <param name="transferType">NetOffice.AccessApi.Enums.AcSharePointListTransferType transferType</param>
		/// <param name="siteAddress">object siteAddress</param>
		/// <param name="listID">object listID</param>
		[CustomMethod]
		[SupportByVersion("Access", 12,14,15,16)]
		void TransferSharePointList(NetOffice.AccessApi.Enums.AcSharePointListTransferType transferType, object siteAddress, object listID);

		/// <summary>
		/// SupportByVersion Access 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff198137.aspx </remarks>
		/// <param name="transferType">NetOffice.AccessApi.Enums.AcSharePointListTransferType transferType</param>
		/// <param name="siteAddress">object siteAddress</param>
		/// <param name="listID">object listID</param>
		/// <param name="viewID">optional object viewID</param>
		[CustomMethod]
		[SupportByVersion("Access", 12,14,15,16)]
		void TransferSharePointList(NetOffice.AccessApi.Enums.AcSharePointListTransferType transferType, object siteAddress, object listID, object viewID);

		/// <summary>
		/// SupportByVersion Access 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff198137.aspx </remarks>
		/// <param name="transferType">NetOffice.AccessApi.Enums.AcSharePointListTransferType transferType</param>
		/// <param name="siteAddress">object siteAddress</param>
		/// <param name="listID">object listID</param>
		/// <param name="viewID">optional object viewID</param>
		/// <param name="tableName">optional object tableName</param>
		[CustomMethod]
		[SupportByVersion("Access", 12,14,15,16)]
		void TransferSharePointList(NetOffice.AccessApi.Enums.AcSharePointListTransferType transferType, object siteAddress, object listID, object viewID, object tableName);

		/// <summary>
		/// SupportByVersion Access 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff844747.aspx </remarks>
		[SupportByVersion("Access", 12,14,15,16)]
		void CloseDatabase();

		/// <summary>
		/// SupportByVersion Access 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff191916.aspx </remarks>
		/// <param name="category">optional object category</param>
		/// <param name="group">optional object group</param>
		[SupportByVersion("Access", 12,14,15,16)]
		void NavigateTo(object category, object group);

		/// <summary>
		/// SupportByVersion Access 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff191916.aspx </remarks>
		[CustomMethod]
		[SupportByVersion("Access", 12,14,15,16)]
		void NavigateTo();

		/// <summary>
		/// SupportByVersion Access 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff191916.aspx </remarks>
		/// <param name="category">optional object category</param>
		[CustomMethod]
		[SupportByVersion("Access", 12,14,15,16)]
		void NavigateTo(object category);

		/// <summary>
		/// SupportByVersion Access 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff836254.aspx </remarks>
		/// <param name="objectType">optional NetOffice.AccessApi.Enums.AcDataObjectType ObjectType = -1</param>
		/// <param name="objectName">optional object objectName</param>
		/// <param name="record">optional NetOffice.AccessApi.Enums.AcRecord Record = 2</param>
		/// <param name="whereCondition">optional object whereCondition</param>
		[SupportByVersion("Access", 12,14,15,16)]
		void SearchForRecord(object objectType, object objectName, object record, object whereCondition);

		/// <summary>
		/// SupportByVersion Access 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff836254.aspx </remarks>
		[CustomMethod]
		[SupportByVersion("Access", 12,14,15,16)]
		void SearchForRecord();

		/// <summary>
		/// SupportByVersion Access 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff836254.aspx </remarks>
		/// <param name="objectType">optional NetOffice.AccessApi.Enums.AcDataObjectType ObjectType = -1</param>
		[CustomMethod]
		[SupportByVersion("Access", 12,14,15,16)]
		void SearchForRecord(object objectType);

		/// <summary>
		/// SupportByVersion Access 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff836254.aspx </remarks>
		/// <param name="objectType">optional NetOffice.AccessApi.Enums.AcDataObjectType ObjectType = -1</param>
		/// <param name="objectName">optional object objectName</param>
		[CustomMethod]
		[SupportByVersion("Access", 12,14,15,16)]
		void SearchForRecord(object objectType, object objectName);

		/// <summary>
		/// SupportByVersion Access 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff836254.aspx </remarks>
		/// <param name="objectType">optional NetOffice.AccessApi.Enums.AcDataObjectType ObjectType = -1</param>
		/// <param name="objectName">optional object objectName</param>
		/// <param name="record">optional NetOffice.AccessApi.Enums.AcRecord Record = 2</param>
		[CustomMethod]
		[SupportByVersion("Access", 12,14,15,16)]
		void SearchForRecord(object objectType, object objectName, object record);

		/// <summary>
		/// SupportByVersion Access 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff192301.aspx </remarks>
		/// <param name="controlName">object controlName</param>
		/// <param name="property">optional NetOffice.AccessApi.Enums.AcProperty Property = 0</param>
		/// <param name="value">optional object value</param>
		[SupportByVersion("Access", 12,14,15,16)]
		void SetProperty(object controlName, object property, object value);

		/// <summary>
		/// SupportByVersion Access 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff192301.aspx </remarks>
		/// <param name="controlName">object controlName</param>
		[CustomMethod]
		[SupportByVersion("Access", 12,14,15,16)]
		void SetProperty(object controlName);

		/// <summary>
		/// SupportByVersion Access 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff192301.aspx </remarks>
		/// <param name="controlName">object controlName</param>
		/// <param name="property">optional NetOffice.AccessApi.Enums.AcProperty Property = 0</param>
		[CustomMethod]
		[SupportByVersion("Access", 12,14,15,16)]
		void SetProperty(object controlName, object property);

		/// <summary>
		/// SupportByVersion Access 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff837036.aspx </remarks>
		[SupportByVersion("Access", 12,14,15,16)]
		void SingleStep();

		/// <summary>
		/// SupportByVersion Access 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff191914.aspx </remarks>
		[SupportByVersion("Access", 12,14,15,16)]
		void ClearMacroError();

		/// <summary>
		/// SupportByVersion Access 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff821741.aspx </remarks>
		/// <param name="show">object show</param>
		/// <param name="category">optional object category</param>
		[SupportByVersion("Access", 12,14,15,16)]
		void SetDisplayedCategories(object show, object category);

		/// <summary>
		/// SupportByVersion Access 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff821741.aspx </remarks>
		/// <param name="show">object show</param>
		[CustomMethod]
		[SupportByVersion("Access", 12,14,15,16)]
		void SetDisplayedCategories(object show);

		/// <summary>
		/// SupportByVersion Access 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff195088.aspx </remarks>
		/// <param name="_lock">object lock</param>
		[SupportByVersion("Access", 12,14,15,16)]
		void LockNavigationPane(object _lock);

		/// <summary>
		/// SupportByVersion Access 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff834375.aspx </remarks>
		/// <param name="savedImportExportName">object savedImportExportName</param>
		[SupportByVersion("Access", 12,14,15,16)]
		void RunSavedImportExport(object savedImportExportName);

		/// <summary>
		/// SupportByVersion Access 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff196381.aspx </remarks>
		/// <param name="objectType">NetOffice.AccessApi.Enums.AcBrowseToObjectType objectType</param>
		/// <param name="objectName">object objectName</param>
		/// <param name="pathtoSubformControl">optional object pathtoSubformControl</param>
		/// <param name="whereCondition">optional object whereCondition</param>
		/// <param name="page">optional object page</param>
		/// <param name="dataMode">optional NetOffice.AccessApi.Enums.AcFormOpenDataMode DataMode = 1</param>
		[SupportByVersion("Access", 14,15,16)]
		void BrowseTo(NetOffice.AccessApi.Enums.AcBrowseToObjectType objectType, object objectName, object pathtoSubformControl, object whereCondition, object page, object dataMode);

		/// <summary>
		/// SupportByVersion Access 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff196381.aspx </remarks>
		/// <param name="objectType">NetOffice.AccessApi.Enums.AcBrowseToObjectType objectType</param>
		/// <param name="objectName">object objectName</param>
		[CustomMethod]
		[SupportByVersion("Access", 14,15,16)]
		void BrowseTo(NetOffice.AccessApi.Enums.AcBrowseToObjectType objectType, object objectName);

		/// <summary>
		/// SupportByVersion Access 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff196381.aspx </remarks>
		/// <param name="objectType">NetOffice.AccessApi.Enums.AcBrowseToObjectType objectType</param>
		/// <param name="objectName">object objectName</param>
		/// <param name="pathtoSubformControl">optional object pathtoSubformControl</param>
		[CustomMethod]
		[SupportByVersion("Access", 14,15,16)]
		void BrowseTo(NetOffice.AccessApi.Enums.AcBrowseToObjectType objectType, object objectName, object pathtoSubformControl);

		/// <summary>
		/// SupportByVersion Access 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff196381.aspx </remarks>
		/// <param name="objectType">NetOffice.AccessApi.Enums.AcBrowseToObjectType objectType</param>
		/// <param name="objectName">object objectName</param>
		/// <param name="pathtoSubformControl">optional object pathtoSubformControl</param>
		/// <param name="whereCondition">optional object whereCondition</param>
		[CustomMethod]
		[SupportByVersion("Access", 14,15,16)]
		void BrowseTo(NetOffice.AccessApi.Enums.AcBrowseToObjectType objectType, object objectName, object pathtoSubformControl, object whereCondition);

		/// <summary>
		/// SupportByVersion Access 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff196381.aspx </remarks>
		/// <param name="objectType">NetOffice.AccessApi.Enums.AcBrowseToObjectType objectType</param>
		/// <param name="objectName">object objectName</param>
		/// <param name="pathtoSubformControl">optional object pathtoSubformControl</param>
		/// <param name="whereCondition">optional object whereCondition</param>
		/// <param name="page">optional object page</param>
		[CustomMethod]
		[SupportByVersion("Access", 14,15,16)]
		void BrowseTo(NetOffice.AccessApi.Enums.AcBrowseToObjectType objectType, object objectName, object pathtoSubformControl, object whereCondition, object page);

		/// <summary>
		/// SupportByVersion Access 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff194182.aspx </remarks>
		/// <param name="name">object name</param>
		/// <param name="expression">object expression</param>
		[SupportByVersion("Access", 14,15,16)]
		void SetParameter(object name, object expression);

		/// <summary>
		/// SupportByVersion Access 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff836068.aspx </remarks>
		/// <param name="macroName">object macroName</param>
		[SupportByVersion("Access", 14,15,16)]
		void RunDataMacro(object macroName);

		/// <summary>
		/// SupportByVersion Access 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff844761.aspx </remarks>
		/// <param name="orderBy">object orderBy</param>
		/// <param name="controlName">optional object controlName</param>
		[SupportByVersion("Access", 14,15,16)]
		void SetOrderBy(object orderBy, object controlName);

		/// <summary>
		/// SupportByVersion Access 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff844761.aspx </remarks>
		/// <param name="orderBy">object orderBy</param>
		[CustomMethod]
		[SupportByVersion("Access", 14,15,16)]
		void SetOrderBy(object orderBy);

		/// <summary>
		/// SupportByVersion Access 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff197950.aspx </remarks>
		/// <param name="filterName">optional object filterName</param>
		/// <param name="whereCondition">optional object whereCondition</param>
		/// <param name="controlName">optional object controlName</param>
		[SupportByVersion("Access", 14,15,16)]
		void SetFilter(object filterName, object whereCondition, object controlName);

		/// <summary>
		/// SupportByVersion Access 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff197950.aspx </remarks>
		[CustomMethod]
		[SupportByVersion("Access", 14,15,16)]
		void SetFilter();

		/// <summary>
		/// SupportByVersion Access 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff197950.aspx </remarks>
		/// <param name="filterName">optional object filterName</param>
		[CustomMethod]
		[SupportByVersion("Access", 14,15,16)]
		void SetFilter(object filterName);

		/// <summary>
		/// SupportByVersion Access 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff197950.aspx </remarks>
		/// <param name="filterName">optional object filterName</param>
		/// <param name="whereCondition">optional object whereCondition</param>
		[CustomMethod]
		[SupportByVersion("Access", 14,15,16)]
		void SetFilter(object filterName, object whereCondition);

		/// <summary>
		/// SupportByVersion Access 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff191907.aspx </remarks>
		[SupportByVersion("Access", 14,15,16)]
		void RefreshRecord();

		#endregion
	}
}
