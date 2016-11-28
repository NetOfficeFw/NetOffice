using System;
using NetRuntimeSystem = System;
using System.Runtime.InteropServices;
using System.Runtime.CompilerServices;
using System.ComponentModel;
using System.Reflection;
using System.Collections.Generic;
using NetOffice;
namespace NetOffice.AccessApi
{
	///<summary>
	/// DispatchInterface DoCmd 
	/// SupportByVersion Access, 9,10,11,12,14,15,16
	/// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff192694.aspx
	///</summary>
	[SupportByVersionAttribute("Access", 9,10,11,12,14,15,16)]
	[EntityTypeAttribute(EntityType.IsDispatchInterface)]
	public class DoCmd : COMObject
	{
		#pragma warning disable
		#region Type Information

        private static Type _type;

		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
        public static Type LateBindingApiWrapperType
        {
            get
            {
                if (null == _type)
                    _type = typeof(DoCmd);
                    
                return _type;
            }
        }
        
        #endregion
        
		#region Construction

		///<param name="factory">current used factory core</param>
		///<param name="parentObject">object there has created the proxy</param>
        ///<param name="comProxy">inner wrapped COM proxy</param>
		public DoCmd(Core factory, ICOMObject parentObject, object comProxy) : base(factory, parentObject, comProxy)
		{
			
		}

        ///<param name="parentObject">object there has created the proxy</param>
        ///<param name="comProxy">inner wrapped COM proxy</param>
        [EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public DoCmd(ICOMObject parentObject, object comProxy) : base(parentObject, comProxy)
		{
		}
		
		///<param name="factory">current used factory core</param>
		///<param name="parentObject">object there has created the proxy</param>
        ///<param name="comProxy">inner wrapped COM proxy</param>
        ///<param name="comProxyType">Type of inner wrapped COM proxy"</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public DoCmd(Core factory, ICOMObject parentObject, object comProxy, NetRuntimeSystem.Type comProxyType) : base(factory, parentObject, comProxy, comProxyType)
		{

		}

		///<param name="parentObject">object there has created the proxy</param>
        ///<param name="comProxy">inner wrapped COM proxy</param>
        ///<param name="comProxyType">Type of inner wrapped COM proxy"</param>
        [EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public DoCmd(ICOMObject parentObject, object comProxy, NetRuntimeSystem.Type comProxyType) : base(parentObject, comProxy, comProxyType)
		{
		}
		
		///<param name="replacedObject">object to replaced. replacedObject are not usable after this action</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public DoCmd(ICOMObject replacedObject) : base(replacedObject)
		{
		}
		
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public DoCmd() : base()
		{
		}
		
		/// <param name="progId">registered ProgID</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public DoCmd(string progId) : base(progId)
		{
		}
		
		#endregion
		
		#region Properties

		#endregion

		#region Methods

		/// <summary>
		/// SupportByVersion Access 9, 10, 11, 12, 14, 15, 16
		/// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff834781.aspx
		/// </summary>
		/// <param name="menuName">object MenuName</param>
		/// <param name="menuMacroName">object MenuMacroName</param>
		/// <param name="statusBarText">object StatusBarText</param>
		[SupportByVersionAttribute("Access", 9,10,11,12,14,15,16)]
		public void AddMenu(object menuName, object menuMacroName, object statusBarText)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(menuName, menuMacroName, statusBarText);
			Invoker.Method(this, "AddMenu", paramsArray);
		}

		/// <summary>
		/// SupportByVersion Access 9, 10, 11, 12, 14, 15, 16
		/// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff197651.aspx
		/// </summary>
		/// <param name="filterName">optional object FilterName</param>
		/// <param name="whereCondition">optional object WhereCondition</param>
		[SupportByVersionAttribute("Access", 9,10,11,12,14,15,16)]
		public void ApplyFilter(object filterName, object whereCondition)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(filterName, whereCondition);
			Invoker.Method(this, "ApplyFilter", paramsArray);
		}

		/// <summary>
		/// SupportByVersion Access 12, 14, 15, 16
		/// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff197651.aspx
		/// </summary>
		/// <param name="filterName">optional object FilterName</param>
		/// <param name="whereCondition">optional object WhereCondition</param>
		/// <param name="controlName">optional object ControlName</param>
		[SupportByVersionAttribute("Access", 12,14,15,16)]
		public void ApplyFilter(object filterName, object whereCondition, object controlName)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(filterName, whereCondition, controlName);
			Invoker.Method(this, "ApplyFilter", paramsArray);
		}

		/// <summary>
		/// SupportByVersion Access 9, 10, 11, 12, 14, 15, 16
		/// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff197651.aspx
		/// </summary>
		[CustomMethodAttribute]
		[SupportByVersionAttribute("Access", 9,10,11,12,14,15,16)]
		public void ApplyFilter()
		{
			object[] paramsArray = null;
			Invoker.Method(this, "ApplyFilter", paramsArray);
		}

		/// <summary>
		/// SupportByVersion Access 9, 10, 11, 12, 14, 15, 16
		/// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff197651.aspx
		/// </summary>
		/// <param name="filterName">optional object FilterName</param>
		[CustomMethodAttribute]
		[SupportByVersionAttribute("Access", 9,10,11,12,14,15,16)]
		public void ApplyFilter(object filterName)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(filterName);
			Invoker.Method(this, "ApplyFilter", paramsArray);
		}

		/// <summary>
		/// SupportByVersion Access 9, 10, 11, 12, 14, 15, 16
		/// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff196680.aspx
		/// </summary>
		[SupportByVersionAttribute("Access", 9,10,11,12,14,15,16)]
		public void Beep()
		{
			object[] paramsArray = null;
			Invoker.Method(this, "Beep", paramsArray);
		}

		/// <summary>
		/// SupportByVersion Access 9, 10, 11, 12, 14, 15, 16
		/// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff836964.aspx
		/// </summary>
		[SupportByVersionAttribute("Access", 9,10,11,12,14,15,16)]
		public void CancelEvent()
		{
			object[] paramsArray = null;
			Invoker.Method(this, "CancelEvent", paramsArray);
		}

		/// <summary>
		/// SupportByVersion Access 9, 10, 11, 12, 14, 15, 16
		/// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff192860.aspx
		/// </summary>
		/// <param name="objectType">optional NetOffice.AccessApi.Enums.AcObjectType ObjectType = -1</param>
		/// <param name="objectName">optional object ObjectName</param>
		/// <param name="save">optional NetOffice.AccessApi.Enums.AcCloseSave Save = 0</param>
		[SupportByVersionAttribute("Access", 9,10,11,12,14,15,16)]
		public void Close(object objectType, object objectName, object save)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(objectType, objectName, save);
			Invoker.Method(this, "Close", paramsArray);
		}

		/// <summary>
		/// SupportByVersion Access 9, 10, 11, 12, 14, 15, 16
		/// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff192860.aspx
		/// </summary>
		[CustomMethodAttribute]
		[SupportByVersionAttribute("Access", 9,10,11,12,14,15,16)]
		public void Close()
		{
			object[] paramsArray = null;
			Invoker.Method(this, "Close", paramsArray);
		}

		/// <summary>
		/// SupportByVersion Access 9, 10, 11, 12, 14, 15, 16
		/// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff192860.aspx
		/// </summary>
		/// <param name="objectType">optional NetOffice.AccessApi.Enums.AcObjectType ObjectType = -1</param>
		[CustomMethodAttribute]
		[SupportByVersionAttribute("Access", 9,10,11,12,14,15,16)]
		public void Close(object objectType)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(objectType);
			Invoker.Method(this, "Close", paramsArray);
		}

		/// <summary>
		/// SupportByVersion Access 9, 10, 11, 12, 14, 15, 16
		/// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff192860.aspx
		/// </summary>
		/// <param name="objectType">optional NetOffice.AccessApi.Enums.AcObjectType ObjectType = -1</param>
		/// <param name="objectName">optional object ObjectName</param>
		[CustomMethodAttribute]
		[SupportByVersionAttribute("Access", 9,10,11,12,14,15,16)]
		public void Close(object objectType, object objectName)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(objectType, objectName);
			Invoker.Method(this, "Close", paramsArray);
		}

		/// <summary>
		/// SupportByVersion Access 9, 10, 11, 12, 14, 15, 16
		/// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff844724.aspx
		/// </summary>
		/// <param name="destinationDatabase">optional object DestinationDatabase</param>
		/// <param name="newName">optional object NewName</param>
		/// <param name="sourceObjectType">optional NetOffice.AccessApi.Enums.AcObjectType SourceObjectType = -1</param>
		/// <param name="sourceObjectName">optional object SourceObjectName</param>
		[SupportByVersionAttribute("Access", 9,10,11,12,14,15,16)]
		public void CopyObject(object destinationDatabase, object newName, object sourceObjectType, object sourceObjectName)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(destinationDatabase, newName, sourceObjectType, sourceObjectName);
			Invoker.Method(this, "CopyObject", paramsArray);
		}

		/// <summary>
		/// SupportByVersion Access 9, 10, 11, 12, 14, 15, 16
		/// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff844724.aspx
		/// </summary>
		[CustomMethodAttribute]
		[SupportByVersionAttribute("Access", 9,10,11,12,14,15,16)]
		public void CopyObject()
		{
			object[] paramsArray = null;
			Invoker.Method(this, "CopyObject", paramsArray);
		}

		/// <summary>
		/// SupportByVersion Access 9, 10, 11, 12, 14, 15, 16
		/// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff844724.aspx
		/// </summary>
		/// <param name="destinationDatabase">optional object DestinationDatabase</param>
		[CustomMethodAttribute]
		[SupportByVersionAttribute("Access", 9,10,11,12,14,15,16)]
		public void CopyObject(object destinationDatabase)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(destinationDatabase);
			Invoker.Method(this, "CopyObject", paramsArray);
		}

		/// <summary>
		/// SupportByVersion Access 9, 10, 11, 12, 14, 15, 16
		/// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff844724.aspx
		/// </summary>
		/// <param name="destinationDatabase">optional object DestinationDatabase</param>
		/// <param name="newName">optional object NewName</param>
		[CustomMethodAttribute]
		[SupportByVersionAttribute("Access", 9,10,11,12,14,15,16)]
		public void CopyObject(object destinationDatabase, object newName)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(destinationDatabase, newName);
			Invoker.Method(this, "CopyObject", paramsArray);
		}

		/// <summary>
		/// SupportByVersion Access 9, 10, 11, 12, 14, 15, 16
		/// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff844724.aspx
		/// </summary>
		/// <param name="destinationDatabase">optional object DestinationDatabase</param>
		/// <param name="newName">optional object NewName</param>
		/// <param name="sourceObjectType">optional NetOffice.AccessApi.Enums.AcObjectType SourceObjectType = -1</param>
		[CustomMethodAttribute]
		[SupportByVersionAttribute("Access", 9,10,11,12,14,15,16)]
		public void CopyObject(object destinationDatabase, object newName, object sourceObjectType)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(destinationDatabase, newName, sourceObjectType);
			Invoker.Method(this, "CopyObject", paramsArray);
		}

		/// <summary>
		/// SupportByVersion Access 9, 10, 11, 12, 14, 15, 16
		/// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff822447.aspx
		/// </summary>
		/// <param name="menuBar">object MenuBar</param>
		/// <param name="menuName">object MenuName</param>
		/// <param name="command">object Command</param>
		/// <param name="subcommand">optional object Subcommand</param>
		/// <param name="version">optional object Version</param>
		[SupportByVersionAttribute("Access", 9,10,11,12,14,15,16)]
		public void DoMenuItem(object menuBar, object menuName, object command, object subcommand, object version)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(menuBar, menuName, command, subcommand, version);
			Invoker.Method(this, "DoMenuItem", paramsArray);
		}

		/// <summary>
		/// SupportByVersion Access 9, 10, 11, 12, 14, 15, 16
		/// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff822447.aspx
		/// </summary>
		/// <param name="menuBar">object MenuBar</param>
		/// <param name="menuName">object MenuName</param>
		/// <param name="command">object Command</param>
		[CustomMethodAttribute]
		[SupportByVersionAttribute("Access", 9,10,11,12,14,15,16)]
		public void DoMenuItem(object menuBar, object menuName, object command)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(menuBar, menuName, command);
			Invoker.Method(this, "DoMenuItem", paramsArray);
		}

		/// <summary>
		/// SupportByVersion Access 9, 10, 11, 12, 14, 15, 16
		/// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff822447.aspx
		/// </summary>
		/// <param name="menuBar">object MenuBar</param>
		/// <param name="menuName">object MenuName</param>
		/// <param name="command">object Command</param>
		/// <param name="subcommand">optional object Subcommand</param>
		[CustomMethodAttribute]
		[SupportByVersionAttribute("Access", 9,10,11,12,14,15,16)]
		public void DoMenuItem(object menuBar, object menuName, object command, object subcommand)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(menuBar, menuName, command, subcommand);
			Invoker.Method(this, "DoMenuItem", paramsArray);
		}

		/// <summary>
		/// SupportByVersion Access 9, 10, 11, 12, 14, 15, 16
		/// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff193863.aspx
		/// </summary>
		/// <param name="echoOn">object EchoOn</param>
		/// <param name="statusBarText">optional object StatusBarText</param>
		[SupportByVersionAttribute("Access", 9,10,11,12,14,15,16)]
		public void Echo(object echoOn, object statusBarText)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(echoOn, statusBarText);
			Invoker.Method(this, "Echo", paramsArray);
		}

		/// <summary>
		/// SupportByVersion Access 9, 10, 11, 12, 14, 15, 16
		/// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff193863.aspx
		/// </summary>
		/// <param name="echoOn">object EchoOn</param>
		[CustomMethodAttribute]
		[SupportByVersionAttribute("Access", 9,10,11,12,14,15,16)]
		public void Echo(object echoOn)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(echoOn);
			Invoker.Method(this, "Echo", paramsArray);
		}

		/// <summary>
		/// SupportByVersion Access 9, 10, 11, 12, 14, 15, 16
		/// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff196453.aspx
		/// </summary>
		[SupportByVersionAttribute("Access", 9,10,11,12,14,15,16)]
		public void FindNext()
		{
			object[] paramsArray = null;
			Invoker.Method(this, "FindNext", paramsArray);
		}

		/// <summary>
		/// SupportByVersion Access 9, 10, 11, 12, 14, 15, 16
		/// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff835361.aspx
		/// </summary>
		/// <param name="findWhat">object FindWhat</param>
		/// <param name="match">optional NetOffice.AccessApi.Enums.AcFindMatch Match = 1</param>
		/// <param name="matchCase">optional object MatchCase</param>
		/// <param name="search">optional NetOffice.AccessApi.Enums.AcSearchDirection Search = 2</param>
		/// <param name="searchAsFormatted">optional object SearchAsFormatted</param>
		/// <param name="onlyCurrentField">optional NetOffice.AccessApi.Enums.AcFindField OnlyCurrentField = -1</param>
		/// <param name="findFirst">optional object FindFirst</param>
		[SupportByVersionAttribute("Access", 9,10,11,12,14,15,16)]
		public void FindRecord(object findWhat, object match, object matchCase, object search, object searchAsFormatted, object onlyCurrentField, object findFirst)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(findWhat, match, matchCase, search, searchAsFormatted, onlyCurrentField, findFirst);
			Invoker.Method(this, "FindRecord", paramsArray);
		}

		/// <summary>
		/// SupportByVersion Access 9, 10, 11, 12, 14, 15, 16
		/// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff835361.aspx
		/// </summary>
		/// <param name="findWhat">object FindWhat</param>
		[CustomMethodAttribute]
		[SupportByVersionAttribute("Access", 9,10,11,12,14,15,16)]
		public void FindRecord(object findWhat)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(findWhat);
			Invoker.Method(this, "FindRecord", paramsArray);
		}

		/// <summary>
		/// SupportByVersion Access 9, 10, 11, 12, 14, 15, 16
		/// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff835361.aspx
		/// </summary>
		/// <param name="findWhat">object FindWhat</param>
		/// <param name="match">optional NetOffice.AccessApi.Enums.AcFindMatch Match = 1</param>
		[CustomMethodAttribute]
		[SupportByVersionAttribute("Access", 9,10,11,12,14,15,16)]
		public void FindRecord(object findWhat, object match)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(findWhat, match);
			Invoker.Method(this, "FindRecord", paramsArray);
		}

		/// <summary>
		/// SupportByVersion Access 9, 10, 11, 12, 14, 15, 16
		/// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff835361.aspx
		/// </summary>
		/// <param name="findWhat">object FindWhat</param>
		/// <param name="match">optional NetOffice.AccessApi.Enums.AcFindMatch Match = 1</param>
		/// <param name="matchCase">optional object MatchCase</param>
		[CustomMethodAttribute]
		[SupportByVersionAttribute("Access", 9,10,11,12,14,15,16)]
		public void FindRecord(object findWhat, object match, object matchCase)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(findWhat, match, matchCase);
			Invoker.Method(this, "FindRecord", paramsArray);
		}

		/// <summary>
		/// SupportByVersion Access 9, 10, 11, 12, 14, 15, 16
		/// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff835361.aspx
		/// </summary>
		/// <param name="findWhat">object FindWhat</param>
		/// <param name="match">optional NetOffice.AccessApi.Enums.AcFindMatch Match = 1</param>
		/// <param name="matchCase">optional object MatchCase</param>
		/// <param name="search">optional NetOffice.AccessApi.Enums.AcSearchDirection Search = 2</param>
		[CustomMethodAttribute]
		[SupportByVersionAttribute("Access", 9,10,11,12,14,15,16)]
		public void FindRecord(object findWhat, object match, object matchCase, object search)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(findWhat, match, matchCase, search);
			Invoker.Method(this, "FindRecord", paramsArray);
		}

		/// <summary>
		/// SupportByVersion Access 9, 10, 11, 12, 14, 15, 16
		/// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff835361.aspx
		/// </summary>
		/// <param name="findWhat">object FindWhat</param>
		/// <param name="match">optional NetOffice.AccessApi.Enums.AcFindMatch Match = 1</param>
		/// <param name="matchCase">optional object MatchCase</param>
		/// <param name="search">optional NetOffice.AccessApi.Enums.AcSearchDirection Search = 2</param>
		/// <param name="searchAsFormatted">optional object SearchAsFormatted</param>
		[CustomMethodAttribute]
		[SupportByVersionAttribute("Access", 9,10,11,12,14,15,16)]
		public void FindRecord(object findWhat, object match, object matchCase, object search, object searchAsFormatted)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(findWhat, match, matchCase, search, searchAsFormatted);
			Invoker.Method(this, "FindRecord", paramsArray);
		}

		/// <summary>
		/// SupportByVersion Access 9, 10, 11, 12, 14, 15, 16
		/// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff835361.aspx
		/// </summary>
		/// <param name="findWhat">object FindWhat</param>
		/// <param name="match">optional NetOffice.AccessApi.Enums.AcFindMatch Match = 1</param>
		/// <param name="matchCase">optional object MatchCase</param>
		/// <param name="search">optional NetOffice.AccessApi.Enums.AcSearchDirection Search = 2</param>
		/// <param name="searchAsFormatted">optional object SearchAsFormatted</param>
		/// <param name="onlyCurrentField">optional NetOffice.AccessApi.Enums.AcFindField OnlyCurrentField = -1</param>
		[CustomMethodAttribute]
		[SupportByVersionAttribute("Access", 9,10,11,12,14,15,16)]
		public void FindRecord(object findWhat, object match, object matchCase, object search, object searchAsFormatted, object onlyCurrentField)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(findWhat, match, matchCase, search, searchAsFormatted, onlyCurrentField);
			Invoker.Method(this, "FindRecord", paramsArray);
		}

		/// <summary>
		/// SupportByVersion Access 9, 10, 11, 12, 14, 15, 16
		/// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff192079.aspx
		/// </summary>
		/// <param name="controlName">object ControlName</param>
		[SupportByVersionAttribute("Access", 9,10,11,12,14,15,16)]
		public void GoToControl(object controlName)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(controlName);
			Invoker.Method(this, "GoToControl", paramsArray);
		}

		/// <summary>
		/// SupportByVersion Access 9, 10, 11, 12, 14, 15, 16
		/// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff192504.aspx
		/// </summary>
		/// <param name="pageNumber">object PageNumber</param>
		/// <param name="right">optional object Right</param>
		/// <param name="down">optional object Down</param>
		[SupportByVersionAttribute("Access", 9,10,11,12,14,15,16)]
		public void GoToPage(object pageNumber, object right, object down)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(pageNumber, right, down);
			Invoker.Method(this, "GoToPage", paramsArray);
		}

		/// <summary>
		/// SupportByVersion Access 9, 10, 11, 12, 14, 15, 16
		/// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff192504.aspx
		/// </summary>
		/// <param name="pageNumber">object PageNumber</param>
		[CustomMethodAttribute]
		[SupportByVersionAttribute("Access", 9,10,11,12,14,15,16)]
		public void GoToPage(object pageNumber)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(pageNumber);
			Invoker.Method(this, "GoToPage", paramsArray);
		}

		/// <summary>
		/// SupportByVersion Access 9, 10, 11, 12, 14, 15, 16
		/// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff192504.aspx
		/// </summary>
		/// <param name="pageNumber">object PageNumber</param>
		/// <param name="right">optional object Right</param>
		[CustomMethodAttribute]
		[SupportByVersionAttribute("Access", 9,10,11,12,14,15,16)]
		public void GoToPage(object pageNumber, object right)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(pageNumber, right);
			Invoker.Method(this, "GoToPage", paramsArray);
		}

		/// <summary>
		/// SupportByVersion Access 9, 10, 11, 12, 14, 15, 16
		/// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff194117.aspx
		/// </summary>
		/// <param name="objectType">optional NetOffice.AccessApi.Enums.AcDataObjectType ObjectType = -1</param>
		/// <param name="objectName">optional object ObjectName</param>
		/// <param name="record">optional NetOffice.AccessApi.Enums.AcRecord Record = 1</param>
		/// <param name="offset">optional object Offset</param>
		[SupportByVersionAttribute("Access", 9,10,11,12,14,15,16)]
		public void GoToRecord(object objectType, object objectName, object record, object offset)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(objectType, objectName, record, offset);
			Invoker.Method(this, "GoToRecord", paramsArray);
		}

		/// <summary>
		/// SupportByVersion Access 9, 10, 11, 12, 14, 15, 16
		/// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff194117.aspx
		/// </summary>
		[CustomMethodAttribute]
		[SupportByVersionAttribute("Access", 9,10,11,12,14,15,16)]
		public void GoToRecord()
		{
			object[] paramsArray = null;
			Invoker.Method(this, "GoToRecord", paramsArray);
		}

		/// <summary>
		/// SupportByVersion Access 9, 10, 11, 12, 14, 15, 16
		/// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff194117.aspx
		/// </summary>
		/// <param name="objectType">optional NetOffice.AccessApi.Enums.AcDataObjectType ObjectType = -1</param>
		[CustomMethodAttribute]
		[SupportByVersionAttribute("Access", 9,10,11,12,14,15,16)]
		public void GoToRecord(object objectType)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(objectType);
			Invoker.Method(this, "GoToRecord", paramsArray);
		}

		/// <summary>
		/// SupportByVersion Access 9, 10, 11, 12, 14, 15, 16
		/// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff194117.aspx
		/// </summary>
		/// <param name="objectType">optional NetOffice.AccessApi.Enums.AcDataObjectType ObjectType = -1</param>
		/// <param name="objectName">optional object ObjectName</param>
		[CustomMethodAttribute]
		[SupportByVersionAttribute("Access", 9,10,11,12,14,15,16)]
		public void GoToRecord(object objectType, object objectName)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(objectType, objectName);
			Invoker.Method(this, "GoToRecord", paramsArray);
		}

		/// <summary>
		/// SupportByVersion Access 9, 10, 11, 12, 14, 15, 16
		/// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff194117.aspx
		/// </summary>
		/// <param name="objectType">optional NetOffice.AccessApi.Enums.AcDataObjectType ObjectType = -1</param>
		/// <param name="objectName">optional object ObjectName</param>
		/// <param name="record">optional NetOffice.AccessApi.Enums.AcRecord Record = 1</param>
		[CustomMethodAttribute]
		[SupportByVersionAttribute("Access", 9,10,11,12,14,15,16)]
		public void GoToRecord(object objectType, object objectName, object record)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(objectType, objectName, record);
			Invoker.Method(this, "GoToRecord", paramsArray);
		}

		/// <summary>
		/// SupportByVersion Access 9, 10, 11, 12, 14, 15, 16
		/// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff835648.aspx
		/// </summary>
		/// <param name="hourglassOn">object HourglassOn</param>
		[SupportByVersionAttribute("Access", 9,10,11,12,14,15,16)]
		public void Hourglass(object hourglassOn)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(hourglassOn);
			Invoker.Method(this, "Hourglass", paramsArray);
		}

		/// <summary>
		/// SupportByVersion Access 9, 10, 11, 12, 14, 15, 16
		/// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff195449.aspx
		/// </summary>
		[SupportByVersionAttribute("Access", 9,10,11,12,14,15,16)]
		public void Maximize()
		{
			object[] paramsArray = null;
			Invoker.Method(this, "Maximize", paramsArray);
		}

		/// <summary>
		/// SupportByVersion Access 9, 10, 11, 12, 14, 15, 16
		/// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff837032.aspx
		/// </summary>
		[SupportByVersionAttribute("Access", 9,10,11,12,14,15,16)]
		public void Minimize()
		{
			object[] paramsArray = null;
			Invoker.Method(this, "Minimize", paramsArray);
		}

		/// <summary>
		/// SupportByVersion Access 9, 10, 11, 12, 14, 15, 16
		/// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff197394.aspx
		/// </summary>
		/// <param name="right">optional object Right</param>
		/// <param name="down">optional object Down</param>
		/// <param name="width">optional object Width</param>
		/// <param name="height">optional object Height</param>
		[SupportByVersionAttribute("Access", 9,10,11,12,14,15,16)]
		public void MoveSize(object right, object down, object width, object height)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(right, down, width, height);
			Invoker.Method(this, "MoveSize", paramsArray);
		}

		/// <summary>
		/// SupportByVersion Access 9, 10, 11, 12, 14, 15, 16
		/// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff197394.aspx
		/// </summary>
		[CustomMethodAttribute]
		[SupportByVersionAttribute("Access", 9,10,11,12,14,15,16)]
		public void MoveSize()
		{
			object[] paramsArray = null;
			Invoker.Method(this, "MoveSize", paramsArray);
		}

		/// <summary>
		/// SupportByVersion Access 9, 10, 11, 12, 14, 15, 16
		/// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff197394.aspx
		/// </summary>
		/// <param name="right">optional object Right</param>
		[CustomMethodAttribute]
		[SupportByVersionAttribute("Access", 9,10,11,12,14,15,16)]
		public void MoveSize(object right)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(right);
			Invoker.Method(this, "MoveSize", paramsArray);
		}

		/// <summary>
		/// SupportByVersion Access 9, 10, 11, 12, 14, 15, 16
		/// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff197394.aspx
		/// </summary>
		/// <param name="right">optional object Right</param>
		/// <param name="down">optional object Down</param>
		[CustomMethodAttribute]
		[SupportByVersionAttribute("Access", 9,10,11,12,14,15,16)]
		public void MoveSize(object right, object down)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(right, down);
			Invoker.Method(this, "MoveSize", paramsArray);
		}

		/// <summary>
		/// SupportByVersion Access 9, 10, 11, 12, 14, 15, 16
		/// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff197394.aspx
		/// </summary>
		/// <param name="right">optional object Right</param>
		/// <param name="down">optional object Down</param>
		/// <param name="width">optional object Width</param>
		[CustomMethodAttribute]
		[SupportByVersionAttribute("Access", 9,10,11,12,14,15,16)]
		public void MoveSize(object right, object down, object width)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(right, down, width);
			Invoker.Method(this, "MoveSize", paramsArray);
		}

		/// <summary>
		/// SupportByVersion Access 9, 10, 11, 12, 14, 15, 16
		/// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff820845.aspx
		/// </summary>
		/// <param name="formName">object FormName</param>
		/// <param name="view">optional NetOffice.AccessApi.Enums.AcFormView View = 0</param>
		/// <param name="filterName">optional object FilterName</param>
		/// <param name="whereCondition">optional object WhereCondition</param>
		/// <param name="dataMode">optional NetOffice.AccessApi.Enums.AcFormOpenDataMode DataMode = -1</param>
		/// <param name="windowMode">optional NetOffice.AccessApi.Enums.AcWindowMode WindowMode = 0</param>
		/// <param name="openArgs">optional object OpenArgs</param>
		[SupportByVersionAttribute("Access", 9,10,11,12,14,15,16)]
		public void OpenForm(object formName, object view, object filterName, object whereCondition, object dataMode, object windowMode, object openArgs)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(formName, view, filterName, whereCondition, dataMode, windowMode, openArgs);
			Invoker.Method(this, "OpenForm", paramsArray);
		}

		/// <summary>
		/// SupportByVersion Access 9, 10, 11, 12, 14, 15, 16
		/// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff820845.aspx
		/// </summary>
		/// <param name="formName">object FormName</param>
		[CustomMethodAttribute]
		[SupportByVersionAttribute("Access", 9,10,11,12,14,15,16)]
		public void OpenForm(object formName)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(formName);
			Invoker.Method(this, "OpenForm", paramsArray);
		}

		/// <summary>
		/// SupportByVersion Access 9, 10, 11, 12, 14, 15, 16
		/// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff820845.aspx
		/// </summary>
		/// <param name="formName">object FormName</param>
		/// <param name="view">optional NetOffice.AccessApi.Enums.AcFormView View = 0</param>
		[CustomMethodAttribute]
		[SupportByVersionAttribute("Access", 9,10,11,12,14,15,16)]
		public void OpenForm(object formName, object view)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(formName, view);
			Invoker.Method(this, "OpenForm", paramsArray);
		}

		/// <summary>
		/// SupportByVersion Access 9, 10, 11, 12, 14, 15, 16
		/// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff820845.aspx
		/// </summary>
		/// <param name="formName">object FormName</param>
		/// <param name="view">optional NetOffice.AccessApi.Enums.AcFormView View = 0</param>
		/// <param name="filterName">optional object FilterName</param>
		[CustomMethodAttribute]
		[SupportByVersionAttribute("Access", 9,10,11,12,14,15,16)]
		public void OpenForm(object formName, object view, object filterName)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(formName, view, filterName);
			Invoker.Method(this, "OpenForm", paramsArray);
		}

		/// <summary>
		/// SupportByVersion Access 9, 10, 11, 12, 14, 15, 16
		/// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff820845.aspx
		/// </summary>
		/// <param name="formName">object FormName</param>
		/// <param name="view">optional NetOffice.AccessApi.Enums.AcFormView View = 0</param>
		/// <param name="filterName">optional object FilterName</param>
		/// <param name="whereCondition">optional object WhereCondition</param>
		[CustomMethodAttribute]
		[SupportByVersionAttribute("Access", 9,10,11,12,14,15,16)]
		public void OpenForm(object formName, object view, object filterName, object whereCondition)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(formName, view, filterName, whereCondition);
			Invoker.Method(this, "OpenForm", paramsArray);
		}

		/// <summary>
		/// SupportByVersion Access 9, 10, 11, 12, 14, 15, 16
		/// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff820845.aspx
		/// </summary>
		/// <param name="formName">object FormName</param>
		/// <param name="view">optional NetOffice.AccessApi.Enums.AcFormView View = 0</param>
		/// <param name="filterName">optional object FilterName</param>
		/// <param name="whereCondition">optional object WhereCondition</param>
		/// <param name="dataMode">optional NetOffice.AccessApi.Enums.AcFormOpenDataMode DataMode = -1</param>
		[CustomMethodAttribute]
		[SupportByVersionAttribute("Access", 9,10,11,12,14,15,16)]
		public void OpenForm(object formName, object view, object filterName, object whereCondition, object dataMode)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(formName, view, filterName, whereCondition, dataMode);
			Invoker.Method(this, "OpenForm", paramsArray);
		}

		/// <summary>
		/// SupportByVersion Access 9, 10, 11, 12, 14, 15, 16
		/// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff820845.aspx
		/// </summary>
		/// <param name="formName">object FormName</param>
		/// <param name="view">optional NetOffice.AccessApi.Enums.AcFormView View = 0</param>
		/// <param name="filterName">optional object FilterName</param>
		/// <param name="whereCondition">optional object WhereCondition</param>
		/// <param name="dataMode">optional NetOffice.AccessApi.Enums.AcFormOpenDataMode DataMode = -1</param>
		/// <param name="windowMode">optional NetOffice.AccessApi.Enums.AcWindowMode WindowMode = 0</param>
		[CustomMethodAttribute]
		[SupportByVersionAttribute("Access", 9,10,11,12,14,15,16)]
		public void OpenForm(object formName, object view, object filterName, object whereCondition, object dataMode, object windowMode)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(formName, view, filterName, whereCondition, dataMode, windowMode);
			Invoker.Method(this, "OpenForm", paramsArray);
		}

		/// <summary>
		/// SupportByVersion Access 9, 10, 11, 12, 14, 15, 16
		/// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff192746.aspx
		/// </summary>
		/// <param name="queryName">object QueryName</param>
		/// <param name="view">optional NetOffice.AccessApi.Enums.AcView View = 0</param>
		/// <param name="dataMode">optional NetOffice.AccessApi.Enums.AcOpenDataMode DataMode = 1</param>
		[SupportByVersionAttribute("Access", 9,10,11,12,14,15,16)]
		public void OpenQuery(object queryName, object view, object dataMode)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(queryName, view, dataMode);
			Invoker.Method(this, "OpenQuery", paramsArray);
		}

		/// <summary>
		/// SupportByVersion Access 9, 10, 11, 12, 14, 15, 16
		/// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff192746.aspx
		/// </summary>
		/// <param name="queryName">object QueryName</param>
		[CustomMethodAttribute]
		[SupportByVersionAttribute("Access", 9,10,11,12,14,15,16)]
		public void OpenQuery(object queryName)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(queryName);
			Invoker.Method(this, "OpenQuery", paramsArray);
		}

		/// <summary>
		/// SupportByVersion Access 9, 10, 11, 12, 14, 15, 16
		/// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff192746.aspx
		/// </summary>
		/// <param name="queryName">object QueryName</param>
		/// <param name="view">optional NetOffice.AccessApi.Enums.AcView View = 0</param>
		[CustomMethodAttribute]
		[SupportByVersionAttribute("Access", 9,10,11,12,14,15,16)]
		public void OpenQuery(object queryName, object view)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(queryName, view);
			Invoker.Method(this, "OpenQuery", paramsArray);
		}

		/// <summary>
		/// SupportByVersion Access 9, 10, 11, 12, 14, 15, 16
		/// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff194975.aspx
		/// </summary>
		/// <param name="tableName">object TableName</param>
		/// <param name="view">optional NetOffice.AccessApi.Enums.AcView View = 0</param>
		/// <param name="dataMode">optional NetOffice.AccessApi.Enums.AcOpenDataMode DataMode = 1</param>
		[SupportByVersionAttribute("Access", 9,10,11,12,14,15,16)]
		public void OpenTable(object tableName, object view, object dataMode)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(tableName, view, dataMode);
			Invoker.Method(this, "OpenTable", paramsArray);
		}

		/// <summary>
		/// SupportByVersion Access 9, 10, 11, 12, 14, 15, 16
		/// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff194975.aspx
		/// </summary>
		/// <param name="tableName">object TableName</param>
		[CustomMethodAttribute]
		[SupportByVersionAttribute("Access", 9,10,11,12,14,15,16)]
		public void OpenTable(object tableName)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(tableName);
			Invoker.Method(this, "OpenTable", paramsArray);
		}

		/// <summary>
		/// SupportByVersion Access 9, 10, 11, 12, 14, 15, 16
		/// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff194975.aspx
		/// </summary>
		/// <param name="tableName">object TableName</param>
		/// <param name="view">optional NetOffice.AccessApi.Enums.AcView View = 0</param>
		[CustomMethodAttribute]
		[SupportByVersionAttribute("Access", 9,10,11,12,14,15,16)]
		public void OpenTable(object tableName, object view)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(tableName, view);
			Invoker.Method(this, "OpenTable", paramsArray);
		}

		/// <summary>
		/// SupportByVersion Access 9, 10, 11, 12, 14, 15, 16
		/// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff192667.aspx
		/// </summary>
		/// <param name="printRange">optional NetOffice.AccessApi.Enums.AcPrintRange PrintRange = 0</param>
		/// <param name="pageFrom">optional object PageFrom</param>
		/// <param name="pageTo">optional object PageTo</param>
		/// <param name="printQuality">optional NetOffice.AccessApi.Enums.AcPrintQuality PrintQuality = 0</param>
		/// <param name="copies">optional object Copies</param>
		/// <param name="collateCopies">optional object CollateCopies</param>
		[SupportByVersionAttribute("Access", 9,10,11,12,14,15,16)]
		public void PrintOut(object printRange, object pageFrom, object pageTo, object printQuality, object copies, object collateCopies)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(printRange, pageFrom, pageTo, printQuality, copies, collateCopies);
			Invoker.Method(this, "PrintOut", paramsArray);
		}

		/// <summary>
		/// SupportByVersion Access 9, 10, 11, 12, 14, 15, 16
		/// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff192667.aspx
		/// </summary>
		[CustomMethodAttribute]
		[SupportByVersionAttribute("Access", 9,10,11,12,14,15,16)]
		public void PrintOut()
		{
			object[] paramsArray = null;
			Invoker.Method(this, "PrintOut", paramsArray);
		}

		/// <summary>
		/// SupportByVersion Access 9, 10, 11, 12, 14, 15, 16
		/// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff192667.aspx
		/// </summary>
		/// <param name="printRange">optional NetOffice.AccessApi.Enums.AcPrintRange PrintRange = 0</param>
		[CustomMethodAttribute]
		[SupportByVersionAttribute("Access", 9,10,11,12,14,15,16)]
		public void PrintOut(object printRange)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(printRange);
			Invoker.Method(this, "PrintOut", paramsArray);
		}

		/// <summary>
		/// SupportByVersion Access 9, 10, 11, 12, 14, 15, 16
		/// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff192667.aspx
		/// </summary>
		/// <param name="printRange">optional NetOffice.AccessApi.Enums.AcPrintRange PrintRange = 0</param>
		/// <param name="pageFrom">optional object PageFrom</param>
		[CustomMethodAttribute]
		[SupportByVersionAttribute("Access", 9,10,11,12,14,15,16)]
		public void PrintOut(object printRange, object pageFrom)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(printRange, pageFrom);
			Invoker.Method(this, "PrintOut", paramsArray);
		}

		/// <summary>
		/// SupportByVersion Access 9, 10, 11, 12, 14, 15, 16
		/// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff192667.aspx
		/// </summary>
		/// <param name="printRange">optional NetOffice.AccessApi.Enums.AcPrintRange PrintRange = 0</param>
		/// <param name="pageFrom">optional object PageFrom</param>
		/// <param name="pageTo">optional object PageTo</param>
		[CustomMethodAttribute]
		[SupportByVersionAttribute("Access", 9,10,11,12,14,15,16)]
		public void PrintOut(object printRange, object pageFrom, object pageTo)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(printRange, pageFrom, pageTo);
			Invoker.Method(this, "PrintOut", paramsArray);
		}

		/// <summary>
		/// SupportByVersion Access 9, 10, 11, 12, 14, 15, 16
		/// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff192667.aspx
		/// </summary>
		/// <param name="printRange">optional NetOffice.AccessApi.Enums.AcPrintRange PrintRange = 0</param>
		/// <param name="pageFrom">optional object PageFrom</param>
		/// <param name="pageTo">optional object PageTo</param>
		/// <param name="printQuality">optional NetOffice.AccessApi.Enums.AcPrintQuality PrintQuality = 0</param>
		[CustomMethodAttribute]
		[SupportByVersionAttribute("Access", 9,10,11,12,14,15,16)]
		public void PrintOut(object printRange, object pageFrom, object pageTo, object printQuality)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(printRange, pageFrom, pageTo, printQuality);
			Invoker.Method(this, "PrintOut", paramsArray);
		}

		/// <summary>
		/// SupportByVersion Access 9, 10, 11, 12, 14, 15, 16
		/// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff192667.aspx
		/// </summary>
		/// <param name="printRange">optional NetOffice.AccessApi.Enums.AcPrintRange PrintRange = 0</param>
		/// <param name="pageFrom">optional object PageFrom</param>
		/// <param name="pageTo">optional object PageTo</param>
		/// <param name="printQuality">optional NetOffice.AccessApi.Enums.AcPrintQuality PrintQuality = 0</param>
		/// <param name="copies">optional object Copies</param>
		[CustomMethodAttribute]
		[SupportByVersionAttribute("Access", 9,10,11,12,14,15,16)]
		public void PrintOut(object printRange, object pageFrom, object pageTo, object printQuality, object copies)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(printRange, pageFrom, pageTo, printQuality, copies);
			Invoker.Method(this, "PrintOut", paramsArray);
		}

		/// <summary>
		/// SupportByVersion Access 9, 10, 11, 12, 14, 15, 16
		/// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff191887.aspx
		/// </summary>
		/// <param name="options">optional NetOffice.AccessApi.Enums.AcQuitOption Options = 1</param>
		[SupportByVersionAttribute("Access", 9,10,11,12,14,15,16)]
		public void Quit(object options)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(options);
			Invoker.Method(this, "Quit", paramsArray);
		}

		/// <summary>
		/// SupportByVersion Access 9, 10, 11, 12, 14, 15, 16
		/// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff191887.aspx
		/// </summary>
		[CustomMethodAttribute]
		[SupportByVersionAttribute("Access", 9,10,11,12,14,15,16)]
		public void Quit()
		{
			object[] paramsArray = null;
			Invoker.Method(this, "Quit", paramsArray);
		}

		/// <summary>
		/// SupportByVersion Access 9, 10, 11, 12, 14, 15, 16
		/// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff195253.aspx
		/// </summary>
		/// <param name="controlName">optional object ControlName</param>
		[SupportByVersionAttribute("Access", 9,10,11,12,14,15,16)]
		public void Requery(object controlName)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(controlName);
			Invoker.Method(this, "Requery", paramsArray);
		}

		/// <summary>
		/// SupportByVersion Access 9, 10, 11, 12, 14, 15, 16
		/// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff195253.aspx
		/// </summary>
		[CustomMethodAttribute]
		[SupportByVersionAttribute("Access", 9,10,11,12,14,15,16)]
		public void Requery()
		{
			object[] paramsArray = null;
			Invoker.Method(this, "Requery", paramsArray);
		}

		/// <summary>
		/// SupportByVersion Access 9, 10, 11, 12, 14, 15, 16
		/// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff195560.aspx
		/// </summary>
		/// <param name="objectType">optional NetOffice.AccessApi.Enums.AcObjectType ObjectType = -1</param>
		/// <param name="objectName">optional object ObjectName</param>
		[SupportByVersionAttribute("Access", 9,10,11,12,14,15,16)]
		public void RepaintObject(object objectType, object objectName)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(objectType, objectName);
			Invoker.Method(this, "RepaintObject", paramsArray);
		}

		/// <summary>
		/// SupportByVersion Access 9, 10, 11, 12, 14, 15, 16
		/// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff195560.aspx
		/// </summary>
		[CustomMethodAttribute]
		[SupportByVersionAttribute("Access", 9,10,11,12,14,15,16)]
		public void RepaintObject()
		{
			object[] paramsArray = null;
			Invoker.Method(this, "RepaintObject", paramsArray);
		}

		/// <summary>
		/// SupportByVersion Access 9, 10, 11, 12, 14, 15, 16
		/// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff195560.aspx
		/// </summary>
		/// <param name="objectType">optional NetOffice.AccessApi.Enums.AcObjectType ObjectType = -1</param>
		[CustomMethodAttribute]
		[SupportByVersionAttribute("Access", 9,10,11,12,14,15,16)]
		public void RepaintObject(object objectType)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(objectType);
			Invoker.Method(this, "RepaintObject", paramsArray);
		}

		/// <summary>
		/// SupportByVersion Access 9, 10, 11, 12, 14, 15, 16
		/// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff823209.aspx
		/// </summary>
		/// <param name="newName">object NewName</param>
		/// <param name="objectType">optional NetOffice.AccessApi.Enums.AcObjectType ObjectType = -1</param>
		/// <param name="oldName">optional object OldName</param>
		[SupportByVersionAttribute("Access", 9,10,11,12,14,15,16)]
		public void Rename(object newName, object objectType, object oldName)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(newName, objectType, oldName);
			Invoker.Method(this, "Rename", paramsArray);
		}

		/// <summary>
		/// SupportByVersion Access 9, 10, 11, 12, 14, 15, 16
		/// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff823209.aspx
		/// </summary>
		/// <param name="newName">object NewName</param>
		[CustomMethodAttribute]
		[SupportByVersionAttribute("Access", 9,10,11,12,14,15,16)]
		public void Rename(object newName)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(newName);
			Invoker.Method(this, "Rename", paramsArray);
		}

		/// <summary>
		/// SupportByVersion Access 9, 10, 11, 12, 14, 15, 16
		/// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff823209.aspx
		/// </summary>
		/// <param name="newName">object NewName</param>
		/// <param name="objectType">optional NetOffice.AccessApi.Enums.AcObjectType ObjectType = -1</param>
		[CustomMethodAttribute]
		[SupportByVersionAttribute("Access", 9,10,11,12,14,15,16)]
		public void Rename(object newName, object objectType)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(newName, objectType);
			Invoker.Method(this, "Rename", paramsArray);
		}

		/// <summary>
		/// SupportByVersion Access 9, 10, 11, 12, 14, 15, 16
		/// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff193174.aspx
		/// </summary>
		[SupportByVersionAttribute("Access", 9,10,11,12,14,15,16)]
		public void Restore()
		{
			object[] paramsArray = null;
			Invoker.Method(this, "Restore", paramsArray);
		}

		/// <summary>
		/// SupportByVersion Access 9, 10, 11, 12, 14, 15, 16
		/// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff192075.aspx
		/// </summary>
		/// <param name="macroName">object MacroName</param>
		/// <param name="repeatCount">optional object RepeatCount</param>
		/// <param name="repeatExpression">optional object RepeatExpression</param>
		[SupportByVersionAttribute("Access", 9,10,11,12,14,15,16)]
		public void RunMacro(object macroName, object repeatCount, object repeatExpression)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(macroName, repeatCount, repeatExpression);
			Invoker.Method(this, "RunMacro", paramsArray);
		}

		/// <summary>
		/// SupportByVersion Access 9, 10, 11, 12, 14, 15, 16
		/// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff192075.aspx
		/// </summary>
		/// <param name="macroName">object MacroName</param>
		[CustomMethodAttribute]
		[SupportByVersionAttribute("Access", 9,10,11,12,14,15,16)]
		public void RunMacro(object macroName)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(macroName);
			Invoker.Method(this, "RunMacro", paramsArray);
		}

		/// <summary>
		/// SupportByVersion Access 9, 10, 11, 12, 14, 15, 16
		/// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff192075.aspx
		/// </summary>
		/// <param name="macroName">object MacroName</param>
		/// <param name="repeatCount">optional object RepeatCount</param>
		[CustomMethodAttribute]
		[SupportByVersionAttribute("Access", 9,10,11,12,14,15,16)]
		public void RunMacro(object macroName, object repeatCount)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(macroName, repeatCount);
			Invoker.Method(this, "RunMacro", paramsArray);
		}

		/// <summary>
		/// SupportByVersion Access 9, 10, 11, 12, 14, 15, 16
		/// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff194626.aspx
		/// </summary>
		/// <param name="sQLStatement">object SQLStatement</param>
		/// <param name="useTransaction">optional object UseTransaction</param>
		[SupportByVersionAttribute("Access", 9,10,11,12,14,15,16)]
		public void RunSQL(object sQLStatement, object useTransaction)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(sQLStatement, useTransaction);
			Invoker.Method(this, "RunSQL", paramsArray);
		}

		/// <summary>
		/// SupportByVersion Access 9, 10, 11, 12, 14, 15, 16
		/// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff194626.aspx
		/// </summary>
		/// <param name="sQLStatement">object SQLStatement</param>
		[CustomMethodAttribute]
		[SupportByVersionAttribute("Access", 9,10,11,12,14,15,16)]
		public void RunSQL(object sQLStatement)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(sQLStatement);
			Invoker.Method(this, "RunSQL", paramsArray);
		}

		/// <summary>
		/// SupportByVersion Access 9, 10, 11, 12, 14, 15, 16
		/// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff835629.aspx
		/// </summary>
		/// <param name="objectType">NetOffice.AccessApi.Enums.AcObjectType ObjectType</param>
		/// <param name="objectName">optional object ObjectName</param>
		/// <param name="inDatabaseWindow">optional object InDatabaseWindow</param>
		[SupportByVersionAttribute("Access", 9,10,11,12,14,15,16)]
		public void SelectObject(NetOffice.AccessApi.Enums.AcObjectType objectType, object objectName, object inDatabaseWindow)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(objectType, objectName, inDatabaseWindow);
			Invoker.Method(this, "SelectObject", paramsArray);
		}

		/// <summary>
		/// SupportByVersion Access 9, 10, 11, 12, 14, 15, 16
		/// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff835629.aspx
		/// </summary>
		/// <param name="objectType">NetOffice.AccessApi.Enums.AcObjectType ObjectType</param>
		[CustomMethodAttribute]
		[SupportByVersionAttribute("Access", 9,10,11,12,14,15,16)]
		public void SelectObject(NetOffice.AccessApi.Enums.AcObjectType objectType)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(objectType);
			Invoker.Method(this, "SelectObject", paramsArray);
		}

		/// <summary>
		/// SupportByVersion Access 9, 10, 11, 12, 14, 15, 16
		/// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff835629.aspx
		/// </summary>
		/// <param name="objectType">NetOffice.AccessApi.Enums.AcObjectType ObjectType</param>
		/// <param name="objectName">optional object ObjectName</param>
		[CustomMethodAttribute]
		[SupportByVersionAttribute("Access", 9,10,11,12,14,15,16)]
		public void SelectObject(NetOffice.AccessApi.Enums.AcObjectType objectType, object objectName)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(objectType, objectName);
			Invoker.Method(this, "SelectObject", paramsArray);
		}

		/// <summary>
		/// SupportByVersion Access 9, 10, 11, 12, 14, 15, 16
		/// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff837275.aspx
		/// </summary>
		/// <param name="warningsOn">object WarningsOn</param>
		[SupportByVersionAttribute("Access", 9,10,11,12,14,15,16)]
		public void SetWarnings(object warningsOn)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(warningsOn);
			Invoker.Method(this, "SetWarnings", paramsArray);
		}

		/// <summary>
		/// SupportByVersion Access 9, 10, 11, 12, 14, 15, 16
		/// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff195994.aspx
		/// </summary>
		[SupportByVersionAttribute("Access", 9,10,11,12,14,15,16)]
		public void ShowAllRecords()
		{
			object[] paramsArray = null;
			Invoker.Method(this, "ShowAllRecords", paramsArray);
		}

		/// <summary>
		/// SupportByVersion Access 9, 10, 11, 12, 14, 15, 16
		/// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff192676.aspx
		/// </summary>
		/// <param name="reportName">object ReportName</param>
		/// <param name="view">optional NetOffice.AccessApi.Enums.AcView View = 0</param>
		/// <param name="filterName">optional object FilterName</param>
		/// <param name="whereCondition">optional object WhereCondition</param>
		[SupportByVersionAttribute("Access", 9,10,11,12,14,15,16)]
		public void OpenReport(object reportName, object view, object filterName, object whereCondition)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(reportName, view, filterName, whereCondition);
			Invoker.Method(this, "OpenReport", paramsArray);
		}

		/// <summary>
		/// SupportByVersion Access 10, 11, 12, 14, 15, 16
		/// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff192676.aspx
		/// </summary>
		/// <param name="reportName">object ReportName</param>
		/// <param name="view">optional NetOffice.AccessApi.Enums.AcView View = 0</param>
		/// <param name="filterName">optional object FilterName</param>
		/// <param name="whereCondition">optional object WhereCondition</param>
		/// <param name="windowMode">optional NetOffice.AccessApi.Enums.AcWindowMode WindowMode = 0</param>
		/// <param name="openArgs">optional object OpenArgs</param>
		[SupportByVersionAttribute("Access", 10,11,12,14,15,16)]
		public void OpenReport(object reportName, object view, object filterName, object whereCondition, object windowMode, object openArgs)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(reportName, view, filterName, whereCondition, windowMode, openArgs);
			Invoker.Method(this, "OpenReport", paramsArray);
		}

		/// <summary>
		/// SupportByVersion Access 9, 10, 11, 12, 14, 15, 16
		/// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff192676.aspx
		/// </summary>
		/// <param name="reportName">object ReportName</param>
		[CustomMethodAttribute]
		[SupportByVersionAttribute("Access", 9,10,11,12,14,15,16)]
		public void OpenReport(object reportName)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(reportName);
			Invoker.Method(this, "OpenReport", paramsArray);
		}

		/// <summary>
		/// SupportByVersion Access 9, 10, 11, 12, 14, 15, 16
		/// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff192676.aspx
		/// </summary>
		/// <param name="reportName">object ReportName</param>
		/// <param name="view">optional NetOffice.AccessApi.Enums.AcView View = 0</param>
		[CustomMethodAttribute]
		[SupportByVersionAttribute("Access", 9,10,11,12,14,15,16)]
		public void OpenReport(object reportName, object view)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(reportName, view);
			Invoker.Method(this, "OpenReport", paramsArray);
		}

		/// <summary>
		/// SupportByVersion Access 9, 10, 11, 12, 14, 15, 16
		/// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff192676.aspx
		/// </summary>
		/// <param name="reportName">object ReportName</param>
		/// <param name="view">optional NetOffice.AccessApi.Enums.AcView View = 0</param>
		/// <param name="filterName">optional object FilterName</param>
		[CustomMethodAttribute]
		[SupportByVersionAttribute("Access", 9,10,11,12,14,15,16)]
		public void OpenReport(object reportName, object view, object filterName)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(reportName, view, filterName);
			Invoker.Method(this, "OpenReport", paramsArray);
		}

		/// <summary>
		/// SupportByVersion Access 10, 11, 12, 14, 15, 16
		/// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff192676.aspx
		/// </summary>
		/// <param name="reportName">object ReportName</param>
		/// <param name="view">optional NetOffice.AccessApi.Enums.AcView View = 0</param>
		/// <param name="filterName">optional object FilterName</param>
		/// <param name="whereCondition">optional object WhereCondition</param>
		/// <param name="windowMode">optional NetOffice.AccessApi.Enums.AcWindowMode WindowMode = 0</param>
		[CustomMethodAttribute]
		[SupportByVersionAttribute("Access", 10,11,12,14,15,16)]
		public void OpenReport(object reportName, object view, object filterName, object whereCondition, object windowMode)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(reportName, view, filterName, whereCondition, windowMode);
			Invoker.Method(this, "OpenReport", paramsArray);
		}

		/// <summary>
		/// SupportByVersion Access 9, 10, 11, 12, 14, 15, 16
		/// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff196455.aspx
		/// </summary>
		/// <param name="transferType">optional NetOffice.AccessApi.Enums.AcDataTransferType TransferType = 0</param>
		/// <param name="databaseType">optional object DatabaseType</param>
		/// <param name="databaseName">optional object DatabaseName</param>
		/// <param name="objectType">optional NetOffice.AccessApi.Enums.AcObjectType ObjectType = 0</param>
		/// <param name="source">optional object Source</param>
		/// <param name="destination">optional object Destination</param>
		/// <param name="structureOnly">optional object StructureOnly</param>
		/// <param name="storeLogin">optional object StoreLogin</param>
		[SupportByVersionAttribute("Access", 9,10,11,12,14,15,16)]
		public void TransferDatabase(object transferType, object databaseType, object databaseName, object objectType, object source, object destination, object structureOnly, object storeLogin)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(transferType, databaseType, databaseName, objectType, source, destination, structureOnly, storeLogin);
			Invoker.Method(this, "TransferDatabase", paramsArray);
		}

		/// <summary>
		/// SupportByVersion Access 9, 10, 11, 12, 14, 15, 16
		/// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff196455.aspx
		/// </summary>
		[CustomMethodAttribute]
		[SupportByVersionAttribute("Access", 9,10,11,12,14,15,16)]
		public void TransferDatabase()
		{
			object[] paramsArray = null;
			Invoker.Method(this, "TransferDatabase", paramsArray);
		}

		/// <summary>
		/// SupportByVersion Access 9, 10, 11, 12, 14, 15, 16
		/// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff196455.aspx
		/// </summary>
		/// <param name="transferType">optional NetOffice.AccessApi.Enums.AcDataTransferType TransferType = 0</param>
		[CustomMethodAttribute]
		[SupportByVersionAttribute("Access", 9,10,11,12,14,15,16)]
		public void TransferDatabase(object transferType)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(transferType);
			Invoker.Method(this, "TransferDatabase", paramsArray);
		}

		/// <summary>
		/// SupportByVersion Access 9, 10, 11, 12, 14, 15, 16
		/// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff196455.aspx
		/// </summary>
		/// <param name="transferType">optional NetOffice.AccessApi.Enums.AcDataTransferType TransferType = 0</param>
		/// <param name="databaseType">optional object DatabaseType</param>
		[CustomMethodAttribute]
		[SupportByVersionAttribute("Access", 9,10,11,12,14,15,16)]
		public void TransferDatabase(object transferType, object databaseType)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(transferType, databaseType);
			Invoker.Method(this, "TransferDatabase", paramsArray);
		}

		/// <summary>
		/// SupportByVersion Access 9, 10, 11, 12, 14, 15, 16
		/// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff196455.aspx
		/// </summary>
		/// <param name="transferType">optional NetOffice.AccessApi.Enums.AcDataTransferType TransferType = 0</param>
		/// <param name="databaseType">optional object DatabaseType</param>
		/// <param name="databaseName">optional object DatabaseName</param>
		[CustomMethodAttribute]
		[SupportByVersionAttribute("Access", 9,10,11,12,14,15,16)]
		public void TransferDatabase(object transferType, object databaseType, object databaseName)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(transferType, databaseType, databaseName);
			Invoker.Method(this, "TransferDatabase", paramsArray);
		}

		/// <summary>
		/// SupportByVersion Access 9, 10, 11, 12, 14, 15, 16
		/// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff196455.aspx
		/// </summary>
		/// <param name="transferType">optional NetOffice.AccessApi.Enums.AcDataTransferType TransferType = 0</param>
		/// <param name="databaseType">optional object DatabaseType</param>
		/// <param name="databaseName">optional object DatabaseName</param>
		/// <param name="objectType">optional NetOffice.AccessApi.Enums.AcObjectType ObjectType = 0</param>
		[CustomMethodAttribute]
		[SupportByVersionAttribute("Access", 9,10,11,12,14,15,16)]
		public void TransferDatabase(object transferType, object databaseType, object databaseName, object objectType)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(transferType, databaseType, databaseName, objectType);
			Invoker.Method(this, "TransferDatabase", paramsArray);
		}

		/// <summary>
		/// SupportByVersion Access 9, 10, 11, 12, 14, 15, 16
		/// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff196455.aspx
		/// </summary>
		/// <param name="transferType">optional NetOffice.AccessApi.Enums.AcDataTransferType TransferType = 0</param>
		/// <param name="databaseType">optional object DatabaseType</param>
		/// <param name="databaseName">optional object DatabaseName</param>
		/// <param name="objectType">optional NetOffice.AccessApi.Enums.AcObjectType ObjectType = 0</param>
		/// <param name="source">optional object Source</param>
		[CustomMethodAttribute]
		[SupportByVersionAttribute("Access", 9,10,11,12,14,15,16)]
		public void TransferDatabase(object transferType, object databaseType, object databaseName, object objectType, object source)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(transferType, databaseType, databaseName, objectType, source);
			Invoker.Method(this, "TransferDatabase", paramsArray);
		}

		/// <summary>
		/// SupportByVersion Access 9, 10, 11, 12, 14, 15, 16
		/// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff196455.aspx
		/// </summary>
		/// <param name="transferType">optional NetOffice.AccessApi.Enums.AcDataTransferType TransferType = 0</param>
		/// <param name="databaseType">optional object DatabaseType</param>
		/// <param name="databaseName">optional object DatabaseName</param>
		/// <param name="objectType">optional NetOffice.AccessApi.Enums.AcObjectType ObjectType = 0</param>
		/// <param name="source">optional object Source</param>
		/// <param name="destination">optional object Destination</param>
		[CustomMethodAttribute]
		[SupportByVersionAttribute("Access", 9,10,11,12,14,15,16)]
		public void TransferDatabase(object transferType, object databaseType, object databaseName, object objectType, object source, object destination)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(transferType, databaseType, databaseName, objectType, source, destination);
			Invoker.Method(this, "TransferDatabase", paramsArray);
		}

		/// <summary>
		/// SupportByVersion Access 9, 10, 11, 12, 14, 15, 16
		/// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff196455.aspx
		/// </summary>
		/// <param name="transferType">optional NetOffice.AccessApi.Enums.AcDataTransferType TransferType = 0</param>
		/// <param name="databaseType">optional object DatabaseType</param>
		/// <param name="databaseName">optional object DatabaseName</param>
		/// <param name="objectType">optional NetOffice.AccessApi.Enums.AcObjectType ObjectType = 0</param>
		/// <param name="source">optional object Source</param>
		/// <param name="destination">optional object Destination</param>
		/// <param name="structureOnly">optional object StructureOnly</param>
		[CustomMethodAttribute]
		[SupportByVersionAttribute("Access", 9,10,11,12,14,15,16)]
		public void TransferDatabase(object transferType, object databaseType, object databaseName, object objectType, object source, object destination, object structureOnly)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(transferType, databaseType, databaseName, objectType, source, destination, structureOnly);
			Invoker.Method(this, "TransferDatabase", paramsArray);
		}

		/// <summary>
		/// SupportByVersion Access 9, 10, 11, 12, 14, 15, 16
		/// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff844793.aspx
		/// </summary>
		/// <param name="transferType">optional NetOffice.AccessApi.Enums.AcDataTransferType TransferType = 0</param>
		/// <param name="spreadsheetType">optional NetOffice.AccessApi.Enums.AcSpreadSheetType SpreadsheetType = 8</param>
		/// <param name="tableName">optional object TableName</param>
		/// <param name="fileName">optional object FileName</param>
		/// <param name="hasFieldNames">optional object HasFieldNames</param>
		/// <param name="range">optional object Range</param>
		/// <param name="useOA">optional object UseOA</param>
		[SupportByVersionAttribute("Access", 9,10,11,12,14,15,16)]
		public void TransferSpreadsheet(object transferType, object spreadsheetType, object tableName, object fileName, object hasFieldNames, object range, object useOA)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(transferType, spreadsheetType, tableName, fileName, hasFieldNames, range, useOA);
			Invoker.Method(this, "TransferSpreadsheet", paramsArray);
		}

		/// <summary>
		/// SupportByVersion Access 9, 10, 11, 12, 14, 15, 16
		/// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff844793.aspx
		/// </summary>
		[CustomMethodAttribute]
		[SupportByVersionAttribute("Access", 9,10,11,12,14,15,16)]
		public void TransferSpreadsheet()
		{
			object[] paramsArray = null;
			Invoker.Method(this, "TransferSpreadsheet", paramsArray);
		}

		/// <summary>
		/// SupportByVersion Access 9, 10, 11, 12, 14, 15, 16
		/// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff844793.aspx
		/// </summary>
		/// <param name="transferType">optional NetOffice.AccessApi.Enums.AcDataTransferType TransferType = 0</param>
		[CustomMethodAttribute]
		[SupportByVersionAttribute("Access", 9,10,11,12,14,15,16)]
		public void TransferSpreadsheet(object transferType)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(transferType);
			Invoker.Method(this, "TransferSpreadsheet", paramsArray);
		}

		/// <summary>
		/// SupportByVersion Access 9, 10, 11, 12, 14, 15, 16
		/// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff844793.aspx
		/// </summary>
		/// <param name="transferType">optional NetOffice.AccessApi.Enums.AcDataTransferType TransferType = 0</param>
		/// <param name="spreadsheetType">optional NetOffice.AccessApi.Enums.AcSpreadSheetType SpreadsheetType = 8</param>
		[CustomMethodAttribute]
		[SupportByVersionAttribute("Access", 9,10,11,12,14,15,16)]
		public void TransferSpreadsheet(object transferType, object spreadsheetType)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(transferType, spreadsheetType);
			Invoker.Method(this, "TransferSpreadsheet", paramsArray);
		}

		/// <summary>
		/// SupportByVersion Access 9, 10, 11, 12, 14, 15, 16
		/// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff844793.aspx
		/// </summary>
		/// <param name="transferType">optional NetOffice.AccessApi.Enums.AcDataTransferType TransferType = 0</param>
		/// <param name="spreadsheetType">optional NetOffice.AccessApi.Enums.AcSpreadSheetType SpreadsheetType = 8</param>
		/// <param name="tableName">optional object TableName</param>
		[CustomMethodAttribute]
		[SupportByVersionAttribute("Access", 9,10,11,12,14,15,16)]
		public void TransferSpreadsheet(object transferType, object spreadsheetType, object tableName)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(transferType, spreadsheetType, tableName);
			Invoker.Method(this, "TransferSpreadsheet", paramsArray);
		}

		/// <summary>
		/// SupportByVersion Access 9, 10, 11, 12, 14, 15, 16
		/// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff844793.aspx
		/// </summary>
		/// <param name="transferType">optional NetOffice.AccessApi.Enums.AcDataTransferType TransferType = 0</param>
		/// <param name="spreadsheetType">optional NetOffice.AccessApi.Enums.AcSpreadSheetType SpreadsheetType = 8</param>
		/// <param name="tableName">optional object TableName</param>
		/// <param name="fileName">optional object FileName</param>
		[CustomMethodAttribute]
		[SupportByVersionAttribute("Access", 9,10,11,12,14,15,16)]
		public void TransferSpreadsheet(object transferType, object spreadsheetType, object tableName, object fileName)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(transferType, spreadsheetType, tableName, fileName);
			Invoker.Method(this, "TransferSpreadsheet", paramsArray);
		}

		/// <summary>
		/// SupportByVersion Access 9, 10, 11, 12, 14, 15, 16
		/// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff844793.aspx
		/// </summary>
		/// <param name="transferType">optional NetOffice.AccessApi.Enums.AcDataTransferType TransferType = 0</param>
		/// <param name="spreadsheetType">optional NetOffice.AccessApi.Enums.AcSpreadSheetType SpreadsheetType = 8</param>
		/// <param name="tableName">optional object TableName</param>
		/// <param name="fileName">optional object FileName</param>
		/// <param name="hasFieldNames">optional object HasFieldNames</param>
		[CustomMethodAttribute]
		[SupportByVersionAttribute("Access", 9,10,11,12,14,15,16)]
		public void TransferSpreadsheet(object transferType, object spreadsheetType, object tableName, object fileName, object hasFieldNames)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(transferType, spreadsheetType, tableName, fileName, hasFieldNames);
			Invoker.Method(this, "TransferSpreadsheet", paramsArray);
		}

		/// <summary>
		/// SupportByVersion Access 9, 10, 11, 12, 14, 15, 16
		/// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff844793.aspx
		/// </summary>
		/// <param name="transferType">optional NetOffice.AccessApi.Enums.AcDataTransferType TransferType = 0</param>
		/// <param name="spreadsheetType">optional NetOffice.AccessApi.Enums.AcSpreadSheetType SpreadsheetType = 8</param>
		/// <param name="tableName">optional object TableName</param>
		/// <param name="fileName">optional object FileName</param>
		/// <param name="hasFieldNames">optional object HasFieldNames</param>
		/// <param name="range">optional object Range</param>
		[CustomMethodAttribute]
		[SupportByVersionAttribute("Access", 9,10,11,12,14,15,16)]
		public void TransferSpreadsheet(object transferType, object spreadsheetType, object tableName, object fileName, object hasFieldNames, object range)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(transferType, spreadsheetType, tableName, fileName, hasFieldNames, range);
			Invoker.Method(this, "TransferSpreadsheet", paramsArray);
		}

		/// <summary>
		/// SupportByVersion Access 9, 10, 11, 12, 14, 15, 16
		/// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff835958.aspx
		/// </summary>
		/// <param name="transferType">optional NetOffice.AccessApi.Enums.AcTextTransferType TransferType = 0</param>
		/// <param name="specificationName">optional object SpecificationName</param>
		/// <param name="tableName">optional object TableName</param>
		/// <param name="fileName">optional object FileName</param>
		/// <param name="hasFieldNames">optional object HasFieldNames</param>
		/// <param name="hTMLTableName">optional object HTMLTableName</param>
		/// <param name="codePage">optional object CodePage</param>
		[SupportByVersionAttribute("Access", 9,10,11,12,14,15,16)]
		public void TransferText(object transferType, object specificationName, object tableName, object fileName, object hasFieldNames, object hTMLTableName, object codePage)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(transferType, specificationName, tableName, fileName, hasFieldNames, hTMLTableName, codePage);
			Invoker.Method(this, "TransferText", paramsArray);
		}

		/// <summary>
		/// SupportByVersion Access 9, 10, 11, 12, 14, 15, 16
		/// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff835958.aspx
		/// </summary>
		[CustomMethodAttribute]
		[SupportByVersionAttribute("Access", 9,10,11,12,14,15,16)]
		public void TransferText()
		{
			object[] paramsArray = null;
			Invoker.Method(this, "TransferText", paramsArray);
		}

		/// <summary>
		/// SupportByVersion Access 9, 10, 11, 12, 14, 15, 16
		/// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff835958.aspx
		/// </summary>
		/// <param name="transferType">optional NetOffice.AccessApi.Enums.AcTextTransferType TransferType = 0</param>
		[CustomMethodAttribute]
		[SupportByVersionAttribute("Access", 9,10,11,12,14,15,16)]
		public void TransferText(object transferType)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(transferType);
			Invoker.Method(this, "TransferText", paramsArray);
		}

		/// <summary>
		/// SupportByVersion Access 9, 10, 11, 12, 14, 15, 16
		/// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff835958.aspx
		/// </summary>
		/// <param name="transferType">optional NetOffice.AccessApi.Enums.AcTextTransferType TransferType = 0</param>
		/// <param name="specificationName">optional object SpecificationName</param>
		[CustomMethodAttribute]
		[SupportByVersionAttribute("Access", 9,10,11,12,14,15,16)]
		public void TransferText(object transferType, object specificationName)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(transferType, specificationName);
			Invoker.Method(this, "TransferText", paramsArray);
		}

		/// <summary>
		/// SupportByVersion Access 9, 10, 11, 12, 14, 15, 16
		/// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff835958.aspx
		/// </summary>
		/// <param name="transferType">optional NetOffice.AccessApi.Enums.AcTextTransferType TransferType = 0</param>
		/// <param name="specificationName">optional object SpecificationName</param>
		/// <param name="tableName">optional object TableName</param>
		[CustomMethodAttribute]
		[SupportByVersionAttribute("Access", 9,10,11,12,14,15,16)]
		public void TransferText(object transferType, object specificationName, object tableName)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(transferType, specificationName, tableName);
			Invoker.Method(this, "TransferText", paramsArray);
		}

		/// <summary>
		/// SupportByVersion Access 9, 10, 11, 12, 14, 15, 16
		/// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff835958.aspx
		/// </summary>
		/// <param name="transferType">optional NetOffice.AccessApi.Enums.AcTextTransferType TransferType = 0</param>
		/// <param name="specificationName">optional object SpecificationName</param>
		/// <param name="tableName">optional object TableName</param>
		/// <param name="fileName">optional object FileName</param>
		[CustomMethodAttribute]
		[SupportByVersionAttribute("Access", 9,10,11,12,14,15,16)]
		public void TransferText(object transferType, object specificationName, object tableName, object fileName)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(transferType, specificationName, tableName, fileName);
			Invoker.Method(this, "TransferText", paramsArray);
		}

		/// <summary>
		/// SupportByVersion Access 9, 10, 11, 12, 14, 15, 16
		/// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff835958.aspx
		/// </summary>
		/// <param name="transferType">optional NetOffice.AccessApi.Enums.AcTextTransferType TransferType = 0</param>
		/// <param name="specificationName">optional object SpecificationName</param>
		/// <param name="tableName">optional object TableName</param>
		/// <param name="fileName">optional object FileName</param>
		/// <param name="hasFieldNames">optional object HasFieldNames</param>
		[CustomMethodAttribute]
		[SupportByVersionAttribute("Access", 9,10,11,12,14,15,16)]
		public void TransferText(object transferType, object specificationName, object tableName, object fileName, object hasFieldNames)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(transferType, specificationName, tableName, fileName, hasFieldNames);
			Invoker.Method(this, "TransferText", paramsArray);
		}

		/// <summary>
		/// SupportByVersion Access 9, 10, 11, 12, 14, 15, 16
		/// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff835958.aspx
		/// </summary>
		/// <param name="transferType">optional NetOffice.AccessApi.Enums.AcTextTransferType TransferType = 0</param>
		/// <param name="specificationName">optional object SpecificationName</param>
		/// <param name="tableName">optional object TableName</param>
		/// <param name="fileName">optional object FileName</param>
		/// <param name="hasFieldNames">optional object HasFieldNames</param>
		/// <param name="hTMLTableName">optional object HTMLTableName</param>
		[CustomMethodAttribute]
		[SupportByVersionAttribute("Access", 9,10,11,12,14,15,16)]
		public void TransferText(object transferType, object specificationName, object tableName, object fileName, object hasFieldNames, object hTMLTableName)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(transferType, specificationName, tableName, fileName, hasFieldNames, hTMLTableName);
			Invoker.Method(this, "TransferText", paramsArray);
		}

		/// <summary>
		/// SupportByVersion Access 9, 10, 11, 12, 14, 15, 16
		/// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff192065.aspx
		/// </summary>
		/// <param name="objectType">NetOffice.AccessApi.Enums.AcOutputObjectType ObjectType</param>
		/// <param name="objectName">optional object ObjectName</param>
		/// <param name="outputFormat">optional object OutputFormat</param>
		/// <param name="outputFile">optional object OutputFile</param>
		/// <param name="autoStart">optional object AutoStart</param>
		/// <param name="templateFile">optional object TemplateFile</param>
		[SupportByVersionAttribute("Access", 9,10,11,12,14,15,16)]
		public void OutputTo(NetOffice.AccessApi.Enums.AcOutputObjectType objectType, object objectName, object outputFormat, object outputFile, object autoStart, object templateFile)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(objectType, objectName, outputFormat, outputFile, autoStart, templateFile);
			Invoker.Method(this, "OutputTo", paramsArray);
		}

		/// <summary>
		/// SupportByVersion Access 10, 11, 12, 14, 15, 16
		/// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff192065.aspx
		/// </summary>
		/// <param name="objectType">NetOffice.AccessApi.Enums.AcOutputObjectType ObjectType</param>
		/// <param name="objectName">optional object ObjectName</param>
		/// <param name="outputFormat">optional object OutputFormat</param>
		/// <param name="outputFile">optional object OutputFile</param>
		/// <param name="autoStart">optional object AutoStart</param>
		/// <param name="templateFile">optional object TemplateFile</param>
		/// <param name="encoding">optional object Encoding</param>
		[SupportByVersionAttribute("Access", 10,11,12,14,15,16)]
		public void OutputTo(NetOffice.AccessApi.Enums.AcOutputObjectType objectType, object objectName, object outputFormat, object outputFile, object autoStart, object templateFile, object encoding)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(objectType, objectName, outputFormat, outputFile, autoStart, templateFile, encoding);
			Invoker.Method(this, "OutputTo", paramsArray);
		}

		/// <summary>
		/// SupportByVersion Access 12, 14, 15, 16
		/// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff192065.aspx
		/// </summary>
		/// <param name="objectType">NetOffice.AccessApi.Enums.AcOutputObjectType ObjectType</param>
		/// <param name="objectName">optional object ObjectName</param>
		/// <param name="outputFormat">optional object OutputFormat</param>
		/// <param name="outputFile">optional object OutputFile</param>
		/// <param name="autoStart">optional object AutoStart</param>
		/// <param name="templateFile">optional object TemplateFile</param>
		/// <param name="encoding">optional object Encoding</param>
		/// <param name="outputQuality">optional NetOffice.AccessApi.Enums.AcExportQuality OutputQuality = 0</param>
		[SupportByVersionAttribute("Access", 12,14,15,16)]
		public void OutputTo(NetOffice.AccessApi.Enums.AcOutputObjectType objectType, object objectName, object outputFormat, object outputFile, object autoStart, object templateFile, object encoding, object outputQuality)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(objectType, objectName, outputFormat, outputFile, autoStart, templateFile, encoding, outputQuality);
			Invoker.Method(this, "OutputTo", paramsArray);
		}

		/// <summary>
		/// SupportByVersion Access 9, 10, 11, 12, 14, 15, 16
		/// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff192065.aspx
		/// </summary>
		/// <param name="objectType">NetOffice.AccessApi.Enums.AcOutputObjectType ObjectType</param>
		[CustomMethodAttribute]
		[SupportByVersionAttribute("Access", 9,10,11,12,14,15,16)]
		public void OutputTo(NetOffice.AccessApi.Enums.AcOutputObjectType objectType)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(objectType);
			Invoker.Method(this, "OutputTo", paramsArray);
		}

		/// <summary>
		/// SupportByVersion Access 9, 10, 11, 12, 14, 15, 16
		/// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff192065.aspx
		/// </summary>
		/// <param name="objectType">NetOffice.AccessApi.Enums.AcOutputObjectType ObjectType</param>
		/// <param name="objectName">optional object ObjectName</param>
		[CustomMethodAttribute]
		[SupportByVersionAttribute("Access", 9,10,11,12,14,15,16)]
		public void OutputTo(NetOffice.AccessApi.Enums.AcOutputObjectType objectType, object objectName)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(objectType, objectName);
			Invoker.Method(this, "OutputTo", paramsArray);
		}

		/// <summary>
		/// SupportByVersion Access 9, 10, 11, 12, 14, 15, 16
		/// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff192065.aspx
		/// </summary>
		/// <param name="objectType">NetOffice.AccessApi.Enums.AcOutputObjectType ObjectType</param>
		/// <param name="objectName">optional object ObjectName</param>
		/// <param name="outputFormat">optional object OutputFormat</param>
		[CustomMethodAttribute]
		[SupportByVersionAttribute("Access", 9,10,11,12,14,15,16)]
		public void OutputTo(NetOffice.AccessApi.Enums.AcOutputObjectType objectType, object objectName, object outputFormat)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(objectType, objectName, outputFormat);
			Invoker.Method(this, "OutputTo", paramsArray);
		}

		/// <summary>
		/// SupportByVersion Access 9, 10, 11, 12, 14, 15, 16
		/// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff192065.aspx
		/// </summary>
		/// <param name="objectType">NetOffice.AccessApi.Enums.AcOutputObjectType ObjectType</param>
		/// <param name="objectName">optional object ObjectName</param>
		/// <param name="outputFormat">optional object OutputFormat</param>
		/// <param name="outputFile">optional object OutputFile</param>
		[CustomMethodAttribute]
		[SupportByVersionAttribute("Access", 9,10,11,12,14,15,16)]
		public void OutputTo(NetOffice.AccessApi.Enums.AcOutputObjectType objectType, object objectName, object outputFormat, object outputFile)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(objectType, objectName, outputFormat, outputFile);
			Invoker.Method(this, "OutputTo", paramsArray);
		}

		/// <summary>
		/// SupportByVersion Access 9, 10, 11, 12, 14, 15, 16
		/// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff192065.aspx
		/// </summary>
		/// <param name="objectType">NetOffice.AccessApi.Enums.AcOutputObjectType ObjectType</param>
		/// <param name="objectName">optional object ObjectName</param>
		/// <param name="outputFormat">optional object OutputFormat</param>
		/// <param name="outputFile">optional object OutputFile</param>
		/// <param name="autoStart">optional object AutoStart</param>
		[CustomMethodAttribute]
		[SupportByVersionAttribute("Access", 9,10,11,12,14,15,16)]
		public void OutputTo(NetOffice.AccessApi.Enums.AcOutputObjectType objectType, object objectName, object outputFormat, object outputFile, object autoStart)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(objectType, objectName, outputFormat, outputFile, autoStart);
			Invoker.Method(this, "OutputTo", paramsArray);
		}

		/// <summary>
		/// SupportByVersion Access 9, 10, 11, 12, 14, 15, 16
		/// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff197376.aspx
		/// </summary>
		/// <param name="objectType">optional NetOffice.AccessApi.Enums.AcObjectType ObjectType = -1</param>
		/// <param name="objectName">optional object ObjectName</param>
		[SupportByVersionAttribute("Access", 9,10,11,12,14,15,16)]
		public void DeleteObject(object objectType, object objectName)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(objectType, objectName);
			Invoker.Method(this, "DeleteObject", paramsArray);
		}

		/// <summary>
		/// SupportByVersion Access 9, 10, 11, 12, 14, 15, 16
		/// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff197376.aspx
		/// </summary>
		[CustomMethodAttribute]
		[SupportByVersionAttribute("Access", 9,10,11,12,14,15,16)]
		public void DeleteObject()
		{
			object[] paramsArray = null;
			Invoker.Method(this, "DeleteObject", paramsArray);
		}

		/// <summary>
		/// SupportByVersion Access 9, 10, 11, 12, 14, 15, 16
		/// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff197376.aspx
		/// </summary>
		/// <param name="objectType">optional NetOffice.AccessApi.Enums.AcObjectType ObjectType = -1</param>
		[CustomMethodAttribute]
		[SupportByVersionAttribute("Access", 9,10,11,12,14,15,16)]
		public void DeleteObject(object objectType)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(objectType);
			Invoker.Method(this, "DeleteObject", paramsArray);
		}

		/// <summary>
		/// SupportByVersion Access 9, 10, 11, 12, 14, 15, 16
		/// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff192698.aspx
		/// </summary>
		/// <param name="moduleName">optional object ModuleName</param>
		/// <param name="procedureName">optional object ProcedureName</param>
		[SupportByVersionAttribute("Access", 9,10,11,12,14,15,16)]
		public void OpenModule(object moduleName, object procedureName)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(moduleName, procedureName);
			Invoker.Method(this, "OpenModule", paramsArray);
		}

		/// <summary>
		/// SupportByVersion Access 9, 10, 11, 12, 14, 15, 16
		/// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff192698.aspx
		/// </summary>
		[CustomMethodAttribute]
		[SupportByVersionAttribute("Access", 9,10,11,12,14,15,16)]
		public void OpenModule()
		{
			object[] paramsArray = null;
			Invoker.Method(this, "OpenModule", paramsArray);
		}

		/// <summary>
		/// SupportByVersion Access 9, 10, 11, 12, 14, 15, 16
		/// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff192698.aspx
		/// </summary>
		/// <param name="moduleName">optional object ModuleName</param>
		[CustomMethodAttribute]
		[SupportByVersionAttribute("Access", 9,10,11,12,14,15,16)]
		public void OpenModule(object moduleName)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(moduleName);
			Invoker.Method(this, "OpenModule", paramsArray);
		}

		/// <summary>
		/// SupportByVersion Access 9, 10, 11, 12, 14, 15, 16
		/// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff197046.aspx
		/// </summary>
		/// <param name="objectType">optional NetOffice.AccessApi.Enums.AcSendObjectType ObjectType = -1</param>
		/// <param name="objectName">optional object ObjectName</param>
		/// <param name="outputFormat">optional object OutputFormat</param>
		/// <param name="to">optional object To</param>
		/// <param name="cc">optional object Cc</param>
		/// <param name="bcc">optional object Bcc</param>
		/// <param name="subject">optional object Subject</param>
		/// <param name="messageText">optional object MessageText</param>
		/// <param name="editMessage">optional object EditMessage</param>
		/// <param name="templateFile">optional object TemplateFile</param>
		[SupportByVersionAttribute("Access", 9,10,11,12,14,15,16)]
		public void SendObject(object objectType, object objectName, object outputFormat, object to, object cc, object bcc, object subject, object messageText, object editMessage, object templateFile)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(objectType, objectName, outputFormat, to, cc, bcc, subject, messageText, editMessage, templateFile);
			Invoker.Method(this, "SendObject", paramsArray);
		}

		/// <summary>
		/// SupportByVersion Access 9, 10, 11, 12, 14, 15, 16
		/// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff197046.aspx
		/// </summary>
		[CustomMethodAttribute]
		[SupportByVersionAttribute("Access", 9,10,11,12,14,15,16)]
		public void SendObject()
		{
			object[] paramsArray = null;
			Invoker.Method(this, "SendObject", paramsArray);
		}

		/// <summary>
		/// SupportByVersion Access 9, 10, 11, 12, 14, 15, 16
		/// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff197046.aspx
		/// </summary>
		/// <param name="objectType">optional NetOffice.AccessApi.Enums.AcSendObjectType ObjectType = -1</param>
		[CustomMethodAttribute]
		[SupportByVersionAttribute("Access", 9,10,11,12,14,15,16)]
		public void SendObject(object objectType)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(objectType);
			Invoker.Method(this, "SendObject", paramsArray);
		}

		/// <summary>
		/// SupportByVersion Access 9, 10, 11, 12, 14, 15, 16
		/// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff197046.aspx
		/// </summary>
		/// <param name="objectType">optional NetOffice.AccessApi.Enums.AcSendObjectType ObjectType = -1</param>
		/// <param name="objectName">optional object ObjectName</param>
		[CustomMethodAttribute]
		[SupportByVersionAttribute("Access", 9,10,11,12,14,15,16)]
		public void SendObject(object objectType, object objectName)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(objectType, objectName);
			Invoker.Method(this, "SendObject", paramsArray);
		}

		/// <summary>
		/// SupportByVersion Access 9, 10, 11, 12, 14, 15, 16
		/// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff197046.aspx
		/// </summary>
		/// <param name="objectType">optional NetOffice.AccessApi.Enums.AcSendObjectType ObjectType = -1</param>
		/// <param name="objectName">optional object ObjectName</param>
		/// <param name="outputFormat">optional object OutputFormat</param>
		[CustomMethodAttribute]
		[SupportByVersionAttribute("Access", 9,10,11,12,14,15,16)]
		public void SendObject(object objectType, object objectName, object outputFormat)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(objectType, objectName, outputFormat);
			Invoker.Method(this, "SendObject", paramsArray);
		}

		/// <summary>
		/// SupportByVersion Access 9, 10, 11, 12, 14, 15, 16
		/// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff197046.aspx
		/// </summary>
		/// <param name="objectType">optional NetOffice.AccessApi.Enums.AcSendObjectType ObjectType = -1</param>
		/// <param name="objectName">optional object ObjectName</param>
		/// <param name="outputFormat">optional object OutputFormat</param>
		/// <param name="to">optional object To</param>
		[CustomMethodAttribute]
		[SupportByVersionAttribute("Access", 9,10,11,12,14,15,16)]
		public void SendObject(object objectType, object objectName, object outputFormat, object to)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(objectType, objectName, outputFormat, to);
			Invoker.Method(this, "SendObject", paramsArray);
		}

		/// <summary>
		/// SupportByVersion Access 9, 10, 11, 12, 14, 15, 16
		/// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff197046.aspx
		/// </summary>
		/// <param name="objectType">optional NetOffice.AccessApi.Enums.AcSendObjectType ObjectType = -1</param>
		/// <param name="objectName">optional object ObjectName</param>
		/// <param name="outputFormat">optional object OutputFormat</param>
		/// <param name="to">optional object To</param>
		/// <param name="cc">optional object Cc</param>
		[CustomMethodAttribute]
		[SupportByVersionAttribute("Access", 9,10,11,12,14,15,16)]
		public void SendObject(object objectType, object objectName, object outputFormat, object to, object cc)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(objectType, objectName, outputFormat, to, cc);
			Invoker.Method(this, "SendObject", paramsArray);
		}

		/// <summary>
		/// SupportByVersion Access 9, 10, 11, 12, 14, 15, 16
		/// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff197046.aspx
		/// </summary>
		/// <param name="objectType">optional NetOffice.AccessApi.Enums.AcSendObjectType ObjectType = -1</param>
		/// <param name="objectName">optional object ObjectName</param>
		/// <param name="outputFormat">optional object OutputFormat</param>
		/// <param name="to">optional object To</param>
		/// <param name="cc">optional object Cc</param>
		/// <param name="bcc">optional object Bcc</param>
		[CustomMethodAttribute]
		[SupportByVersionAttribute("Access", 9,10,11,12,14,15,16)]
		public void SendObject(object objectType, object objectName, object outputFormat, object to, object cc, object bcc)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(objectType, objectName, outputFormat, to, cc, bcc);
			Invoker.Method(this, "SendObject", paramsArray);
		}

		/// <summary>
		/// SupportByVersion Access 9, 10, 11, 12, 14, 15, 16
		/// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff197046.aspx
		/// </summary>
		/// <param name="objectType">optional NetOffice.AccessApi.Enums.AcSendObjectType ObjectType = -1</param>
		/// <param name="objectName">optional object ObjectName</param>
		/// <param name="outputFormat">optional object OutputFormat</param>
		/// <param name="to">optional object To</param>
		/// <param name="cc">optional object Cc</param>
		/// <param name="bcc">optional object Bcc</param>
		/// <param name="subject">optional object Subject</param>
		[CustomMethodAttribute]
		[SupportByVersionAttribute("Access", 9,10,11,12,14,15,16)]
		public void SendObject(object objectType, object objectName, object outputFormat, object to, object cc, object bcc, object subject)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(objectType, objectName, outputFormat, to, cc, bcc, subject);
			Invoker.Method(this, "SendObject", paramsArray);
		}

		/// <summary>
		/// SupportByVersion Access 9, 10, 11, 12, 14, 15, 16
		/// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff197046.aspx
		/// </summary>
		/// <param name="objectType">optional NetOffice.AccessApi.Enums.AcSendObjectType ObjectType = -1</param>
		/// <param name="objectName">optional object ObjectName</param>
		/// <param name="outputFormat">optional object OutputFormat</param>
		/// <param name="to">optional object To</param>
		/// <param name="cc">optional object Cc</param>
		/// <param name="bcc">optional object Bcc</param>
		/// <param name="subject">optional object Subject</param>
		/// <param name="messageText">optional object MessageText</param>
		[CustomMethodAttribute]
		[SupportByVersionAttribute("Access", 9,10,11,12,14,15,16)]
		public void SendObject(object objectType, object objectName, object outputFormat, object to, object cc, object bcc, object subject, object messageText)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(objectType, objectName, outputFormat, to, cc, bcc, subject, messageText);
			Invoker.Method(this, "SendObject", paramsArray);
		}

		/// <summary>
		/// SupportByVersion Access 9, 10, 11, 12, 14, 15, 16
		/// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff197046.aspx
		/// </summary>
		/// <param name="objectType">optional NetOffice.AccessApi.Enums.AcSendObjectType ObjectType = -1</param>
		/// <param name="objectName">optional object ObjectName</param>
		/// <param name="outputFormat">optional object OutputFormat</param>
		/// <param name="to">optional object To</param>
		/// <param name="cc">optional object Cc</param>
		/// <param name="bcc">optional object Bcc</param>
		/// <param name="subject">optional object Subject</param>
		/// <param name="messageText">optional object MessageText</param>
		/// <param name="editMessage">optional object EditMessage</param>
		[CustomMethodAttribute]
		[SupportByVersionAttribute("Access", 9,10,11,12,14,15,16)]
		public void SendObject(object objectType, object objectName, object outputFormat, object to, object cc, object bcc, object subject, object messageText, object editMessage)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(objectType, objectName, outputFormat, to, cc, bcc, subject, messageText, editMessage);
			Invoker.Method(this, "SendObject", paramsArray);
		}

		/// <summary>
		/// SupportByVersion Access 9, 10, 11, 12, 14, 15, 16
		/// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff194957.aspx
		/// </summary>
		/// <param name="toolbarName">object ToolbarName</param>
		/// <param name="show">optional NetOffice.AccessApi.Enums.AcShowToolbar Show = 0</param>
		[SupportByVersionAttribute("Access", 9,10,11,12,14,15,16)]
		public void ShowToolbar(object toolbarName, object show)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(toolbarName, show);
			Invoker.Method(this, "ShowToolbar", paramsArray);
		}

		/// <summary>
		/// SupportByVersion Access 9, 10, 11, 12, 14, 15, 16
		/// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff194957.aspx
		/// </summary>
		/// <param name="toolbarName">object ToolbarName</param>
		[CustomMethodAttribute]
		[SupportByVersionAttribute("Access", 9,10,11,12,14,15,16)]
		public void ShowToolbar(object toolbarName)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(toolbarName);
			Invoker.Method(this, "ShowToolbar", paramsArray);
		}

		/// <summary>
		/// SupportByVersion Access 9, 10, 11, 12, 14, 15, 16
		/// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff196435.aspx
		/// </summary>
		/// <param name="objectType">optional NetOffice.AccessApi.Enums.AcObjectType ObjectType = -1</param>
		/// <param name="objectName">optional object ObjectName</param>
		[SupportByVersionAttribute("Access", 9,10,11,12,14,15,16)]
		public void Save(object objectType, object objectName)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(objectType, objectName);
			Invoker.Method(this, "Save", paramsArray);
		}

		/// <summary>
		/// SupportByVersion Access 9, 10, 11, 12, 14, 15, 16
		/// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff196435.aspx
		/// </summary>
		[CustomMethodAttribute]
		[SupportByVersionAttribute("Access", 9,10,11,12,14,15,16)]
		public void Save()
		{
			object[] paramsArray = null;
			Invoker.Method(this, "Save", paramsArray);
		}

		/// <summary>
		/// SupportByVersion Access 9, 10, 11, 12, 14, 15, 16
		/// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff196435.aspx
		/// </summary>
		/// <param name="objectType">optional NetOffice.AccessApi.Enums.AcObjectType ObjectType = -1</param>
		[CustomMethodAttribute]
		[SupportByVersionAttribute("Access", 9,10,11,12,14,15,16)]
		public void Save(object objectType)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(objectType);
			Invoker.Method(this, "Save", paramsArray);
		}

		/// <summary>
		/// SupportByVersion Access 9, 10, 11, 12, 14, 15, 16
		/// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff195265.aspx
		/// </summary>
		/// <param name="menuIndex">object MenuIndex</param>
		/// <param name="commandIndex">optional object CommandIndex</param>
		/// <param name="subcommandIndex">optional object SubcommandIndex</param>
		/// <param name="flag">optional object Flag</param>
		[SupportByVersionAttribute("Access", 9,10,11,12,14,15,16)]
		public void SetMenuItem(object menuIndex, object commandIndex, object subcommandIndex, object flag)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(menuIndex, commandIndex, subcommandIndex, flag);
			Invoker.Method(this, "SetMenuItem", paramsArray);
		}

		/// <summary>
		/// SupportByVersion Access 9, 10, 11, 12, 14, 15, 16
		/// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff195265.aspx
		/// </summary>
		/// <param name="menuIndex">object MenuIndex</param>
		[CustomMethodAttribute]
		[SupportByVersionAttribute("Access", 9,10,11,12,14,15,16)]
		public void SetMenuItem(object menuIndex)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(menuIndex);
			Invoker.Method(this, "SetMenuItem", paramsArray);
		}

		/// <summary>
		/// SupportByVersion Access 9, 10, 11, 12, 14, 15, 16
		/// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff195265.aspx
		/// </summary>
		/// <param name="menuIndex">object MenuIndex</param>
		/// <param name="commandIndex">optional object CommandIndex</param>
		[CustomMethodAttribute]
		[SupportByVersionAttribute("Access", 9,10,11,12,14,15,16)]
		public void SetMenuItem(object menuIndex, object commandIndex)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(menuIndex, commandIndex);
			Invoker.Method(this, "SetMenuItem", paramsArray);
		}

		/// <summary>
		/// SupportByVersion Access 9, 10, 11, 12, 14, 15, 16
		/// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff195265.aspx
		/// </summary>
		/// <param name="menuIndex">object MenuIndex</param>
		/// <param name="commandIndex">optional object CommandIndex</param>
		/// <param name="subcommandIndex">optional object SubcommandIndex</param>
		[CustomMethodAttribute]
		[SupportByVersionAttribute("Access", 9,10,11,12,14,15,16)]
		public void SetMenuItem(object menuIndex, object commandIndex, object subcommandIndex)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(menuIndex, commandIndex, subcommandIndex);
			Invoker.Method(this, "SetMenuItem", paramsArray);
		}

		/// <summary>
		/// SupportByVersion Access 9, 10, 11, 12, 14, 15, 16
		/// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff194612.aspx
		/// </summary>
		/// <param name="command">NetOffice.AccessApi.Enums.AcCommand Command</param>
		[SupportByVersionAttribute("Access", 9,10,11,12,14,15,16)]
		public void RunCommand(NetOffice.AccessApi.Enums.AcCommand command)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(command);
			Invoker.Method(this, "RunCommand", paramsArray);
		}

		/// <summary>
		/// SupportByVersion Access 9, 10, 11, 12, 14, 15, 16
		/// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff845421.aspx
		/// </summary>
		/// <param name="dataAccessPageName">object DataAccessPageName</param>
		/// <param name="view">optional NetOffice.AccessApi.Enums.AcDataAccessPageView View = 0</param>
		[SupportByVersionAttribute("Access", 9,10,11,12,14,15,16)]
		public void OpenDataAccessPage(object dataAccessPageName, object view)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(dataAccessPageName, view);
			Invoker.Method(this, "OpenDataAccessPage", paramsArray);
		}

		/// <summary>
		/// SupportByVersion Access 9, 10, 11, 12, 14, 15, 16
		/// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff845421.aspx
		/// </summary>
		/// <param name="dataAccessPageName">object DataAccessPageName</param>
		[CustomMethodAttribute]
		[SupportByVersionAttribute("Access", 9,10,11,12,14,15,16)]
		public void OpenDataAccessPage(object dataAccessPageName)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(dataAccessPageName);
			Invoker.Method(this, "OpenDataAccessPage", paramsArray);
		}

		/// <summary>
		/// SupportByVersion Access 9, 10, 11, 12, 14, 15, 16
		/// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff197347.aspx
		/// </summary>
		/// <param name="viewName">object ViewName</param>
		/// <param name="view">optional NetOffice.AccessApi.Enums.AcView View = 0</param>
		/// <param name="dataMode">optional NetOffice.AccessApi.Enums.AcOpenDataMode DataMode = 1</param>
		[SupportByVersionAttribute("Access", 9,10,11,12,14,15,16)]
		public void OpenView(object viewName, object view, object dataMode)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(viewName, view, dataMode);
			Invoker.Method(this, "OpenView", paramsArray);
		}

		/// <summary>
		/// SupportByVersion Access 9, 10, 11, 12, 14, 15, 16
		/// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff197347.aspx
		/// </summary>
		/// <param name="viewName">object ViewName</param>
		[CustomMethodAttribute]
		[SupportByVersionAttribute("Access", 9,10,11,12,14,15,16)]
		public void OpenView(object viewName)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(viewName);
			Invoker.Method(this, "OpenView", paramsArray);
		}

		/// <summary>
		/// SupportByVersion Access 9, 10, 11, 12, 14, 15, 16
		/// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff197347.aspx
		/// </summary>
		/// <param name="viewName">object ViewName</param>
		/// <param name="view">optional NetOffice.AccessApi.Enums.AcView View = 0</param>
		[CustomMethodAttribute]
		[SupportByVersionAttribute("Access", 9,10,11,12,14,15,16)]
		public void OpenView(object viewName, object view)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(viewName, view);
			Invoker.Method(this, "OpenView", paramsArray);
		}

		/// <summary>
		/// SupportByVersion Access 9, 10, 11, 12, 14, 15, 16
		/// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff821439.aspx
		/// </summary>
		/// <param name="diagramName">object DiagramName</param>
		[SupportByVersionAttribute("Access", 9,10,11,12,14,15,16)]
		public void OpenDiagram(object diagramName)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(diagramName);
			Invoker.Method(this, "OpenDiagram", paramsArray);
		}

		/// <summary>
		/// SupportByVersion Access 9, 10, 11, 12, 14, 15, 16
		/// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff197412.aspx
		/// </summary>
		/// <param name="procedureName">object ProcedureName</param>
		/// <param name="view">optional NetOffice.AccessApi.Enums.AcView View = 0</param>
		/// <param name="dataMode">optional NetOffice.AccessApi.Enums.AcOpenDataMode DataMode = 1</param>
		[SupportByVersionAttribute("Access", 9,10,11,12,14,15,16)]
		public void OpenStoredProcedure(object procedureName, object view, object dataMode)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(procedureName, view, dataMode);
			Invoker.Method(this, "OpenStoredProcedure", paramsArray);
		}

		/// <summary>
		/// SupportByVersion Access 9, 10, 11, 12, 14, 15, 16
		/// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff197412.aspx
		/// </summary>
		/// <param name="procedureName">object ProcedureName</param>
		[CustomMethodAttribute]
		[SupportByVersionAttribute("Access", 9,10,11,12,14,15,16)]
		public void OpenStoredProcedure(object procedureName)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(procedureName);
			Invoker.Method(this, "OpenStoredProcedure", paramsArray);
		}

		/// <summary>
		/// SupportByVersion Access 9, 10, 11, 12, 14, 15, 16
		/// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff197412.aspx
		/// </summary>
		/// <param name="procedureName">object ProcedureName</param>
		/// <param name="view">optional NetOffice.AccessApi.Enums.AcView View = 0</param>
		[CustomMethodAttribute]
		[SupportByVersionAttribute("Access", 9,10,11,12,14,15,16)]
		public void OpenStoredProcedure(object procedureName, object view)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(procedureName, view);
			Invoker.Method(this, "OpenStoredProcedure", paramsArray);
		}

		/// <summary>
		/// SupportByVersion Access 10, 11, 12, 14, 15, 16
		/// 
		/// </summary>
		/// <param name="reportName">object ReportName</param>
		/// <param name="view">optional NetOffice.AccessApi.Enums.AcView View = 0</param>
		/// <param name="filterName">optional object FilterName</param>
		/// <param name="whereCondition">optional object WhereCondition</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		[SupportByVersionAttribute("Access", 10,11,12,14,15,16)]
		public void OpenReportOld0(object reportName, object view, object filterName, object whereCondition)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(reportName, view, filterName, whereCondition);
			Invoker.Method(this, "OpenReportOld0", paramsArray);
		}

		/// <summary>
		/// SupportByVersion Access 10, 11, 12, 14, 15, 16
		/// 
		/// </summary>
		/// <param name="reportName">object ReportName</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		[CustomMethodAttribute]
		[SupportByVersionAttribute("Access", 10,11,12,14,15,16)]
		public void OpenReportOld0(object reportName)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(reportName);
			Invoker.Method(this, "OpenReportOld0", paramsArray);
		}

		/// <summary>
		/// SupportByVersion Access 10, 11, 12, 14, 15, 16
		/// 
		/// </summary>
		/// <param name="reportName">object ReportName</param>
		/// <param name="view">optional NetOffice.AccessApi.Enums.AcView View = 0</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		[CustomMethodAttribute]
		[SupportByVersionAttribute("Access", 10,11,12,14,15,16)]
		public void OpenReportOld0(object reportName, object view)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(reportName, view);
			Invoker.Method(this, "OpenReportOld0", paramsArray);
		}

		/// <summary>
		/// SupportByVersion Access 10, 11, 12, 14, 15, 16
		/// 
		/// </summary>
		/// <param name="reportName">object ReportName</param>
		/// <param name="view">optional NetOffice.AccessApi.Enums.AcView View = 0</param>
		/// <param name="filterName">optional object FilterName</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		[CustomMethodAttribute]
		[SupportByVersionAttribute("Access", 10,11,12,14,15,16)]
		public void OpenReportOld0(object reportName, object view, object filterName)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(reportName, view, filterName);
			Invoker.Method(this, "OpenReportOld0", paramsArray);
		}

		/// <summary>
		/// SupportByVersion Access 10, 11, 12, 14, 15, 16
		/// 
		/// </summary>
		/// <param name="objectType">NetOffice.AccessApi.Enums.AcOutputObjectType ObjectType</param>
		/// <param name="objectName">optional object ObjectName</param>
		/// <param name="outputFormat">optional object OutputFormat</param>
		/// <param name="outputFile">optional object OutputFile</param>
		/// <param name="autoStart">optional object AutoStart</param>
		/// <param name="templateFile">optional object TemplateFile</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		[SupportByVersionAttribute("Access", 10,11,12,14,15,16)]
		public void OutputToOld0(NetOffice.AccessApi.Enums.AcOutputObjectType objectType, object objectName, object outputFormat, object outputFile, object autoStart, object templateFile)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(objectType, objectName, outputFormat, outputFile, autoStart, templateFile);
			Invoker.Method(this, "OutputToOld0", paramsArray);
		}

		/// <summary>
		/// SupportByVersion Access 10, 11, 12, 14, 15, 16
		/// 
		/// </summary>
		/// <param name="objectType">NetOffice.AccessApi.Enums.AcOutputObjectType ObjectType</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		[CustomMethodAttribute]
		[SupportByVersionAttribute("Access", 10,11,12,14,15,16)]
		public void OutputToOld0(NetOffice.AccessApi.Enums.AcOutputObjectType objectType)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(objectType);
			Invoker.Method(this, "OutputToOld0", paramsArray);
		}

		/// <summary>
		/// SupportByVersion Access 10, 11, 12, 14, 15, 16
		/// 
		/// </summary>
		/// <param name="objectType">NetOffice.AccessApi.Enums.AcOutputObjectType ObjectType</param>
		/// <param name="objectName">optional object ObjectName</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		[CustomMethodAttribute]
		[SupportByVersionAttribute("Access", 10,11,12,14,15,16)]
		public void OutputToOld0(NetOffice.AccessApi.Enums.AcOutputObjectType objectType, object objectName)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(objectType, objectName);
			Invoker.Method(this, "OutputToOld0", paramsArray);
		}

		/// <summary>
		/// SupportByVersion Access 10, 11, 12, 14, 15, 16
		/// 
		/// </summary>
		/// <param name="objectType">NetOffice.AccessApi.Enums.AcOutputObjectType ObjectType</param>
		/// <param name="objectName">optional object ObjectName</param>
		/// <param name="outputFormat">optional object OutputFormat</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		[CustomMethodAttribute]
		[SupportByVersionAttribute("Access", 10,11,12,14,15,16)]
		public void OutputToOld0(NetOffice.AccessApi.Enums.AcOutputObjectType objectType, object objectName, object outputFormat)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(objectType, objectName, outputFormat);
			Invoker.Method(this, "OutputToOld0", paramsArray);
		}

		/// <summary>
		/// SupportByVersion Access 10, 11, 12, 14, 15, 16
		/// 
		/// </summary>
		/// <param name="objectType">NetOffice.AccessApi.Enums.AcOutputObjectType ObjectType</param>
		/// <param name="objectName">optional object ObjectName</param>
		/// <param name="outputFormat">optional object OutputFormat</param>
		/// <param name="outputFile">optional object OutputFile</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		[CustomMethodAttribute]
		[SupportByVersionAttribute("Access", 10,11,12,14,15,16)]
		public void OutputToOld0(NetOffice.AccessApi.Enums.AcOutputObjectType objectType, object objectName, object outputFormat, object outputFile)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(objectType, objectName, outputFormat, outputFile);
			Invoker.Method(this, "OutputToOld0", paramsArray);
		}

		/// <summary>
		/// SupportByVersion Access 10, 11, 12, 14, 15, 16
		/// 
		/// </summary>
		/// <param name="objectType">NetOffice.AccessApi.Enums.AcOutputObjectType ObjectType</param>
		/// <param name="objectName">optional object ObjectName</param>
		/// <param name="outputFormat">optional object OutputFormat</param>
		/// <param name="outputFile">optional object OutputFile</param>
		/// <param name="autoStart">optional object AutoStart</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		[CustomMethodAttribute]
		[SupportByVersionAttribute("Access", 10,11,12,14,15,16)]
		public void OutputToOld0(NetOffice.AccessApi.Enums.AcOutputObjectType objectType, object objectName, object outputFormat, object outputFile, object autoStart)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(objectType, objectName, outputFormat, outputFile, autoStart);
			Invoker.Method(this, "OutputToOld0", paramsArray);
		}

		/// <summary>
		/// SupportByVersion Access 10, 11, 12, 14, 15, 16
		/// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff835048.aspx
		/// </summary>
		/// <param name="server">object Server</param>
		/// <param name="database">object Database</param>
		/// <param name="useTrustedConnection">optional object UseTrustedConnection</param>
		/// <param name="login">optional object Login</param>
		/// <param name="password">optional object Password</param>
		/// <param name="transferCopyData">optional object TransferCopyData</param>
		[SupportByVersionAttribute("Access", 10,11,12,14,15,16)]
		public void TransferSQLDatabase(object server, object database, object useTrustedConnection, object login, object password, object transferCopyData)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(server, database, useTrustedConnection, login, password, transferCopyData);
			Invoker.Method(this, "TransferSQLDatabase", paramsArray);
		}

		/// <summary>
		/// SupportByVersion Access 10, 11, 12, 14, 15, 16
		/// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff835048.aspx
		/// </summary>
		/// <param name="server">object Server</param>
		/// <param name="database">object Database</param>
		[CustomMethodAttribute]
		[SupportByVersionAttribute("Access", 10,11,12,14,15,16)]
		public void TransferSQLDatabase(object server, object database)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(server, database);
			Invoker.Method(this, "TransferSQLDatabase", paramsArray);
		}

		/// <summary>
		/// SupportByVersion Access 10, 11, 12, 14, 15, 16
		/// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff835048.aspx
		/// </summary>
		/// <param name="server">object Server</param>
		/// <param name="database">object Database</param>
		/// <param name="useTrustedConnection">optional object UseTrustedConnection</param>
		[CustomMethodAttribute]
		[SupportByVersionAttribute("Access", 10,11,12,14,15,16)]
		public void TransferSQLDatabase(object server, object database, object useTrustedConnection)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(server, database, useTrustedConnection);
			Invoker.Method(this, "TransferSQLDatabase", paramsArray);
		}

		/// <summary>
		/// SupportByVersion Access 10, 11, 12, 14, 15, 16
		/// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff835048.aspx
		/// </summary>
		/// <param name="server">object Server</param>
		/// <param name="database">object Database</param>
		/// <param name="useTrustedConnection">optional object UseTrustedConnection</param>
		/// <param name="login">optional object Login</param>
		[CustomMethodAttribute]
		[SupportByVersionAttribute("Access", 10,11,12,14,15,16)]
		public void TransferSQLDatabase(object server, object database, object useTrustedConnection, object login)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(server, database, useTrustedConnection, login);
			Invoker.Method(this, "TransferSQLDatabase", paramsArray);
		}

		/// <summary>
		/// SupportByVersion Access 10, 11, 12, 14, 15, 16
		/// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff835048.aspx
		/// </summary>
		/// <param name="server">object Server</param>
		/// <param name="database">object Database</param>
		/// <param name="useTrustedConnection">optional object UseTrustedConnection</param>
		/// <param name="login">optional object Login</param>
		/// <param name="password">optional object Password</param>
		[CustomMethodAttribute]
		[SupportByVersionAttribute("Access", 10,11,12,14,15,16)]
		public void TransferSQLDatabase(object server, object database, object useTrustedConnection, object login, object password)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(server, database, useTrustedConnection, login, password);
			Invoker.Method(this, "TransferSQLDatabase", paramsArray);
		}

		/// <summary>
		/// SupportByVersion Access 10, 11, 12, 14, 15, 16
		/// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff845497.aspx
		/// </summary>
		/// <param name="databaseFileName">object DatabaseFileName</param>
		/// <param name="overwriteExistingFile">optional object OverwriteExistingFile</param>
		/// <param name="disconnectAllUsers">optional object DisconnectAllUsers</param>
		[SupportByVersionAttribute("Access", 10,11,12,14,15,16)]
		public void CopyDatabaseFile(object databaseFileName, object overwriteExistingFile, object disconnectAllUsers)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(databaseFileName, overwriteExistingFile, disconnectAllUsers);
			Invoker.Method(this, "CopyDatabaseFile", paramsArray);
		}

		/// <summary>
		/// SupportByVersion Access 10, 11, 12, 14, 15, 16
		/// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff845497.aspx
		/// </summary>
		/// <param name="databaseFileName">object DatabaseFileName</param>
		[CustomMethodAttribute]
		[SupportByVersionAttribute("Access", 10,11,12,14,15,16)]
		public void CopyDatabaseFile(object databaseFileName)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(databaseFileName);
			Invoker.Method(this, "CopyDatabaseFile", paramsArray);
		}

		/// <summary>
		/// SupportByVersion Access 10, 11, 12, 14, 15, 16
		/// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff845497.aspx
		/// </summary>
		/// <param name="databaseFileName">object DatabaseFileName</param>
		/// <param name="overwriteExistingFile">optional object OverwriteExistingFile</param>
		[CustomMethodAttribute]
		[SupportByVersionAttribute("Access", 10,11,12,14,15,16)]
		public void CopyDatabaseFile(object databaseFileName, object overwriteExistingFile)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(databaseFileName, overwriteExistingFile);
			Invoker.Method(this, "CopyDatabaseFile", paramsArray);
		}

		/// <summary>
		/// SupportByVersion Access 10, 11, 12, 14, 15, 16
		/// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff194192.aspx
		/// </summary>
		/// <param name="functionName">object FunctionName</param>
		/// <param name="view">optional NetOffice.AccessApi.Enums.AcView View = 0</param>
		/// <param name="dataMode">optional NetOffice.AccessApi.Enums.AcOpenDataMode DataMode = 1</param>
		[SupportByVersionAttribute("Access", 10,11,12,14,15,16)]
		public void OpenFunction(object functionName, object view, object dataMode)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(functionName, view, dataMode);
			Invoker.Method(this, "OpenFunction", paramsArray);
		}

		/// <summary>
		/// SupportByVersion Access 10, 11, 12, 14, 15, 16
		/// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff194192.aspx
		/// </summary>
		/// <param name="functionName">object FunctionName</param>
		[CustomMethodAttribute]
		[SupportByVersionAttribute("Access", 10,11,12,14,15,16)]
		public void OpenFunction(object functionName)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(functionName);
			Invoker.Method(this, "OpenFunction", paramsArray);
		}

		/// <summary>
		/// SupportByVersion Access 10, 11, 12, 14, 15, 16
		/// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff194192.aspx
		/// </summary>
		/// <param name="functionName">object FunctionName</param>
		/// <param name="view">optional NetOffice.AccessApi.Enums.AcView View = 0</param>
		[CustomMethodAttribute]
		[SupportByVersionAttribute("Access", 10,11,12,14,15,16)]
		public void OpenFunction(object functionName, object view)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(functionName, view);
			Invoker.Method(this, "OpenFunction", paramsArray);
		}

		/// <summary>
		/// SupportByVersion Access 12, 14, 15, 16
		/// 
		/// </summary>
		/// <param name="filterName">optional object FilterName</param>
		/// <param name="whereCondition">optional object WhereCondition</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		[SupportByVersionAttribute("Access", 12,14,15,16)]
		public void ApplyFilterOld0(object filterName, object whereCondition)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(filterName, whereCondition);
			Invoker.Method(this, "ApplyFilterOld0", paramsArray);
		}

		/// <summary>
		/// SupportByVersion Access 12, 14, 15, 16
		/// 
		/// </summary>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		[CustomMethodAttribute]
		[SupportByVersionAttribute("Access", 12,14,15,16)]
		public void ApplyFilterOld0()
		{
			object[] paramsArray = null;
			Invoker.Method(this, "ApplyFilterOld0", paramsArray);
		}

		/// <summary>
		/// SupportByVersion Access 12, 14, 15, 16
		/// 
		/// </summary>
		/// <param name="filterName">optional object FilterName</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		[CustomMethodAttribute]
		[SupportByVersionAttribute("Access", 12,14,15,16)]
		public void ApplyFilterOld0(object filterName)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(filterName);
			Invoker.Method(this, "ApplyFilterOld0", paramsArray);
		}

		/// <summary>
		/// SupportByVersion Access 12, 14, 15, 16
		/// 
		/// </summary>
		/// <param name="objectType">NetOffice.AccessApi.Enums.AcOutputObjectType ObjectType</param>
		/// <param name="objectName">optional object ObjectName</param>
		/// <param name="outputFormat">optional object OutputFormat</param>
		/// <param name="outputFile">optional object OutputFile</param>
		/// <param name="autoStart">optional object AutoStart</param>
		/// <param name="templateFile">optional object TemplateFile</param>
		/// <param name="encoding">optional object Encoding</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		[SupportByVersionAttribute("Access", 12,14,15,16)]
		public void OutputToOld1(NetOffice.AccessApi.Enums.AcOutputObjectType objectType, object objectName, object outputFormat, object outputFile, object autoStart, object templateFile, object encoding)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(objectType, objectName, outputFormat, outputFile, autoStart, templateFile, encoding);
			Invoker.Method(this, "OutputToOld1", paramsArray);
		}

		/// <summary>
		/// SupportByVersion Access 12, 14, 15, 16
		/// 
		/// </summary>
		/// <param name="objectType">NetOffice.AccessApi.Enums.AcOutputObjectType ObjectType</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		[CustomMethodAttribute]
		[SupportByVersionAttribute("Access", 12,14,15,16)]
		public void OutputToOld1(NetOffice.AccessApi.Enums.AcOutputObjectType objectType)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(objectType);
			Invoker.Method(this, "OutputToOld1", paramsArray);
		}

		/// <summary>
		/// SupportByVersion Access 12, 14, 15, 16
		/// 
		/// </summary>
		/// <param name="objectType">NetOffice.AccessApi.Enums.AcOutputObjectType ObjectType</param>
		/// <param name="objectName">optional object ObjectName</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		[CustomMethodAttribute]
		[SupportByVersionAttribute("Access", 12,14,15,16)]
		public void OutputToOld1(NetOffice.AccessApi.Enums.AcOutputObjectType objectType, object objectName)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(objectType, objectName);
			Invoker.Method(this, "OutputToOld1", paramsArray);
		}

		/// <summary>
		/// SupportByVersion Access 12, 14, 15, 16
		/// 
		/// </summary>
		/// <param name="objectType">NetOffice.AccessApi.Enums.AcOutputObjectType ObjectType</param>
		/// <param name="objectName">optional object ObjectName</param>
		/// <param name="outputFormat">optional object OutputFormat</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		[CustomMethodAttribute]
		[SupportByVersionAttribute("Access", 12,14,15,16)]
		public void OutputToOld1(NetOffice.AccessApi.Enums.AcOutputObjectType objectType, object objectName, object outputFormat)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(objectType, objectName, outputFormat);
			Invoker.Method(this, "OutputToOld1", paramsArray);
		}

		/// <summary>
		/// SupportByVersion Access 12, 14, 15, 16
		/// 
		/// </summary>
		/// <param name="objectType">NetOffice.AccessApi.Enums.AcOutputObjectType ObjectType</param>
		/// <param name="objectName">optional object ObjectName</param>
		/// <param name="outputFormat">optional object OutputFormat</param>
		/// <param name="outputFile">optional object OutputFile</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		[CustomMethodAttribute]
		[SupportByVersionAttribute("Access", 12,14,15,16)]
		public void OutputToOld1(NetOffice.AccessApi.Enums.AcOutputObjectType objectType, object objectName, object outputFormat, object outputFile)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(objectType, objectName, outputFormat, outputFile);
			Invoker.Method(this, "OutputToOld1", paramsArray);
		}

		/// <summary>
		/// SupportByVersion Access 12, 14, 15, 16
		/// 
		/// </summary>
		/// <param name="objectType">NetOffice.AccessApi.Enums.AcOutputObjectType ObjectType</param>
		/// <param name="objectName">optional object ObjectName</param>
		/// <param name="outputFormat">optional object OutputFormat</param>
		/// <param name="outputFile">optional object OutputFile</param>
		/// <param name="autoStart">optional object AutoStart</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		[CustomMethodAttribute]
		[SupportByVersionAttribute("Access", 12,14,15,16)]
		public void OutputToOld1(NetOffice.AccessApi.Enums.AcOutputObjectType objectType, object objectName, object outputFormat, object outputFile, object autoStart)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(objectType, objectName, outputFormat, outputFile, autoStart);
			Invoker.Method(this, "OutputToOld1", paramsArray);
		}

		/// <summary>
		/// SupportByVersion Access 12, 14, 15, 16
		/// 
		/// </summary>
		/// <param name="objectType">NetOffice.AccessApi.Enums.AcOutputObjectType ObjectType</param>
		/// <param name="objectName">optional object ObjectName</param>
		/// <param name="outputFormat">optional object OutputFormat</param>
		/// <param name="outputFile">optional object OutputFile</param>
		/// <param name="autoStart">optional object AutoStart</param>
		/// <param name="templateFile">optional object TemplateFile</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		[CustomMethodAttribute]
		[SupportByVersionAttribute("Access", 12,14,15,16)]
		public void OutputToOld1(NetOffice.AccessApi.Enums.AcOutputObjectType objectType, object objectName, object outputFormat, object outputFile, object autoStart, object templateFile)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(objectType, objectName, outputFormat, outputFile, autoStart, templateFile);
			Invoker.Method(this, "OutputToOld1", paramsArray);
		}

		/// <summary>
		/// SupportByVersion Access 12, 14, 15, 16
		/// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff198137.aspx
		/// </summary>
		/// <param name="transferType">NetOffice.AccessApi.Enums.AcSharePointListTransferType TransferType</param>
		/// <param name="siteAddress">object SiteAddress</param>
		/// <param name="listID">object ListID</param>
		/// <param name="viewID">optional object ViewID</param>
		/// <param name="tableName">optional object TableName</param>
		/// <param name="getLookupDisplayValues">optional object GetLookupDisplayValues</param>
		[SupportByVersionAttribute("Access", 12,14,15,16)]
		public void TransferSharePointList(NetOffice.AccessApi.Enums.AcSharePointListTransferType transferType, object siteAddress, object listID, object viewID, object tableName, object getLookupDisplayValues)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(transferType, siteAddress, listID, viewID, tableName, getLookupDisplayValues);
			Invoker.Method(this, "TransferSharePointList", paramsArray);
		}

		/// <summary>
		/// SupportByVersion Access 12, 14, 15, 16
		/// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff198137.aspx
		/// </summary>
		/// <param name="transferType">NetOffice.AccessApi.Enums.AcSharePointListTransferType TransferType</param>
		/// <param name="siteAddress">object SiteAddress</param>
		/// <param name="listID">object ListID</param>
		[CustomMethodAttribute]
		[SupportByVersionAttribute("Access", 12,14,15,16)]
		public void TransferSharePointList(NetOffice.AccessApi.Enums.AcSharePointListTransferType transferType, object siteAddress, object listID)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(transferType, siteAddress, listID);
			Invoker.Method(this, "TransferSharePointList", paramsArray);
		}

		/// <summary>
		/// SupportByVersion Access 12, 14, 15, 16
		/// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff198137.aspx
		/// </summary>
		/// <param name="transferType">NetOffice.AccessApi.Enums.AcSharePointListTransferType TransferType</param>
		/// <param name="siteAddress">object SiteAddress</param>
		/// <param name="listID">object ListID</param>
		/// <param name="viewID">optional object ViewID</param>
		[CustomMethodAttribute]
		[SupportByVersionAttribute("Access", 12,14,15,16)]
		public void TransferSharePointList(NetOffice.AccessApi.Enums.AcSharePointListTransferType transferType, object siteAddress, object listID, object viewID)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(transferType, siteAddress, listID, viewID);
			Invoker.Method(this, "TransferSharePointList", paramsArray);
		}

		/// <summary>
		/// SupportByVersion Access 12, 14, 15, 16
		/// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff198137.aspx
		/// </summary>
		/// <param name="transferType">NetOffice.AccessApi.Enums.AcSharePointListTransferType TransferType</param>
		/// <param name="siteAddress">object SiteAddress</param>
		/// <param name="listID">object ListID</param>
		/// <param name="viewID">optional object ViewID</param>
		/// <param name="tableName">optional object TableName</param>
		[CustomMethodAttribute]
		[SupportByVersionAttribute("Access", 12,14,15,16)]
		public void TransferSharePointList(NetOffice.AccessApi.Enums.AcSharePointListTransferType transferType, object siteAddress, object listID, object viewID, object tableName)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(transferType, siteAddress, listID, viewID, tableName);
			Invoker.Method(this, "TransferSharePointList", paramsArray);
		}

		/// <summary>
		/// SupportByVersion Access 12, 14, 15, 16
		/// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff844747.aspx
		/// </summary>
		[SupportByVersionAttribute("Access", 12,14,15,16)]
		public void CloseDatabase()
		{
			object[] paramsArray = null;
			Invoker.Method(this, "CloseDatabase", paramsArray);
		}

		/// <summary>
		/// SupportByVersion Access 12, 14, 15, 16
		/// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff191916.aspx
		/// </summary>
		/// <param name="category">optional object Category</param>
		/// <param name="group">optional object Group</param>
		[SupportByVersionAttribute("Access", 12,14,15,16)]
		public void NavigateTo(object category, object group)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(category, group);
			Invoker.Method(this, "NavigateTo", paramsArray);
		}

		/// <summary>
		/// SupportByVersion Access 12, 14, 15, 16
		/// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff191916.aspx
		/// </summary>
		[CustomMethodAttribute]
		[SupportByVersionAttribute("Access", 12,14,15,16)]
		public void NavigateTo()
		{
			object[] paramsArray = null;
			Invoker.Method(this, "NavigateTo", paramsArray);
		}

		/// <summary>
		/// SupportByVersion Access 12, 14, 15, 16
		/// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff191916.aspx
		/// </summary>
		/// <param name="category">optional object Category</param>
		[CustomMethodAttribute]
		[SupportByVersionAttribute("Access", 12,14,15,16)]
		public void NavigateTo(object category)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(category);
			Invoker.Method(this, "NavigateTo", paramsArray);
		}

		/// <summary>
		/// SupportByVersion Access 12, 14, 15, 16
		/// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff836254.aspx
		/// </summary>
		/// <param name="objectType">optional NetOffice.AccessApi.Enums.AcDataObjectType ObjectType = -1</param>
		/// <param name="objectName">optional object ObjectName</param>
		/// <param name="record">optional NetOffice.AccessApi.Enums.AcRecord Record = 2</param>
		/// <param name="whereCondition">optional object WhereCondition</param>
		[SupportByVersionAttribute("Access", 12,14,15,16)]
		public void SearchForRecord(object objectType, object objectName, object record, object whereCondition)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(objectType, objectName, record, whereCondition);
			Invoker.Method(this, "SearchForRecord", paramsArray);
		}

		/// <summary>
		/// SupportByVersion Access 12, 14, 15, 16
		/// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff836254.aspx
		/// </summary>
		[CustomMethodAttribute]
		[SupportByVersionAttribute("Access", 12,14,15,16)]
		public void SearchForRecord()
		{
			object[] paramsArray = null;
			Invoker.Method(this, "SearchForRecord", paramsArray);
		}

		/// <summary>
		/// SupportByVersion Access 12, 14, 15, 16
		/// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff836254.aspx
		/// </summary>
		/// <param name="objectType">optional NetOffice.AccessApi.Enums.AcDataObjectType ObjectType = -1</param>
		[CustomMethodAttribute]
		[SupportByVersionAttribute("Access", 12,14,15,16)]
		public void SearchForRecord(object objectType)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(objectType);
			Invoker.Method(this, "SearchForRecord", paramsArray);
		}

		/// <summary>
		/// SupportByVersion Access 12, 14, 15, 16
		/// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff836254.aspx
		/// </summary>
		/// <param name="objectType">optional NetOffice.AccessApi.Enums.AcDataObjectType ObjectType = -1</param>
		/// <param name="objectName">optional object ObjectName</param>
		[CustomMethodAttribute]
		[SupportByVersionAttribute("Access", 12,14,15,16)]
		public void SearchForRecord(object objectType, object objectName)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(objectType, objectName);
			Invoker.Method(this, "SearchForRecord", paramsArray);
		}

		/// <summary>
		/// SupportByVersion Access 12, 14, 15, 16
		/// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff836254.aspx
		/// </summary>
		/// <param name="objectType">optional NetOffice.AccessApi.Enums.AcDataObjectType ObjectType = -1</param>
		/// <param name="objectName">optional object ObjectName</param>
		/// <param name="record">optional NetOffice.AccessApi.Enums.AcRecord Record = 2</param>
		[CustomMethodAttribute]
		[SupportByVersionAttribute("Access", 12,14,15,16)]
		public void SearchForRecord(object objectType, object objectName, object record)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(objectType, objectName, record);
			Invoker.Method(this, "SearchForRecord", paramsArray);
		}

		/// <summary>
		/// SupportByVersion Access 12, 14, 15, 16
		/// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff192301.aspx
		/// </summary>
		/// <param name="controlName">object ControlName</param>
		/// <param name="property">optional NetOffice.AccessApi.Enums.AcProperty Property = 0</param>
		/// <param name="value">optional object Value</param>
		[SupportByVersionAttribute("Access", 12,14,15,16)]
		public void SetProperty(object controlName, object property, object value)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(controlName, property, value);
			Invoker.Method(this, "SetProperty", paramsArray);
		}

		/// <summary>
		/// SupportByVersion Access 12, 14, 15, 16
		/// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff192301.aspx
		/// </summary>
		/// <param name="controlName">object ControlName</param>
		[CustomMethodAttribute]
		[SupportByVersionAttribute("Access", 12,14,15,16)]
		public void SetProperty(object controlName)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(controlName);
			Invoker.Method(this, "SetProperty", paramsArray);
		}

		/// <summary>
		/// SupportByVersion Access 12, 14, 15, 16
		/// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff192301.aspx
		/// </summary>
		/// <param name="controlName">object ControlName</param>
		/// <param name="property">optional NetOffice.AccessApi.Enums.AcProperty Property = 0</param>
		[CustomMethodAttribute]
		[SupportByVersionAttribute("Access", 12,14,15,16)]
		public void SetProperty(object controlName, object property)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(controlName, property);
			Invoker.Method(this, "SetProperty", paramsArray);
		}

		/// <summary>
		/// SupportByVersion Access 12, 14, 15, 16
		/// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff837036.aspx
		/// </summary>
		[SupportByVersionAttribute("Access", 12,14,15,16)]
		public void SingleStep()
		{
			object[] paramsArray = null;
			Invoker.Method(this, "SingleStep", paramsArray);
		}

		/// <summary>
		/// SupportByVersion Access 12, 14, 15, 16
		/// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff191914.aspx
		/// </summary>
		[SupportByVersionAttribute("Access", 12,14,15,16)]
		public void ClearMacroError()
		{
			object[] paramsArray = null;
			Invoker.Method(this, "ClearMacroError", paramsArray);
		}

		/// <summary>
		/// SupportByVersion Access 12, 14, 15, 16
		/// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff821741.aspx
		/// </summary>
		/// <param name="show">object Show</param>
		/// <param name="category">optional object Category</param>
		[SupportByVersionAttribute("Access", 12,14,15,16)]
		public void SetDisplayedCategories(object show, object category)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(show, category);
			Invoker.Method(this, "SetDisplayedCategories", paramsArray);
		}

		/// <summary>
		/// SupportByVersion Access 12, 14, 15, 16
		/// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff821741.aspx
		/// </summary>
		/// <param name="show">object Show</param>
		[CustomMethodAttribute]
		[SupportByVersionAttribute("Access", 12,14,15,16)]
		public void SetDisplayedCategories(object show)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(show);
			Invoker.Method(this, "SetDisplayedCategories", paramsArray);
		}

		/// <summary>
		/// SupportByVersion Access 12, 14, 15, 16
		/// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff195088.aspx
		/// </summary>
		/// <param name="_lock">object Lock</param>
		[SupportByVersionAttribute("Access", 12,14,15,16)]
		public void LockNavigationPane(object _lock)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(_lock);
			Invoker.Method(this, "LockNavigationPane", paramsArray);
		}

		/// <summary>
		/// SupportByVersion Access 12, 14, 15, 16
		/// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff834375.aspx
		/// </summary>
		/// <param name="savedImportExportName">object SavedImportExportName</param>
		[SupportByVersionAttribute("Access", 12,14,15,16)]
		public void RunSavedImportExport(object savedImportExportName)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(savedImportExportName);
			Invoker.Method(this, "RunSavedImportExport", paramsArray);
		}

		/// <summary>
		/// SupportByVersion Access 14, 15, 16
		/// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff196381.aspx
		/// </summary>
		/// <param name="objectType">NetOffice.AccessApi.Enums.AcBrowseToObjectType ObjectType</param>
		/// <param name="objectName">object ObjectName</param>
		/// <param name="pathtoSubformControl">optional object PathtoSubformControl</param>
		/// <param name="whereCondition">optional object WhereCondition</param>
		/// <param name="page">optional object Page</param>
		/// <param name="dataMode">optional NetOffice.AccessApi.Enums.AcFormOpenDataMode DataMode = 1</param>
		[SupportByVersionAttribute("Access", 14,15,16)]
		public void BrowseTo(NetOffice.AccessApi.Enums.AcBrowseToObjectType objectType, object objectName, object pathtoSubformControl, object whereCondition, object page, object dataMode)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(objectType, objectName, pathtoSubformControl, whereCondition, page, dataMode);
			Invoker.Method(this, "BrowseTo", paramsArray);
		}

		/// <summary>
		/// SupportByVersion Access 14, 15, 16
		/// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff196381.aspx
		/// </summary>
		/// <param name="objectType">NetOffice.AccessApi.Enums.AcBrowseToObjectType ObjectType</param>
		/// <param name="objectName">object ObjectName</param>
		[CustomMethodAttribute]
		[SupportByVersionAttribute("Access", 14,15,16)]
		public void BrowseTo(NetOffice.AccessApi.Enums.AcBrowseToObjectType objectType, object objectName)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(objectType, objectName);
			Invoker.Method(this, "BrowseTo", paramsArray);
		}

		/// <summary>
		/// SupportByVersion Access 14, 15, 16
		/// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff196381.aspx
		/// </summary>
		/// <param name="objectType">NetOffice.AccessApi.Enums.AcBrowseToObjectType ObjectType</param>
		/// <param name="objectName">object ObjectName</param>
		/// <param name="pathtoSubformControl">optional object PathtoSubformControl</param>
		[CustomMethodAttribute]
		[SupportByVersionAttribute("Access", 14,15,16)]
		public void BrowseTo(NetOffice.AccessApi.Enums.AcBrowseToObjectType objectType, object objectName, object pathtoSubformControl)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(objectType, objectName, pathtoSubformControl);
			Invoker.Method(this, "BrowseTo", paramsArray);
		}

		/// <summary>
		/// SupportByVersion Access 14, 15, 16
		/// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff196381.aspx
		/// </summary>
		/// <param name="objectType">NetOffice.AccessApi.Enums.AcBrowseToObjectType ObjectType</param>
		/// <param name="objectName">object ObjectName</param>
		/// <param name="pathtoSubformControl">optional object PathtoSubformControl</param>
		/// <param name="whereCondition">optional object WhereCondition</param>
		[CustomMethodAttribute]
		[SupportByVersionAttribute("Access", 14,15,16)]
		public void BrowseTo(NetOffice.AccessApi.Enums.AcBrowseToObjectType objectType, object objectName, object pathtoSubformControl, object whereCondition)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(objectType, objectName, pathtoSubformControl, whereCondition);
			Invoker.Method(this, "BrowseTo", paramsArray);
		}

		/// <summary>
		/// SupportByVersion Access 14, 15, 16
		/// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff196381.aspx
		/// </summary>
		/// <param name="objectType">NetOffice.AccessApi.Enums.AcBrowseToObjectType ObjectType</param>
		/// <param name="objectName">object ObjectName</param>
		/// <param name="pathtoSubformControl">optional object PathtoSubformControl</param>
		/// <param name="whereCondition">optional object WhereCondition</param>
		/// <param name="page">optional object Page</param>
		[CustomMethodAttribute]
		[SupportByVersionAttribute("Access", 14,15,16)]
		public void BrowseTo(NetOffice.AccessApi.Enums.AcBrowseToObjectType objectType, object objectName, object pathtoSubformControl, object whereCondition, object page)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(objectType, objectName, pathtoSubformControl, whereCondition, page);
			Invoker.Method(this, "BrowseTo", paramsArray);
		}

		/// <summary>
		/// SupportByVersion Access 14, 15, 16
		/// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff194182.aspx
		/// </summary>
		/// <param name="name">object Name</param>
		/// <param name="expression">object Expression</param>
		[SupportByVersionAttribute("Access", 14,15,16)]
		public void SetParameter(object name, object expression)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(name, expression);
			Invoker.Method(this, "SetParameter", paramsArray);
		}

		/// <summary>
		/// SupportByVersion Access 14, 15, 16
		/// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff836068.aspx
		/// </summary>
		/// <param name="macroName">object MacroName</param>
		[SupportByVersionAttribute("Access", 14,15,16)]
		public void RunDataMacro(object macroName)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(macroName);
			Invoker.Method(this, "RunDataMacro", paramsArray);
		}

		/// <summary>
		/// SupportByVersion Access 14, 15, 16
		/// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff844761.aspx
		/// </summary>
		/// <param name="orderBy">object OrderBy</param>
		/// <param name="controlName">optional object ControlName</param>
		[SupportByVersionAttribute("Access", 14,15,16)]
		public void SetOrderBy(object orderBy, object controlName)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(orderBy, controlName);
			Invoker.Method(this, "SetOrderBy", paramsArray);
		}

		/// <summary>
		/// SupportByVersion Access 14, 15, 16
		/// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff844761.aspx
		/// </summary>
		/// <param name="orderBy">object OrderBy</param>
		[CustomMethodAttribute]
		[SupportByVersionAttribute("Access", 14,15,16)]
		public void SetOrderBy(object orderBy)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(orderBy);
			Invoker.Method(this, "SetOrderBy", paramsArray);
		}

		/// <summary>
		/// SupportByVersion Access 14, 15, 16
		/// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff197950.aspx
		/// </summary>
		/// <param name="filterName">optional object FilterName</param>
		/// <param name="whereCondition">optional object WhereCondition</param>
		/// <param name="controlName">optional object ControlName</param>
		[SupportByVersionAttribute("Access", 14,15,16)]
		public void SetFilter(object filterName, object whereCondition, object controlName)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(filterName, whereCondition, controlName);
			Invoker.Method(this, "SetFilter", paramsArray);
		}

		/// <summary>
		/// SupportByVersion Access 14, 15, 16
		/// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff197950.aspx
		/// </summary>
		[CustomMethodAttribute]
		[SupportByVersionAttribute("Access", 14,15,16)]
		public void SetFilter()
		{
			object[] paramsArray = null;
			Invoker.Method(this, "SetFilter", paramsArray);
		}

		/// <summary>
		/// SupportByVersion Access 14, 15, 16
		/// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff197950.aspx
		/// </summary>
		/// <param name="filterName">optional object FilterName</param>
		[CustomMethodAttribute]
		[SupportByVersionAttribute("Access", 14,15,16)]
		public void SetFilter(object filterName)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(filterName);
			Invoker.Method(this, "SetFilter", paramsArray);
		}

		/// <summary>
		/// SupportByVersion Access 14, 15, 16
		/// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff197950.aspx
		/// </summary>
		/// <param name="filterName">optional object FilterName</param>
		/// <param name="whereCondition">optional object WhereCondition</param>
		[CustomMethodAttribute]
		[SupportByVersionAttribute("Access", 14,15,16)]
		public void SetFilter(object filterName, object whereCondition)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(filterName, whereCondition);
			Invoker.Method(this, "SetFilter", paramsArray);
		}

		/// <summary>
		/// SupportByVersion Access 14, 15, 16
		/// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff191907.aspx
		/// </summary>
		[SupportByVersionAttribute("Access", 14,15,16)]
		public void RefreshRecord()
		{
			object[] paramsArray = null;
			Invoker.Method(this, "RefreshRecord", paramsArray);
		}

		#endregion
		#pragma warning restore
	}
}