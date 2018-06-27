using System;
using NetRuntimeSystem = System;
using System.ComponentModel;
using NetOffice.Attributes;
using NetOffice.AccessApi;

namespace NetOffice.AccessApi.Behind
{
	/// <summary>
	/// DispatchInterface DoCmd 
	/// SupportByVersion Access, 9,10,11,12,14,15,16
	/// </summary>
	/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff192694.aspx </remarks>
	[SupportByVersion("Access", 9,10,11,12,14,15,16)]
	[EntityType(EntityType.IsDispatchInterface)]
 	public class DoCmd : COMObject, NetOffice.AccessApi.DoCmd
	{
		#pragma warning disable

		#region Type Information

        /// <summary>
        /// Contract Type
        /// </summary>
        [EditorBrowsable(EditorBrowsableState.Advanced), Browsable(false), Category("NetOffice"), CoreOverridden]
        public override Type ContractType
        {
            get
            {
                if(null == _contractType)
                    _contractType = typeof(NetOffice.AccessApi.DoCmd);
                return _contractType;
            }
        }
        private static Type _contractType;


		/// <summary>
		/// Instance Type
		/// </summary>
		[EditorBrowsable(EditorBrowsableState.Advanced), Browsable(false), Category("NetOffice"), CoreOverridden]
		public override Type InstanceType
		{
			get
			{
				return LateBindingApiWrapperType;
			}
		}

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
        
		#region Ctor

		/// <summary>
		/// Stub Ctor, not indented to use
		/// </summary>
		public DoCmd() : base()
		{

		}

		#endregion
		
		#region Properties

		#endregion

		#region Methods

		/// <summary>
		/// SupportByVersion Access 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff834781.aspx </remarks>
		/// <param name="menuName">object menuName</param>
		/// <param name="menuMacroName">object menuMacroName</param>
		/// <param name="statusBarText">object statusBarText</param>
		[SupportByVersion("Access", 9,10,11,12,14,15,16)]
		public virtual void AddMenu(object menuName, object menuMacroName, object statusBarText)
		{
			 InvokerService.InvokeInternal.ExecuteMethod(this, "AddMenu", menuName, menuMacroName, statusBarText);
		}

		/// <summary>
		/// SupportByVersion Access 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff197651.aspx </remarks>
		/// <param name="filterName">optional object filterName</param>
		/// <param name="whereCondition">optional object whereCondition</param>
		[SupportByVersion("Access", 9,10,11,12,14,15,16)]
		public virtual void ApplyFilter(object filterName, object whereCondition)
		{
			 InvokerService.InvokeInternal.ExecuteMethod(this, "ApplyFilter", filterName, whereCondition);
		}

		/// <summary>
		/// SupportByVersion Access 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff197651.aspx </remarks>
		/// <param name="filterName">optional object filterName</param>
		/// <param name="whereCondition">optional object whereCondition</param>
		/// <param name="controlName">optional object controlName</param>
		[SupportByVersion("Access", 12,14,15,16)]
		public virtual void ApplyFilter(object filterName, object whereCondition, object controlName)
		{
			 InvokerService.InvokeInternal.ExecuteMethod(this, "ApplyFilter", filterName, whereCondition, controlName);
		}

		/// <summary>
		/// SupportByVersion Access 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff197651.aspx </remarks>
		[CustomMethod]
		[SupportByVersion("Access", 9,10,11,12,14,15,16)]
		public virtual void ApplyFilter()
		{
			 InvokerService.InvokeInternal.ExecuteMethod(this, "ApplyFilter");
		}

		/// <summary>
		/// SupportByVersion Access 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff197651.aspx </remarks>
		/// <param name="filterName">optional object filterName</param>
		[CustomMethod]
		[SupportByVersion("Access", 9,10,11,12,14,15,16)]
		public virtual void ApplyFilter(object filterName)
		{
			 InvokerService.InvokeInternal.ExecuteMethod(this, "ApplyFilter", filterName);
		}

		/// <summary>
		/// SupportByVersion Access 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff196680.aspx </remarks>
		[SupportByVersion("Access", 9,10,11,12,14,15,16)]
		public virtual void Beep()
		{
			 InvokerService.InvokeInternal.ExecuteMethod(this, "Beep");
		}

		/// <summary>
		/// SupportByVersion Access 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff836964.aspx </remarks>
		[SupportByVersion("Access", 9,10,11,12,14,15,16)]
		public virtual void CancelEvent()
		{
			 InvokerService.InvokeInternal.ExecuteMethod(this, "CancelEvent");
		}

		/// <summary>
		/// SupportByVersion Access 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff192860.aspx </remarks>
		/// <param name="objectType">optional NetOffice.AccessApi.Enums.AcObjectType ObjectType = -1</param>
		/// <param name="objectName">optional object objectName</param>
		/// <param name="save">optional NetOffice.AccessApi.Enums.AcCloseSave Save = 0</param>
		[SupportByVersion("Access", 9,10,11,12,14,15,16)]
		public virtual void Close(object objectType, object objectName, object save)
		{
			 InvokerService.InvokeInternal.ExecuteMethod(this, "Close", objectType, objectName, save);
		}

		/// <summary>
		/// SupportByVersion Access 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff192860.aspx </remarks>
		[CustomMethod]
		[SupportByVersion("Access", 9,10,11,12,14,15,16)]
		public virtual void Close()
		{
			 InvokerService.InvokeInternal.ExecuteMethod(this, "Close");
		}

		/// <summary>
		/// SupportByVersion Access 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff192860.aspx </remarks>
		/// <param name="objectType">optional NetOffice.AccessApi.Enums.AcObjectType ObjectType = -1</param>
		[CustomMethod]
		[SupportByVersion("Access", 9,10,11,12,14,15,16)]
		public virtual void Close(object objectType)
		{
			 InvokerService.InvokeInternal.ExecuteMethod(this, "Close", objectType);
		}

		/// <summary>
		/// SupportByVersion Access 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff192860.aspx </remarks>
		/// <param name="objectType">optional NetOffice.AccessApi.Enums.AcObjectType ObjectType = -1</param>
		/// <param name="objectName">optional object objectName</param>
		[CustomMethod]
		[SupportByVersion("Access", 9,10,11,12,14,15,16)]
		public virtual void Close(object objectType, object objectName)
		{
			 InvokerService.InvokeInternal.ExecuteMethod(this, "Close", objectType, objectName);
		}

		/// <summary>
		/// SupportByVersion Access 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff844724.aspx </remarks>
		/// <param name="destinationDatabase">optional object destinationDatabase</param>
		/// <param name="newName">optional object newName</param>
		/// <param name="sourceObjectType">optional NetOffice.AccessApi.Enums.AcObjectType SourceObjectType = -1</param>
		/// <param name="sourceObjectName">optional object sourceObjectName</param>
		[SupportByVersion("Access", 9,10,11,12,14,15,16)]
		public virtual void CopyObject(object destinationDatabase, object newName, object sourceObjectType, object sourceObjectName)
		{
			 InvokerService.InvokeInternal.ExecuteMethod(this, "CopyObject", destinationDatabase, newName, sourceObjectType, sourceObjectName);
		}

		/// <summary>
		/// SupportByVersion Access 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff844724.aspx </remarks>
		[CustomMethod]
		[SupportByVersion("Access", 9,10,11,12,14,15,16)]
		public virtual void CopyObject()
		{
			 InvokerService.InvokeInternal.ExecuteMethod(this, "CopyObject");
		}

		/// <summary>
		/// SupportByVersion Access 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff844724.aspx </remarks>
		/// <param name="destinationDatabase">optional object destinationDatabase</param>
		[CustomMethod]
		[SupportByVersion("Access", 9,10,11,12,14,15,16)]
		public virtual void CopyObject(object destinationDatabase)
		{
			 InvokerService.InvokeInternal.ExecuteMethod(this, "CopyObject", destinationDatabase);
		}

		/// <summary>
		/// SupportByVersion Access 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff844724.aspx </remarks>
		/// <param name="destinationDatabase">optional object destinationDatabase</param>
		/// <param name="newName">optional object newName</param>
		[CustomMethod]
		[SupportByVersion("Access", 9,10,11,12,14,15,16)]
		public virtual void CopyObject(object destinationDatabase, object newName)
		{
			 InvokerService.InvokeInternal.ExecuteMethod(this, "CopyObject", destinationDatabase, newName);
		}

		/// <summary>
		/// SupportByVersion Access 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff844724.aspx </remarks>
		/// <param name="destinationDatabase">optional object destinationDatabase</param>
		/// <param name="newName">optional object newName</param>
		/// <param name="sourceObjectType">optional NetOffice.AccessApi.Enums.AcObjectType SourceObjectType = -1</param>
		[CustomMethod]
		[SupportByVersion("Access", 9,10,11,12,14,15,16)]
		public virtual void CopyObject(object destinationDatabase, object newName, object sourceObjectType)
		{
			 InvokerService.InvokeInternal.ExecuteMethod(this, "CopyObject", destinationDatabase, newName, sourceObjectType);
		}

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
		public virtual void DoMenuItem(object menuBar, object menuName, object command, object subcommand, object version)
		{
			 InvokerService.InvokeInternal.ExecuteMethod(this, "DoMenuItem", new object[]{ menuBar, menuName, command, subcommand, version });
		}

		/// <summary>
		/// SupportByVersion Access 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff822447.aspx </remarks>
		/// <param name="menuBar">object menuBar</param>
		/// <param name="menuName">object menuName</param>
		/// <param name="command">object command</param>
		[CustomMethod]
		[SupportByVersion("Access", 9,10,11,12,14,15,16)]
		public virtual void DoMenuItem(object menuBar, object menuName, object command)
		{
			 InvokerService.InvokeInternal.ExecuteMethod(this, "DoMenuItem", menuBar, menuName, command);
		}

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
		public virtual void DoMenuItem(object menuBar, object menuName, object command, object subcommand)
		{
			 InvokerService.InvokeInternal.ExecuteMethod(this, "DoMenuItem", menuBar, menuName, command, subcommand);
		}

		/// <summary>
		/// SupportByVersion Access 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff193863.aspx </remarks>
		/// <param name="echoOn">object echoOn</param>
		/// <param name="statusBarText">optional object statusBarText</param>
		[SupportByVersion("Access", 9,10,11,12,14,15,16)]
		public virtual void Echo(object echoOn, object statusBarText)
		{
			 InvokerService.InvokeInternal.ExecuteMethod(this, "Echo", echoOn, statusBarText);
		}

		/// <summary>
		/// SupportByVersion Access 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff193863.aspx </remarks>
		/// <param name="echoOn">object echoOn</param>
		[CustomMethod]
		[SupportByVersion("Access", 9,10,11,12,14,15,16)]
		public virtual void Echo(object echoOn)
		{
			 InvokerService.InvokeInternal.ExecuteMethod(this, "Echo", echoOn);
		}

		/// <summary>
		/// SupportByVersion Access 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff196453.aspx </remarks>
		[SupportByVersion("Access", 9,10,11,12,14,15,16)]
		public virtual void FindNext()
		{
			 InvokerService.InvokeInternal.ExecuteMethod(this, "FindNext");
		}

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
		public virtual void FindRecord(object findWhat, object match, object matchCase, object search, object searchAsFormatted, object onlyCurrentField, object findFirst)
		{
			 InvokerService.InvokeInternal.ExecuteMethod(this, "FindRecord", new object[]{ findWhat, match, matchCase, search, searchAsFormatted, onlyCurrentField, findFirst });
		}

		/// <summary>
		/// SupportByVersion Access 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff835361.aspx </remarks>
		/// <param name="findWhat">object findWhat</param>
		[CustomMethod]
		[SupportByVersion("Access", 9,10,11,12,14,15,16)]
		public virtual void FindRecord(object findWhat)
		{
			 InvokerService.InvokeInternal.ExecuteMethod(this, "FindRecord", findWhat);
		}

		/// <summary>
		/// SupportByVersion Access 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff835361.aspx </remarks>
		/// <param name="findWhat">object findWhat</param>
		/// <param name="match">optional NetOffice.AccessApi.Enums.AcFindMatch Match = 1</param>
		[CustomMethod]
		[SupportByVersion("Access", 9,10,11,12,14,15,16)]
		public virtual void FindRecord(object findWhat, object match)
		{
			 InvokerService.InvokeInternal.ExecuteMethod(this, "FindRecord", findWhat, match);
		}

		/// <summary>
		/// SupportByVersion Access 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff835361.aspx </remarks>
		/// <param name="findWhat">object findWhat</param>
		/// <param name="match">optional NetOffice.AccessApi.Enums.AcFindMatch Match = 1</param>
		/// <param name="matchCase">optional object matchCase</param>
		[CustomMethod]
		[SupportByVersion("Access", 9,10,11,12,14,15,16)]
		public virtual void FindRecord(object findWhat, object match, object matchCase)
		{
			 InvokerService.InvokeInternal.ExecuteMethod(this, "FindRecord", findWhat, match, matchCase);
		}

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
		public virtual void FindRecord(object findWhat, object match, object matchCase, object search)
		{
			 InvokerService.InvokeInternal.ExecuteMethod(this, "FindRecord", findWhat, match, matchCase, search);
		}

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
		public virtual void FindRecord(object findWhat, object match, object matchCase, object search, object searchAsFormatted)
		{
			 InvokerService.InvokeInternal.ExecuteMethod(this, "FindRecord", new object[]{ findWhat, match, matchCase, search, searchAsFormatted });
		}

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
		public virtual void FindRecord(object findWhat, object match, object matchCase, object search, object searchAsFormatted, object onlyCurrentField)
		{
			 InvokerService.InvokeInternal.ExecuteMethod(this, "FindRecord", new object[]{ findWhat, match, matchCase, search, searchAsFormatted, onlyCurrentField });
		}

		/// <summary>
		/// SupportByVersion Access 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff192079.aspx </remarks>
		/// <param name="controlName">object controlName</param>
		[SupportByVersion("Access", 9,10,11,12,14,15,16)]
		public virtual void GoToControl(object controlName)
		{
			 InvokerService.InvokeInternal.ExecuteMethod(this, "GoToControl", controlName);
		}

		/// <summary>
		/// SupportByVersion Access 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff192504.aspx </remarks>
		/// <param name="pageNumber">object pageNumber</param>
		/// <param name="right">optional object right</param>
		/// <param name="down">optional object down</param>
		[SupportByVersion("Access", 9,10,11,12,14,15,16)]
		public virtual void GoToPage(object pageNumber, object right, object down)
		{
			 InvokerService.InvokeInternal.ExecuteMethod(this, "GoToPage", pageNumber, right, down);
		}

		/// <summary>
		/// SupportByVersion Access 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff192504.aspx </remarks>
		/// <param name="pageNumber">object pageNumber</param>
		[CustomMethod]
		[SupportByVersion("Access", 9,10,11,12,14,15,16)]
		public virtual void GoToPage(object pageNumber)
		{
			 InvokerService.InvokeInternal.ExecuteMethod(this, "GoToPage", pageNumber);
		}

		/// <summary>
		/// SupportByVersion Access 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff192504.aspx </remarks>
		/// <param name="pageNumber">object pageNumber</param>
		/// <param name="right">optional object right</param>
		[CustomMethod]
		[SupportByVersion("Access", 9,10,11,12,14,15,16)]
		public virtual void GoToPage(object pageNumber, object right)
		{
			 InvokerService.InvokeInternal.ExecuteMethod(this, "GoToPage", pageNumber, right);
		}

		/// <summary>
		/// SupportByVersion Access 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff194117.aspx </remarks>
		/// <param name="objectType">optional NetOffice.AccessApi.Enums.AcDataObjectType ObjectType = -1</param>
		/// <param name="objectName">optional object objectName</param>
		/// <param name="record">optional NetOffice.AccessApi.Enums.AcRecord Record = 1</param>
		/// <param name="offset">optional object offset</param>
		[SupportByVersion("Access", 9,10,11,12,14,15,16)]
		public virtual void GoToRecord(object objectType, object objectName, object record, object offset)
		{
			 InvokerService.InvokeInternal.ExecuteMethod(this, "GoToRecord", objectType, objectName, record, offset);
		}

		/// <summary>
		/// SupportByVersion Access 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff194117.aspx </remarks>
		[CustomMethod]
		[SupportByVersion("Access", 9,10,11,12,14,15,16)]
		public virtual void GoToRecord()
		{
			 InvokerService.InvokeInternal.ExecuteMethod(this, "GoToRecord");
		}

		/// <summary>
		/// SupportByVersion Access 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff194117.aspx </remarks>
		/// <param name="objectType">optional NetOffice.AccessApi.Enums.AcDataObjectType ObjectType = -1</param>
		[CustomMethod]
		[SupportByVersion("Access", 9,10,11,12,14,15,16)]
		public virtual void GoToRecord(object objectType)
		{
			 InvokerService.InvokeInternal.ExecuteMethod(this, "GoToRecord", objectType);
		}

		/// <summary>
		/// SupportByVersion Access 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff194117.aspx </remarks>
		/// <param name="objectType">optional NetOffice.AccessApi.Enums.AcDataObjectType ObjectType = -1</param>
		/// <param name="objectName">optional object objectName</param>
		[CustomMethod]
		[SupportByVersion("Access", 9,10,11,12,14,15,16)]
		public virtual void GoToRecord(object objectType, object objectName)
		{
			 InvokerService.InvokeInternal.ExecuteMethod(this, "GoToRecord", objectType, objectName);
		}

		/// <summary>
		/// SupportByVersion Access 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff194117.aspx </remarks>
		/// <param name="objectType">optional NetOffice.AccessApi.Enums.AcDataObjectType ObjectType = -1</param>
		/// <param name="objectName">optional object objectName</param>
		/// <param name="record">optional NetOffice.AccessApi.Enums.AcRecord Record = 1</param>
		[CustomMethod]
		[SupportByVersion("Access", 9,10,11,12,14,15,16)]
		public virtual void GoToRecord(object objectType, object objectName, object record)
		{
			 InvokerService.InvokeInternal.ExecuteMethod(this, "GoToRecord", objectType, objectName, record);
		}

		/// <summary>
		/// SupportByVersion Access 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff835648.aspx </remarks>
		/// <param name="hourglassOn">object hourglassOn</param>
		[SupportByVersion("Access", 9,10,11,12,14,15,16)]
		public virtual void Hourglass(object hourglassOn)
		{
			 InvokerService.InvokeInternal.ExecuteMethod(this, "Hourglass", hourglassOn);
		}

		/// <summary>
		/// SupportByVersion Access 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff195449.aspx </remarks>
		[SupportByVersion("Access", 9,10,11,12,14,15,16)]
		public virtual void Maximize()
		{
			 InvokerService.InvokeInternal.ExecuteMethod(this, "Maximize");
		}

		/// <summary>
		/// SupportByVersion Access 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff837032.aspx </remarks>
		[SupportByVersion("Access", 9,10,11,12,14,15,16)]
		public virtual void Minimize()
		{
			 InvokerService.InvokeInternal.ExecuteMethod(this, "Minimize");
		}

		/// <summary>
		/// SupportByVersion Access 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff197394.aspx </remarks>
		/// <param name="right">optional object right</param>
		/// <param name="down">optional object down</param>
		/// <param name="width">optional object width</param>
		/// <param name="height">optional object height</param>
		[SupportByVersion("Access", 9,10,11,12,14,15,16)]
		public virtual void MoveSize(object right, object down, object width, object height)
		{
			 InvokerService.InvokeInternal.ExecuteMethod(this, "MoveSize", right, down, width, height);
		}

		/// <summary>
		/// SupportByVersion Access 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff197394.aspx </remarks>
		[CustomMethod]
		[SupportByVersion("Access", 9,10,11,12,14,15,16)]
		public virtual void MoveSize()
		{
			 InvokerService.InvokeInternal.ExecuteMethod(this, "MoveSize");
		}

		/// <summary>
		/// SupportByVersion Access 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff197394.aspx </remarks>
		/// <param name="right">optional object right</param>
		[CustomMethod]
		[SupportByVersion("Access", 9,10,11,12,14,15,16)]
		public virtual void MoveSize(object right)
		{
			 InvokerService.InvokeInternal.ExecuteMethod(this, "MoveSize", right);
		}

		/// <summary>
		/// SupportByVersion Access 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff197394.aspx </remarks>
		/// <param name="right">optional object right</param>
		/// <param name="down">optional object down</param>
		[CustomMethod]
		[SupportByVersion("Access", 9,10,11,12,14,15,16)]
		public virtual void MoveSize(object right, object down)
		{
			 InvokerService.InvokeInternal.ExecuteMethod(this, "MoveSize", right, down);
		}

		/// <summary>
		/// SupportByVersion Access 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff197394.aspx </remarks>
		/// <param name="right">optional object right</param>
		/// <param name="down">optional object down</param>
		/// <param name="width">optional object width</param>
		[CustomMethod]
		[SupportByVersion("Access", 9,10,11,12,14,15,16)]
		public virtual void MoveSize(object right, object down, object width)
		{
			 InvokerService.InvokeInternal.ExecuteMethod(this, "MoveSize", right, down, width);
		}

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
		public virtual void OpenForm(object formName, object view, object filterName, object whereCondition, object dataMode, object windowMode, object openArgs)
		{
			 InvokerService.InvokeInternal.ExecuteMethod(this, "OpenForm", new object[]{ formName, view, filterName, whereCondition, dataMode, windowMode, openArgs });
		}

		/// <summary>
		/// SupportByVersion Access 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff820845.aspx </remarks>
		/// <param name="formName">object formName</param>
		[CustomMethod]
		[SupportByVersion("Access", 9,10,11,12,14,15,16)]
		public virtual void OpenForm(object formName)
		{
			 InvokerService.InvokeInternal.ExecuteMethod(this, "OpenForm", formName);
		}

		/// <summary>
		/// SupportByVersion Access 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff820845.aspx </remarks>
		/// <param name="formName">object formName</param>
		/// <param name="view">optional NetOffice.AccessApi.Enums.AcFormView View = 0</param>
		[CustomMethod]
		[SupportByVersion("Access", 9,10,11,12,14,15,16)]
		public virtual void OpenForm(object formName, object view)
		{
			 InvokerService.InvokeInternal.ExecuteMethod(this, "OpenForm", formName, view);
		}

		/// <summary>
		/// SupportByVersion Access 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff820845.aspx </remarks>
		/// <param name="formName">object formName</param>
		/// <param name="view">optional NetOffice.AccessApi.Enums.AcFormView View = 0</param>
		/// <param name="filterName">optional object filterName</param>
		[CustomMethod]
		[SupportByVersion("Access", 9,10,11,12,14,15,16)]
		public virtual void OpenForm(object formName, object view, object filterName)
		{
			 InvokerService.InvokeInternal.ExecuteMethod(this, "OpenForm", formName, view, filterName);
		}

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
		public virtual void OpenForm(object formName, object view, object filterName, object whereCondition)
		{
			 InvokerService.InvokeInternal.ExecuteMethod(this, "OpenForm", formName, view, filterName, whereCondition);
		}

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
		public virtual void OpenForm(object formName, object view, object filterName, object whereCondition, object dataMode)
		{
			 InvokerService.InvokeInternal.ExecuteMethod(this, "OpenForm", new object[]{ formName, view, filterName, whereCondition, dataMode });
		}

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
		public virtual void OpenForm(object formName, object view, object filterName, object whereCondition, object dataMode, object windowMode)
		{
			 InvokerService.InvokeInternal.ExecuteMethod(this, "OpenForm", new object[]{ formName, view, filterName, whereCondition, dataMode, windowMode });
		}

		/// <summary>
		/// SupportByVersion Access 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff192746.aspx </remarks>
		/// <param name="queryName">object queryName</param>
		/// <param name="view">optional NetOffice.AccessApi.Enums.AcView View = 0</param>
		/// <param name="dataMode">optional NetOffice.AccessApi.Enums.AcOpenDataMode DataMode = 1</param>
		[SupportByVersion("Access", 9,10,11,12,14,15,16)]
		public virtual void OpenQuery(object queryName, object view, object dataMode)
		{
			 InvokerService.InvokeInternal.ExecuteMethod(this, "OpenQuery", queryName, view, dataMode);
		}

		/// <summary>
		/// SupportByVersion Access 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff192746.aspx </remarks>
		/// <param name="queryName">object queryName</param>
		[CustomMethod]
		[SupportByVersion("Access", 9,10,11,12,14,15,16)]
		public virtual void OpenQuery(object queryName)
		{
			 InvokerService.InvokeInternal.ExecuteMethod(this, "OpenQuery", queryName);
		}

		/// <summary>
		/// SupportByVersion Access 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff192746.aspx </remarks>
		/// <param name="queryName">object queryName</param>
		/// <param name="view">optional NetOffice.AccessApi.Enums.AcView View = 0</param>
		[CustomMethod]
		[SupportByVersion("Access", 9,10,11,12,14,15,16)]
		public virtual void OpenQuery(object queryName, object view)
		{
			 InvokerService.InvokeInternal.ExecuteMethod(this, "OpenQuery", queryName, view);
		}

		/// <summary>
		/// SupportByVersion Access 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff194975.aspx </remarks>
		/// <param name="tableName">object tableName</param>
		/// <param name="view">optional NetOffice.AccessApi.Enums.AcView View = 0</param>
		/// <param name="dataMode">optional NetOffice.AccessApi.Enums.AcOpenDataMode DataMode = 1</param>
		[SupportByVersion("Access", 9,10,11,12,14,15,16)]
		public virtual void OpenTable(object tableName, object view, object dataMode)
		{
			 InvokerService.InvokeInternal.ExecuteMethod(this, "OpenTable", tableName, view, dataMode);
		}

		/// <summary>
		/// SupportByVersion Access 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff194975.aspx </remarks>
		/// <param name="tableName">object tableName</param>
		[CustomMethod]
		[SupportByVersion("Access", 9,10,11,12,14,15,16)]
		public virtual void OpenTable(object tableName)
		{
			 InvokerService.InvokeInternal.ExecuteMethod(this, "OpenTable", tableName);
		}

		/// <summary>
		/// SupportByVersion Access 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff194975.aspx </remarks>
		/// <param name="tableName">object tableName</param>
		/// <param name="view">optional NetOffice.AccessApi.Enums.AcView View = 0</param>
		[CustomMethod]
		[SupportByVersion("Access", 9,10,11,12,14,15,16)]
		public virtual void OpenTable(object tableName, object view)
		{
			 InvokerService.InvokeInternal.ExecuteMethod(this, "OpenTable", tableName, view);
		}

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
		public virtual void PrintOut(object printRange, object pageFrom, object pageTo, object printQuality, object copies, object collateCopies)
		{
			 InvokerService.InvokeInternal.ExecuteMethod(this, "PrintOut", new object[]{ printRange, pageFrom, pageTo, printQuality, copies, collateCopies });
		}

		/// <summary>
		/// SupportByVersion Access 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff192667.aspx </remarks>
		[CustomMethod]
		[SupportByVersion("Access", 9,10,11,12,14,15,16)]
		public virtual void PrintOut()
		{
			 InvokerService.InvokeInternal.ExecuteMethod(this, "PrintOut");
		}

		/// <summary>
		/// SupportByVersion Access 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff192667.aspx </remarks>
		/// <param name="printRange">optional NetOffice.AccessApi.Enums.AcPrintRange PrintRange = 0</param>
		[CustomMethod]
		[SupportByVersion("Access", 9,10,11,12,14,15,16)]
		public virtual void PrintOut(object printRange)
		{
			 InvokerService.InvokeInternal.ExecuteMethod(this, "PrintOut", printRange);
		}

		/// <summary>
		/// SupportByVersion Access 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff192667.aspx </remarks>
		/// <param name="printRange">optional NetOffice.AccessApi.Enums.AcPrintRange PrintRange = 0</param>
		/// <param name="pageFrom">optional object pageFrom</param>
		[CustomMethod]
		[SupportByVersion("Access", 9,10,11,12,14,15,16)]
		public virtual void PrintOut(object printRange, object pageFrom)
		{
			 InvokerService.InvokeInternal.ExecuteMethod(this, "PrintOut", printRange, pageFrom);
		}

		/// <summary>
		/// SupportByVersion Access 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff192667.aspx </remarks>
		/// <param name="printRange">optional NetOffice.AccessApi.Enums.AcPrintRange PrintRange = 0</param>
		/// <param name="pageFrom">optional object pageFrom</param>
		/// <param name="pageTo">optional object pageTo</param>
		[CustomMethod]
		[SupportByVersion("Access", 9,10,11,12,14,15,16)]
		public virtual void PrintOut(object printRange, object pageFrom, object pageTo)
		{
			 InvokerService.InvokeInternal.ExecuteMethod(this, "PrintOut", printRange, pageFrom, pageTo);
		}

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
		public virtual void PrintOut(object printRange, object pageFrom, object pageTo, object printQuality)
		{
			 InvokerService.InvokeInternal.ExecuteMethod(this, "PrintOut", printRange, pageFrom, pageTo, printQuality);
		}

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
		public virtual void PrintOut(object printRange, object pageFrom, object pageTo, object printQuality, object copies)
		{
			 InvokerService.InvokeInternal.ExecuteMethod(this, "PrintOut", new object[]{ printRange, pageFrom, pageTo, printQuality, copies });
		}

		/// <summary>
		/// SupportByVersion Access 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff191887.aspx </remarks>
		/// <param name="options">optional NetOffice.AccessApi.Enums.AcQuitOption Options = 1</param>
		[SupportByVersion("Access", 9,10,11,12,14,15,16)]
		public virtual void Quit(object options)
		{
			 InvokerService.InvokeInternal.ExecuteMethod(this, "Quit", options);
		}

		/// <summary>
		/// SupportByVersion Access 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff191887.aspx </remarks>
		[CustomMethod]
		[SupportByVersion("Access", 9,10,11,12,14,15,16)]
		public virtual void Quit()
		{
			 InvokerService.InvokeInternal.ExecuteMethod(this, "Quit");
		}

		/// <summary>
		/// SupportByVersion Access 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff195253.aspx </remarks>
		/// <param name="controlName">optional object controlName</param>
		[SupportByVersion("Access", 9,10,11,12,14,15,16)]
		public virtual void Requery(object controlName)
		{
			 InvokerService.InvokeInternal.ExecuteMethod(this, "Requery", controlName);
		}

		/// <summary>
		/// SupportByVersion Access 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff195253.aspx </remarks>
		[CustomMethod]
		[SupportByVersion("Access", 9,10,11,12,14,15,16)]
		public virtual void Requery()
		{
			 InvokerService.InvokeInternal.ExecuteMethod(this, "Requery");
		}

		/// <summary>
		/// SupportByVersion Access 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff195560.aspx </remarks>
		/// <param name="objectType">optional NetOffice.AccessApi.Enums.AcObjectType ObjectType = -1</param>
		/// <param name="objectName">optional object objectName</param>
		[SupportByVersion("Access", 9,10,11,12,14,15,16)]
		public virtual void RepaintObject(object objectType, object objectName)
		{
			 InvokerService.InvokeInternal.ExecuteMethod(this, "RepaintObject", objectType, objectName);
		}

		/// <summary>
		/// SupportByVersion Access 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff195560.aspx </remarks>
		[CustomMethod]
		[SupportByVersion("Access", 9,10,11,12,14,15,16)]
		public virtual void RepaintObject()
		{
			 InvokerService.InvokeInternal.ExecuteMethod(this, "RepaintObject");
		}

		/// <summary>
		/// SupportByVersion Access 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff195560.aspx </remarks>
		/// <param name="objectType">optional NetOffice.AccessApi.Enums.AcObjectType ObjectType = -1</param>
		[CustomMethod]
		[SupportByVersion("Access", 9,10,11,12,14,15,16)]
		public virtual void RepaintObject(object objectType)
		{
			 InvokerService.InvokeInternal.ExecuteMethod(this, "RepaintObject", objectType);
		}

		/// <summary>
		/// SupportByVersion Access 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff823209.aspx </remarks>
		/// <param name="newName">object newName</param>
		/// <param name="objectType">optional NetOffice.AccessApi.Enums.AcObjectType ObjectType = -1</param>
		/// <param name="oldName">optional object oldName</param>
		[SupportByVersion("Access", 9,10,11,12,14,15,16)]
		public virtual void Rename(object newName, object objectType, object oldName)
		{
			 InvokerService.InvokeInternal.ExecuteMethod(this, "Rename", newName, objectType, oldName);
		}

		/// <summary>
		/// SupportByVersion Access 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff823209.aspx </remarks>
		/// <param name="newName">object newName</param>
		[CustomMethod]
		[SupportByVersion("Access", 9,10,11,12,14,15,16)]
		public virtual void Rename(object newName)
		{
			 InvokerService.InvokeInternal.ExecuteMethod(this, "Rename", newName);
		}

		/// <summary>
		/// SupportByVersion Access 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff823209.aspx </remarks>
		/// <param name="newName">object newName</param>
		/// <param name="objectType">optional NetOffice.AccessApi.Enums.AcObjectType ObjectType = -1</param>
		[CustomMethod]
		[SupportByVersion("Access", 9,10,11,12,14,15,16)]
		public virtual void Rename(object newName, object objectType)
		{
			 InvokerService.InvokeInternal.ExecuteMethod(this, "Rename", newName, objectType);
		}

		/// <summary>
		/// SupportByVersion Access 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff193174.aspx </remarks>
		[SupportByVersion("Access", 9,10,11,12,14,15,16)]
		public virtual void Restore()
		{
			 InvokerService.InvokeInternal.ExecuteMethod(this, "Restore");
		}

		/// <summary>
		/// SupportByVersion Access 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff192075.aspx </remarks>
		/// <param name="macroName">object macroName</param>
		/// <param name="repeatCount">optional object repeatCount</param>
		/// <param name="repeatExpression">optional object repeatExpression</param>
		[SupportByVersion("Access", 9,10,11,12,14,15,16)]
		public virtual void RunMacro(object macroName, object repeatCount, object repeatExpression)
		{
			 InvokerService.InvokeInternal.ExecuteMethod(this, "RunMacro", macroName, repeatCount, repeatExpression);
		}

		/// <summary>
		/// SupportByVersion Access 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff192075.aspx </remarks>
		/// <param name="macroName">object macroName</param>
		[CustomMethod]
		[SupportByVersion("Access", 9,10,11,12,14,15,16)]
		public virtual void RunMacro(object macroName)
		{
			 InvokerService.InvokeInternal.ExecuteMethod(this, "RunMacro", macroName);
		}

		/// <summary>
		/// SupportByVersion Access 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff192075.aspx </remarks>
		/// <param name="macroName">object macroName</param>
		/// <param name="repeatCount">optional object repeatCount</param>
		[CustomMethod]
		[SupportByVersion("Access", 9,10,11,12,14,15,16)]
		public virtual void RunMacro(object macroName, object repeatCount)
		{
			 InvokerService.InvokeInternal.ExecuteMethod(this, "RunMacro", macroName, repeatCount);
		}

		/// <summary>
		/// SupportByVersion Access 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff194626.aspx </remarks>
		/// <param name="sQLStatement">object sQLStatement</param>
		/// <param name="useTransaction">optional object useTransaction</param>
		[SupportByVersion("Access", 9,10,11,12,14,15,16)]
		public virtual void RunSQL(object sQLStatement, object useTransaction)
		{
			 InvokerService.InvokeInternal.ExecuteMethod(this, "RunSQL", sQLStatement, useTransaction);
		}

		/// <summary>
		/// SupportByVersion Access 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff194626.aspx </remarks>
		/// <param name="sQLStatement">object sQLStatement</param>
		[CustomMethod]
		[SupportByVersion("Access", 9,10,11,12,14,15,16)]
		public virtual void RunSQL(object sQLStatement)
		{
			 InvokerService.InvokeInternal.ExecuteMethod(this, "RunSQL", sQLStatement);
		}

		/// <summary>
		/// SupportByVersion Access 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff835629.aspx </remarks>
		/// <param name="objectType">NetOffice.AccessApi.Enums.AcObjectType objectType</param>
		/// <param name="objectName">optional object objectName</param>
		/// <param name="inDatabaseWindow">optional object inDatabaseWindow</param>
		[SupportByVersion("Access", 9,10,11,12,14,15,16)]
		public virtual void SelectObject(NetOffice.AccessApi.Enums.AcObjectType objectType, object objectName, object inDatabaseWindow)
		{
			 InvokerService.InvokeInternal.ExecuteMethod(this, "SelectObject", objectType, objectName, inDatabaseWindow);
		}

		/// <summary>
		/// SupportByVersion Access 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff835629.aspx </remarks>
		/// <param name="objectType">NetOffice.AccessApi.Enums.AcObjectType objectType</param>
		[CustomMethod]
		[SupportByVersion("Access", 9,10,11,12,14,15,16)]
		public virtual void SelectObject(NetOffice.AccessApi.Enums.AcObjectType objectType)
		{
			 InvokerService.InvokeInternal.ExecuteMethod(this, "SelectObject", objectType);
		}

		/// <summary>
		/// SupportByVersion Access 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff835629.aspx </remarks>
		/// <param name="objectType">NetOffice.AccessApi.Enums.AcObjectType objectType</param>
		/// <param name="objectName">optional object objectName</param>
		[CustomMethod]
		[SupportByVersion("Access", 9,10,11,12,14,15,16)]
		public virtual void SelectObject(NetOffice.AccessApi.Enums.AcObjectType objectType, object objectName)
		{
			 InvokerService.InvokeInternal.ExecuteMethod(this, "SelectObject", objectType, objectName);
		}

		/// <summary>
		/// SupportByVersion Access 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff837275.aspx </remarks>
		/// <param name="warningsOn">object warningsOn</param>
		[SupportByVersion("Access", 9,10,11,12,14,15,16)]
		public virtual void SetWarnings(object warningsOn)
		{
			 InvokerService.InvokeInternal.ExecuteMethod(this, "SetWarnings", warningsOn);
		}

		/// <summary>
		/// SupportByVersion Access 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff195994.aspx </remarks>
		[SupportByVersion("Access", 9,10,11,12,14,15,16)]
		public virtual void ShowAllRecords()
		{
			 InvokerService.InvokeInternal.ExecuteMethod(this, "ShowAllRecords");
		}

		/// <summary>
		/// SupportByVersion Access 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff192676.aspx </remarks>
		/// <param name="reportName">object reportName</param>
		/// <param name="view">optional NetOffice.AccessApi.Enums.AcView View = 0</param>
		/// <param name="filterName">optional object filterName</param>
		/// <param name="whereCondition">optional object whereCondition</param>
		[SupportByVersion("Access", 9,10,11,12,14,15,16)]
		public virtual void OpenReport(object reportName, object view, object filterName, object whereCondition)
		{
			 InvokerService.InvokeInternal.ExecuteMethod(this, "OpenReport", reportName, view, filterName, whereCondition);
		}

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
		public virtual void OpenReport(object reportName, object view, object filterName, object whereCondition, object windowMode, object openArgs)
		{
			 InvokerService.InvokeInternal.ExecuteMethod(this, "OpenReport", new object[]{ reportName, view, filterName, whereCondition, windowMode, openArgs });
		}

		/// <summary>
		/// SupportByVersion Access 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff192676.aspx </remarks>
		/// <param name="reportName">object reportName</param>
		[CustomMethod]
		[SupportByVersion("Access", 9,10,11,12,14,15,16)]
		public virtual void OpenReport(object reportName)
		{
			 InvokerService.InvokeInternal.ExecuteMethod(this, "OpenReport", reportName);
		}

		/// <summary>
		/// SupportByVersion Access 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff192676.aspx </remarks>
		/// <param name="reportName">object reportName</param>
		/// <param name="view">optional NetOffice.AccessApi.Enums.AcView View = 0</param>
		[CustomMethod]
		[SupportByVersion("Access", 9,10,11,12,14,15,16)]
		public virtual void OpenReport(object reportName, object view)
		{
			 InvokerService.InvokeInternal.ExecuteMethod(this, "OpenReport", reportName, view);
		}

		/// <summary>
		/// SupportByVersion Access 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff192676.aspx </remarks>
		/// <param name="reportName">object reportName</param>
		/// <param name="view">optional NetOffice.AccessApi.Enums.AcView View = 0</param>
		/// <param name="filterName">optional object filterName</param>
		[CustomMethod]
		[SupportByVersion("Access", 9,10,11,12,14,15,16)]
		public virtual void OpenReport(object reportName, object view, object filterName)
		{
			 InvokerService.InvokeInternal.ExecuteMethod(this, "OpenReport", reportName, view, filterName);
		}

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
		public virtual void OpenReport(object reportName, object view, object filterName, object whereCondition, object windowMode)
		{
			 InvokerService.InvokeInternal.ExecuteMethod(this, "OpenReport", new object[]{ reportName, view, filterName, whereCondition, windowMode });
		}

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
		public virtual void TransferDatabase(object transferType, object databaseType, object databaseName, object objectType, object source, object destination, object structureOnly, object storeLogin)
		{
			 InvokerService.InvokeInternal.ExecuteMethod(this, "TransferDatabase", new object[]{ transferType, databaseType, databaseName, objectType, source, destination, structureOnly, storeLogin });
		}

		/// <summary>
		/// SupportByVersion Access 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff196455.aspx </remarks>
		[CustomMethod]
		[SupportByVersion("Access", 9,10,11,12,14,15,16)]
		public virtual void TransferDatabase()
		{
			 InvokerService.InvokeInternal.ExecuteMethod(this, "TransferDatabase");
		}

		/// <summary>
		/// SupportByVersion Access 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff196455.aspx </remarks>
		/// <param name="transferType">optional NetOffice.AccessApi.Enums.AcDataTransferType TransferType = 0</param>
		[CustomMethod]
		[SupportByVersion("Access", 9,10,11,12,14,15,16)]
		public virtual void TransferDatabase(object transferType)
		{
			 InvokerService.InvokeInternal.ExecuteMethod(this, "TransferDatabase", transferType);
		}

		/// <summary>
		/// SupportByVersion Access 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff196455.aspx </remarks>
		/// <param name="transferType">optional NetOffice.AccessApi.Enums.AcDataTransferType TransferType = 0</param>
		/// <param name="databaseType">optional object databaseType</param>
		[CustomMethod]
		[SupportByVersion("Access", 9,10,11,12,14,15,16)]
		public virtual void TransferDatabase(object transferType, object databaseType)
		{
			 InvokerService.InvokeInternal.ExecuteMethod(this, "TransferDatabase", transferType, databaseType);
		}

		/// <summary>
		/// SupportByVersion Access 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff196455.aspx </remarks>
		/// <param name="transferType">optional NetOffice.AccessApi.Enums.AcDataTransferType TransferType = 0</param>
		/// <param name="databaseType">optional object databaseType</param>
		/// <param name="databaseName">optional object databaseName</param>
		[CustomMethod]
		[SupportByVersion("Access", 9,10,11,12,14,15,16)]
		public virtual void TransferDatabase(object transferType, object databaseType, object databaseName)
		{
			 InvokerService.InvokeInternal.ExecuteMethod(this, "TransferDatabase", transferType, databaseType, databaseName);
		}

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
		public virtual void TransferDatabase(object transferType, object databaseType, object databaseName, object objectType)
		{
			 InvokerService.InvokeInternal.ExecuteMethod(this, "TransferDatabase", transferType, databaseType, databaseName, objectType);
		}

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
		public virtual void TransferDatabase(object transferType, object databaseType, object databaseName, object objectType, object source)
		{
			 InvokerService.InvokeInternal.ExecuteMethod(this, "TransferDatabase", new object[]{ transferType, databaseType, databaseName, objectType, source });
		}

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
		public virtual void TransferDatabase(object transferType, object databaseType, object databaseName, object objectType, object source, object destination)
		{
			 InvokerService.InvokeInternal.ExecuteMethod(this, "TransferDatabase", new object[]{ transferType, databaseType, databaseName, objectType, source, destination });
		}

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
		public virtual void TransferDatabase(object transferType, object databaseType, object databaseName, object objectType, object source, object destination, object structureOnly)
		{
			 InvokerService.InvokeInternal.ExecuteMethod(this, "TransferDatabase", new object[]{ transferType, databaseType, databaseName, objectType, source, destination, structureOnly });
		}

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
		public virtual void TransferSpreadsheet(object transferType, object spreadsheetType, object tableName, object fileName, object hasFieldNames, object range, object useOA)
		{
			 InvokerService.InvokeInternal.ExecuteMethod(this, "TransferSpreadsheet", new object[]{ transferType, spreadsheetType, tableName, fileName, hasFieldNames, range, useOA });
		}

		/// <summary>
		/// SupportByVersion Access 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff844793.aspx </remarks>
		[CustomMethod]
		[SupportByVersion("Access", 9,10,11,12,14,15,16)]
		public virtual void TransferSpreadsheet()
		{
			 InvokerService.InvokeInternal.ExecuteMethod(this, "TransferSpreadsheet");
		}

		/// <summary>
		/// SupportByVersion Access 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff844793.aspx </remarks>
		/// <param name="transferType">optional NetOffice.AccessApi.Enums.AcDataTransferType TransferType = 0</param>
		[CustomMethod]
		[SupportByVersion("Access", 9,10,11,12,14,15,16)]
		public virtual void TransferSpreadsheet(object transferType)
		{
			 InvokerService.InvokeInternal.ExecuteMethod(this, "TransferSpreadsheet", transferType);
		}

		/// <summary>
		/// SupportByVersion Access 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff844793.aspx </remarks>
		/// <param name="transferType">optional NetOffice.AccessApi.Enums.AcDataTransferType TransferType = 0</param>
		/// <param name="spreadsheetType">optional NetOffice.AccessApi.Enums.AcSpreadSheetType SpreadsheetType = 8</param>
		[CustomMethod]
		[SupportByVersion("Access", 9,10,11,12,14,15,16)]
		public virtual void TransferSpreadsheet(object transferType, object spreadsheetType)
		{
			 InvokerService.InvokeInternal.ExecuteMethod(this, "TransferSpreadsheet", transferType, spreadsheetType);
		}

		/// <summary>
		/// SupportByVersion Access 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff844793.aspx </remarks>
		/// <param name="transferType">optional NetOffice.AccessApi.Enums.AcDataTransferType TransferType = 0</param>
		/// <param name="spreadsheetType">optional NetOffice.AccessApi.Enums.AcSpreadSheetType SpreadsheetType = 8</param>
		/// <param name="tableName">optional object tableName</param>
		[CustomMethod]
		[SupportByVersion("Access", 9,10,11,12,14,15,16)]
		public virtual void TransferSpreadsheet(object transferType, object spreadsheetType, object tableName)
		{
			 InvokerService.InvokeInternal.ExecuteMethod(this, "TransferSpreadsheet", transferType, spreadsheetType, tableName);
		}

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
		public virtual void TransferSpreadsheet(object transferType, object spreadsheetType, object tableName, object fileName)
		{
			 InvokerService.InvokeInternal.ExecuteMethod(this, "TransferSpreadsheet", transferType, spreadsheetType, tableName, fileName);
		}

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
		public virtual void TransferSpreadsheet(object transferType, object spreadsheetType, object tableName, object fileName, object hasFieldNames)
		{
			 InvokerService.InvokeInternal.ExecuteMethod(this, "TransferSpreadsheet", new object[]{ transferType, spreadsheetType, tableName, fileName, hasFieldNames });
		}

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
		public virtual void TransferSpreadsheet(object transferType, object spreadsheetType, object tableName, object fileName, object hasFieldNames, object range)
		{
			 InvokerService.InvokeInternal.ExecuteMethod(this, "TransferSpreadsheet", new object[]{ transferType, spreadsheetType, tableName, fileName, hasFieldNames, range });
		}

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
		public virtual void TransferText(object transferType, object specificationName, object tableName, object fileName, object hasFieldNames, object hTMLTableName, object codePage)
		{
			 InvokerService.InvokeInternal.ExecuteMethod(this, "TransferText", new object[]{ transferType, specificationName, tableName, fileName, hasFieldNames, hTMLTableName, codePage });
		}

		/// <summary>
		/// SupportByVersion Access 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff835958.aspx </remarks>
		[CustomMethod]
		[SupportByVersion("Access", 9,10,11,12,14,15,16)]
		public virtual void TransferText()
		{
			 InvokerService.InvokeInternal.ExecuteMethod(this, "TransferText");
		}

		/// <summary>
		/// SupportByVersion Access 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff835958.aspx </remarks>
		/// <param name="transferType">optional NetOffice.AccessApi.Enums.AcTextTransferType TransferType = 0</param>
		[CustomMethod]
		[SupportByVersion("Access", 9,10,11,12,14,15,16)]
		public virtual void TransferText(object transferType)
		{
			 InvokerService.InvokeInternal.ExecuteMethod(this, "TransferText", transferType);
		}

		/// <summary>
		/// SupportByVersion Access 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff835958.aspx </remarks>
		/// <param name="transferType">optional NetOffice.AccessApi.Enums.AcTextTransferType TransferType = 0</param>
		/// <param name="specificationName">optional object specificationName</param>
		[CustomMethod]
		[SupportByVersion("Access", 9,10,11,12,14,15,16)]
		public virtual void TransferText(object transferType, object specificationName)
		{
			 InvokerService.InvokeInternal.ExecuteMethod(this, "TransferText", transferType, specificationName);
		}

		/// <summary>
		/// SupportByVersion Access 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff835958.aspx </remarks>
		/// <param name="transferType">optional NetOffice.AccessApi.Enums.AcTextTransferType TransferType = 0</param>
		/// <param name="specificationName">optional object specificationName</param>
		/// <param name="tableName">optional object tableName</param>
		[CustomMethod]
		[SupportByVersion("Access", 9,10,11,12,14,15,16)]
		public virtual void TransferText(object transferType, object specificationName, object tableName)
		{
			 InvokerService.InvokeInternal.ExecuteMethod(this, "TransferText", transferType, specificationName, tableName);
		}

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
		public virtual void TransferText(object transferType, object specificationName, object tableName, object fileName)
		{
			 InvokerService.InvokeInternal.ExecuteMethod(this, "TransferText", transferType, specificationName, tableName, fileName);
		}

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
		public virtual void TransferText(object transferType, object specificationName, object tableName, object fileName, object hasFieldNames)
		{
			 InvokerService.InvokeInternal.ExecuteMethod(this, "TransferText", new object[]{ transferType, specificationName, tableName, fileName, hasFieldNames });
		}

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
		public virtual void TransferText(object transferType, object specificationName, object tableName, object fileName, object hasFieldNames, object hTMLTableName)
		{
			 InvokerService.InvokeInternal.ExecuteMethod(this, "TransferText", new object[]{ transferType, specificationName, tableName, fileName, hasFieldNames, hTMLTableName });
		}

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
		public virtual void OutputTo(NetOffice.AccessApi.Enums.AcOutputObjectType objectType, object objectName, object outputFormat, object outputFile, object autoStart, object templateFile)
		{
			 InvokerService.InvokeInternal.ExecuteMethod(this, "OutputTo", new object[]{ objectType, objectName, outputFormat, outputFile, autoStart, templateFile });
		}

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
		public virtual void OutputTo(NetOffice.AccessApi.Enums.AcOutputObjectType objectType, object objectName, object outputFormat, object outputFile, object autoStart, object templateFile, object encoding)
		{
			 InvokerService.InvokeInternal.ExecuteMethod(this, "OutputTo", new object[]{ objectType, objectName, outputFormat, outputFile, autoStart, templateFile, encoding });
		}

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
		public virtual void OutputTo(NetOffice.AccessApi.Enums.AcOutputObjectType objectType, object objectName, object outputFormat, object outputFile, object autoStart, object templateFile, object encoding, object outputQuality)
		{
			 InvokerService.InvokeInternal.ExecuteMethod(this, "OutputTo", new object[]{ objectType, objectName, outputFormat, outputFile, autoStart, templateFile, encoding, outputQuality });
		}

		/// <summary>
		/// SupportByVersion Access 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff192065.aspx </remarks>
		/// <param name="objectType">NetOffice.AccessApi.Enums.AcOutputObjectType objectType</param>
		[CustomMethod]
		[SupportByVersion("Access", 9,10,11,12,14,15,16)]
		public virtual void OutputTo(NetOffice.AccessApi.Enums.AcOutputObjectType objectType)
		{
			 InvokerService.InvokeInternal.ExecuteMethod(this, "OutputTo", objectType);
		}

		/// <summary>
		/// SupportByVersion Access 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff192065.aspx </remarks>
		/// <param name="objectType">NetOffice.AccessApi.Enums.AcOutputObjectType objectType</param>
		/// <param name="objectName">optional object objectName</param>
		[CustomMethod]
		[SupportByVersion("Access", 9,10,11,12,14,15,16)]
		public virtual void OutputTo(NetOffice.AccessApi.Enums.AcOutputObjectType objectType, object objectName)
		{
			 InvokerService.InvokeInternal.ExecuteMethod(this, "OutputTo", objectType, objectName);
		}

		/// <summary>
		/// SupportByVersion Access 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff192065.aspx </remarks>
		/// <param name="objectType">NetOffice.AccessApi.Enums.AcOutputObjectType objectType</param>
		/// <param name="objectName">optional object objectName</param>
		/// <param name="outputFormat">optional object outputFormat</param>
		[CustomMethod]
		[SupportByVersion("Access", 9,10,11,12,14,15,16)]
		public virtual void OutputTo(NetOffice.AccessApi.Enums.AcOutputObjectType objectType, object objectName, object outputFormat)
		{
			 InvokerService.InvokeInternal.ExecuteMethod(this, "OutputTo", objectType, objectName, outputFormat);
		}

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
		public virtual void OutputTo(NetOffice.AccessApi.Enums.AcOutputObjectType objectType, object objectName, object outputFormat, object outputFile)
		{
			 InvokerService.InvokeInternal.ExecuteMethod(this, "OutputTo", objectType, objectName, outputFormat, outputFile);
		}

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
		public virtual void OutputTo(NetOffice.AccessApi.Enums.AcOutputObjectType objectType, object objectName, object outputFormat, object outputFile, object autoStart)
		{
			 InvokerService.InvokeInternal.ExecuteMethod(this, "OutputTo", new object[]{ objectType, objectName, outputFormat, outputFile, autoStart });
		}

		/// <summary>
		/// SupportByVersion Access 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff197376.aspx </remarks>
		/// <param name="objectType">optional NetOffice.AccessApi.Enums.AcObjectType ObjectType = -1</param>
		/// <param name="objectName">optional object objectName</param>
		[SupportByVersion("Access", 9,10,11,12,14,15,16)]
		public virtual void DeleteObject(object objectType, object objectName)
		{
			 InvokerService.InvokeInternal.ExecuteMethod(this, "DeleteObject", objectType, objectName);
		}

		/// <summary>
		/// SupportByVersion Access 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff197376.aspx </remarks>
		[CustomMethod]
		[SupportByVersion("Access", 9,10,11,12,14,15,16)]
		public virtual void DeleteObject()
		{
			 InvokerService.InvokeInternal.ExecuteMethod(this, "DeleteObject");
		}

		/// <summary>
		/// SupportByVersion Access 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff197376.aspx </remarks>
		/// <param name="objectType">optional NetOffice.AccessApi.Enums.AcObjectType ObjectType = -1</param>
		[CustomMethod]
		[SupportByVersion("Access", 9,10,11,12,14,15,16)]
		public virtual void DeleteObject(object objectType)
		{
			 InvokerService.InvokeInternal.ExecuteMethod(this, "DeleteObject", objectType);
		}

		/// <summary>
		/// SupportByVersion Access 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff192698.aspx </remarks>
		/// <param name="moduleName">optional object moduleName</param>
		/// <param name="procedureName">optional object procedureName</param>
		[SupportByVersion("Access", 9,10,11,12,14,15,16)]
		public virtual void OpenModule(object moduleName, object procedureName)
		{
			 InvokerService.InvokeInternal.ExecuteMethod(this, "OpenModule", moduleName, procedureName);
		}

		/// <summary>
		/// SupportByVersion Access 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff192698.aspx </remarks>
		[CustomMethod]
		[SupportByVersion("Access", 9,10,11,12,14,15,16)]
		public virtual void OpenModule()
		{
			 InvokerService.InvokeInternal.ExecuteMethod(this, "OpenModule");
		}

		/// <summary>
		/// SupportByVersion Access 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff192698.aspx </remarks>
		/// <param name="moduleName">optional object moduleName</param>
		[CustomMethod]
		[SupportByVersion("Access", 9,10,11,12,14,15,16)]
		public virtual void OpenModule(object moduleName)
		{
			 InvokerService.InvokeInternal.ExecuteMethod(this, "OpenModule", moduleName);
		}

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
		public virtual void SendObject(object objectType, object objectName, object outputFormat, object to, object cc, object bcc, object subject, object messageText, object editMessage, object templateFile)
		{
			 InvokerService.InvokeInternal.ExecuteMethod(this, "SendObject", new object[]{ objectType, objectName, outputFormat, to, cc, bcc, subject, messageText, editMessage, templateFile });
		}

		/// <summary>
		/// SupportByVersion Access 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff197046.aspx </remarks>
		[CustomMethod]
		[SupportByVersion("Access", 9,10,11,12,14,15,16)]
		public virtual void SendObject()
		{
			 InvokerService.InvokeInternal.ExecuteMethod(this, "SendObject");
		}

		/// <summary>
		/// SupportByVersion Access 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff197046.aspx </remarks>
		/// <param name="objectType">optional NetOffice.AccessApi.Enums.AcSendObjectType ObjectType = -1</param>
		[CustomMethod]
		[SupportByVersion("Access", 9,10,11,12,14,15,16)]
		public virtual void SendObject(object objectType)
		{
			 InvokerService.InvokeInternal.ExecuteMethod(this, "SendObject", objectType);
		}

		/// <summary>
		/// SupportByVersion Access 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff197046.aspx </remarks>
		/// <param name="objectType">optional NetOffice.AccessApi.Enums.AcSendObjectType ObjectType = -1</param>
		/// <param name="objectName">optional object objectName</param>
		[CustomMethod]
		[SupportByVersion("Access", 9,10,11,12,14,15,16)]
		public virtual void SendObject(object objectType, object objectName)
		{
			 InvokerService.InvokeInternal.ExecuteMethod(this, "SendObject", objectType, objectName);
		}

		/// <summary>
		/// SupportByVersion Access 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff197046.aspx </remarks>
		/// <param name="objectType">optional NetOffice.AccessApi.Enums.AcSendObjectType ObjectType = -1</param>
		/// <param name="objectName">optional object objectName</param>
		/// <param name="outputFormat">optional object outputFormat</param>
		[CustomMethod]
		[SupportByVersion("Access", 9,10,11,12,14,15,16)]
		public virtual void SendObject(object objectType, object objectName, object outputFormat)
		{
			 InvokerService.InvokeInternal.ExecuteMethod(this, "SendObject", objectType, objectName, outputFormat);
		}

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
		public virtual void SendObject(object objectType, object objectName, object outputFormat, object to)
		{
			 InvokerService.InvokeInternal.ExecuteMethod(this, "SendObject", objectType, objectName, outputFormat, to);
		}

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
		public virtual void SendObject(object objectType, object objectName, object outputFormat, object to, object cc)
		{
			 InvokerService.InvokeInternal.ExecuteMethod(this, "SendObject", new object[]{ objectType, objectName, outputFormat, to, cc });
		}

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
		public virtual void SendObject(object objectType, object objectName, object outputFormat, object to, object cc, object bcc)
		{
			 InvokerService.InvokeInternal.ExecuteMethod(this, "SendObject", new object[]{ objectType, objectName, outputFormat, to, cc, bcc });
		}

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
		public virtual void SendObject(object objectType, object objectName, object outputFormat, object to, object cc, object bcc, object subject)
		{
			 InvokerService.InvokeInternal.ExecuteMethod(this, "SendObject", new object[]{ objectType, objectName, outputFormat, to, cc, bcc, subject });
		}

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
		public virtual void SendObject(object objectType, object objectName, object outputFormat, object to, object cc, object bcc, object subject, object messageText)
		{
			 InvokerService.InvokeInternal.ExecuteMethod(this, "SendObject", new object[]{ objectType, objectName, outputFormat, to, cc, bcc, subject, messageText });
		}

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
		public virtual void SendObject(object objectType, object objectName, object outputFormat, object to, object cc, object bcc, object subject, object messageText, object editMessage)
		{
			 InvokerService.InvokeInternal.ExecuteMethod(this, "SendObject", new object[]{ objectType, objectName, outputFormat, to, cc, bcc, subject, messageText, editMessage });
		}

		/// <summary>
		/// SupportByVersion Access 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff194957.aspx </remarks>
		/// <param name="toolbarName">object toolbarName</param>
		/// <param name="show">optional NetOffice.AccessApi.Enums.AcShowToolbar Show = 0</param>
		[SupportByVersion("Access", 9,10,11,12,14,15,16)]
		public virtual void ShowToolbar(object toolbarName, object show)
		{
			 InvokerService.InvokeInternal.ExecuteMethod(this, "ShowToolbar", toolbarName, show);
		}

		/// <summary>
		/// SupportByVersion Access 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff194957.aspx </remarks>
		/// <param name="toolbarName">object toolbarName</param>
		[CustomMethod]
		[SupportByVersion("Access", 9,10,11,12,14,15,16)]
		public virtual void ShowToolbar(object toolbarName)
		{
			 InvokerService.InvokeInternal.ExecuteMethod(this, "ShowToolbar", toolbarName);
		}

		/// <summary>
		/// SupportByVersion Access 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff196435.aspx </remarks>
		/// <param name="objectType">optional NetOffice.AccessApi.Enums.AcObjectType ObjectType = -1</param>
		/// <param name="objectName">optional object objectName</param>
		[SupportByVersion("Access", 9,10,11,12,14,15,16)]
		public virtual void Save(object objectType, object objectName)
		{
			 InvokerService.InvokeInternal.ExecuteMethod(this, "Save", objectType, objectName);
		}

		/// <summary>
		/// SupportByVersion Access 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff196435.aspx </remarks>
		[CustomMethod]
		[SupportByVersion("Access", 9,10,11,12,14,15,16)]
		public virtual void Save()
		{
			 InvokerService.InvokeInternal.ExecuteMethod(this, "Save");
		}

		/// <summary>
		/// SupportByVersion Access 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff196435.aspx </remarks>
		/// <param name="objectType">optional NetOffice.AccessApi.Enums.AcObjectType ObjectType = -1</param>
		[CustomMethod]
		[SupportByVersion("Access", 9,10,11,12,14,15,16)]
		public virtual void Save(object objectType)
		{
			 InvokerService.InvokeInternal.ExecuteMethod(this, "Save", objectType);
		}

		/// <summary>
		/// SupportByVersion Access 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff195265.aspx </remarks>
		/// <param name="menuIndex">object menuIndex</param>
		/// <param name="commandIndex">optional object commandIndex</param>
		/// <param name="subcommandIndex">optional object subcommandIndex</param>
		/// <param name="flag">optional object flag</param>
		[SupportByVersion("Access", 9,10,11,12,14,15,16)]
		public virtual void SetMenuItem(object menuIndex, object commandIndex, object subcommandIndex, object flag)
		{
			 InvokerService.InvokeInternal.ExecuteMethod(this, "SetMenuItem", menuIndex, commandIndex, subcommandIndex, flag);
		}

		/// <summary>
		/// SupportByVersion Access 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff195265.aspx </remarks>
		/// <param name="menuIndex">object menuIndex</param>
		[CustomMethod]
		[SupportByVersion("Access", 9,10,11,12,14,15,16)]
		public virtual void SetMenuItem(object menuIndex)
		{
			 InvokerService.InvokeInternal.ExecuteMethod(this, "SetMenuItem", menuIndex);
		}

		/// <summary>
		/// SupportByVersion Access 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff195265.aspx </remarks>
		/// <param name="menuIndex">object menuIndex</param>
		/// <param name="commandIndex">optional object commandIndex</param>
		[CustomMethod]
		[SupportByVersion("Access", 9,10,11,12,14,15,16)]
		public virtual void SetMenuItem(object menuIndex, object commandIndex)
		{
			 InvokerService.InvokeInternal.ExecuteMethod(this, "SetMenuItem", menuIndex, commandIndex);
		}

		/// <summary>
		/// SupportByVersion Access 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff195265.aspx </remarks>
		/// <param name="menuIndex">object menuIndex</param>
		/// <param name="commandIndex">optional object commandIndex</param>
		/// <param name="subcommandIndex">optional object subcommandIndex</param>
		[CustomMethod]
		[SupportByVersion("Access", 9,10,11,12,14,15,16)]
		public virtual void SetMenuItem(object menuIndex, object commandIndex, object subcommandIndex)
		{
			 InvokerService.InvokeInternal.ExecuteMethod(this, "SetMenuItem", menuIndex, commandIndex, subcommandIndex);
		}

		/// <summary>
		/// SupportByVersion Access 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff194612.aspx </remarks>
		/// <param name="command">NetOffice.AccessApi.Enums.AcCommand command</param>
		[SupportByVersion("Access", 9,10,11,12,14,15,16)]
		public virtual void RunCommand(NetOffice.AccessApi.Enums.AcCommand command)
		{
			 InvokerService.InvokeInternal.ExecuteMethod(this, "RunCommand", command);
		}

		/// <summary>
		/// SupportByVersion Access 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff845421.aspx </remarks>
		/// <param name="dataAccessPageName">object dataAccessPageName</param>
		/// <param name="view">optional NetOffice.AccessApi.Enums.AcDataAccessPageView View = 0</param>
		[SupportByVersion("Access", 9,10,11,12,14,15,16)]
		public virtual void OpenDataAccessPage(object dataAccessPageName, object view)
		{
			 InvokerService.InvokeInternal.ExecuteMethod(this, "OpenDataAccessPage", dataAccessPageName, view);
		}

		/// <summary>
		/// SupportByVersion Access 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff845421.aspx </remarks>
		/// <param name="dataAccessPageName">object dataAccessPageName</param>
		[CustomMethod]
		[SupportByVersion("Access", 9,10,11,12,14,15,16)]
		public virtual void OpenDataAccessPage(object dataAccessPageName)
		{
			 InvokerService.InvokeInternal.ExecuteMethod(this, "OpenDataAccessPage", dataAccessPageName);
		}

		/// <summary>
		/// SupportByVersion Access 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff197347.aspx </remarks>
		/// <param name="viewName">object viewName</param>
		/// <param name="view">optional NetOffice.AccessApi.Enums.AcView View = 0</param>
		/// <param name="dataMode">optional NetOffice.AccessApi.Enums.AcOpenDataMode DataMode = 1</param>
		[SupportByVersion("Access", 9,10,11,12,14,15,16)]
		public virtual void OpenView(object viewName, object view, object dataMode)
		{
			 InvokerService.InvokeInternal.ExecuteMethod(this, "OpenView", viewName, view, dataMode);
		}

		/// <summary>
		/// SupportByVersion Access 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff197347.aspx </remarks>
		/// <param name="viewName">object viewName</param>
		[CustomMethod]
		[SupportByVersion("Access", 9,10,11,12,14,15,16)]
		public virtual void OpenView(object viewName)
		{
			 InvokerService.InvokeInternal.ExecuteMethod(this, "OpenView", viewName);
		}

		/// <summary>
		/// SupportByVersion Access 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff197347.aspx </remarks>
		/// <param name="viewName">object viewName</param>
		/// <param name="view">optional NetOffice.AccessApi.Enums.AcView View = 0</param>
		[CustomMethod]
		[SupportByVersion("Access", 9,10,11,12,14,15,16)]
		public virtual void OpenView(object viewName, object view)
		{
			 InvokerService.InvokeInternal.ExecuteMethod(this, "OpenView", viewName, view);
		}

		/// <summary>
		/// SupportByVersion Access 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff821439.aspx </remarks>
		/// <param name="diagramName">object diagramName</param>
		[SupportByVersion("Access", 9,10,11,12,14,15,16)]
		public virtual void OpenDiagram(object diagramName)
		{
			 InvokerService.InvokeInternal.ExecuteMethod(this, "OpenDiagram", diagramName);
		}

		/// <summary>
		/// SupportByVersion Access 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff197412.aspx </remarks>
		/// <param name="procedureName">object procedureName</param>
		/// <param name="view">optional NetOffice.AccessApi.Enums.AcView View = 0</param>
		/// <param name="dataMode">optional NetOffice.AccessApi.Enums.AcOpenDataMode DataMode = 1</param>
		[SupportByVersion("Access", 9,10,11,12,14,15,16)]
		public virtual void OpenStoredProcedure(object procedureName, object view, object dataMode)
		{
			 InvokerService.InvokeInternal.ExecuteMethod(this, "OpenStoredProcedure", procedureName, view, dataMode);
		}

		/// <summary>
		/// SupportByVersion Access 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff197412.aspx </remarks>
		/// <param name="procedureName">object procedureName</param>
		[CustomMethod]
		[SupportByVersion("Access", 9,10,11,12,14,15,16)]
		public virtual void OpenStoredProcedure(object procedureName)
		{
			 InvokerService.InvokeInternal.ExecuteMethod(this, "OpenStoredProcedure", procedureName);
		}

		/// <summary>
		/// SupportByVersion Access 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff197412.aspx </remarks>
		/// <param name="procedureName">object procedureName</param>
		/// <param name="view">optional NetOffice.AccessApi.Enums.AcView View = 0</param>
		[CustomMethod]
		[SupportByVersion("Access", 9,10,11,12,14,15,16)]
		public virtual void OpenStoredProcedure(object procedureName, object view)
		{
			 InvokerService.InvokeInternal.ExecuteMethod(this, "OpenStoredProcedure", procedureName, view);
		}

		/// <summary>
		/// SupportByVersion Access 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <param name="reportName">object reportName</param>
		/// <param name="view">optional NetOffice.AccessApi.Enums.AcView View = 0</param>
		/// <param name="filterName">optional object filterName</param>
		/// <param name="whereCondition">optional object whereCondition</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		[SupportByVersion("Access", 10,11,12,14,15,16)]
		public virtual void OpenReportOld0(object reportName, object view, object filterName, object whereCondition)
		{
			 InvokerService.InvokeInternal.ExecuteMethod(this, "OpenReportOld0", reportName, view, filterName, whereCondition);
		}

		/// <summary>
		/// SupportByVersion Access 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <param name="reportName">object reportName</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		[CustomMethod]
		[SupportByVersion("Access", 10,11,12,14,15,16)]
		public virtual void OpenReportOld0(object reportName)
		{
			 InvokerService.InvokeInternal.ExecuteMethod(this, "OpenReportOld0", reportName);
		}

		/// <summary>
		/// SupportByVersion Access 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <param name="reportName">object reportName</param>
		/// <param name="view">optional NetOffice.AccessApi.Enums.AcView View = 0</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		[CustomMethod]
		[SupportByVersion("Access", 10,11,12,14,15,16)]
		public virtual void OpenReportOld0(object reportName, object view)
		{
			 InvokerService.InvokeInternal.ExecuteMethod(this, "OpenReportOld0", reportName, view);
		}

		/// <summary>
		/// SupportByVersion Access 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <param name="reportName">object reportName</param>
		/// <param name="view">optional NetOffice.AccessApi.Enums.AcView View = 0</param>
		/// <param name="filterName">optional object filterName</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		[CustomMethod]
		[SupportByVersion("Access", 10,11,12,14,15,16)]
		public virtual void OpenReportOld0(object reportName, object view, object filterName)
		{
			 InvokerService.InvokeInternal.ExecuteMethod(this, "OpenReportOld0", reportName, view, filterName);
		}

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
		public virtual void OutputToOld0(NetOffice.AccessApi.Enums.AcOutputObjectType objectType, object objectName, object outputFormat, object outputFile, object autoStart, object templateFile)
		{
			 InvokerService.InvokeInternal.ExecuteMethod(this, "OutputToOld0", new object[]{ objectType, objectName, outputFormat, outputFile, autoStart, templateFile });
		}

		/// <summary>
		/// SupportByVersion Access 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <param name="objectType">NetOffice.AccessApi.Enums.AcOutputObjectType objectType</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		[CustomMethod]
		[SupportByVersion("Access", 10,11,12,14,15,16)]
		public virtual void OutputToOld0(NetOffice.AccessApi.Enums.AcOutputObjectType objectType)
		{
			 InvokerService.InvokeInternal.ExecuteMethod(this, "OutputToOld0", objectType);
		}

		/// <summary>
		/// SupportByVersion Access 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <param name="objectType">NetOffice.AccessApi.Enums.AcOutputObjectType objectType</param>
		/// <param name="objectName">optional object objectName</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		[CustomMethod]
		[SupportByVersion("Access", 10,11,12,14,15,16)]
		public virtual void OutputToOld0(NetOffice.AccessApi.Enums.AcOutputObjectType objectType, object objectName)
		{
			 InvokerService.InvokeInternal.ExecuteMethod(this, "OutputToOld0", objectType, objectName);
		}

		/// <summary>
		/// SupportByVersion Access 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <param name="objectType">NetOffice.AccessApi.Enums.AcOutputObjectType objectType</param>
		/// <param name="objectName">optional object objectName</param>
		/// <param name="outputFormat">optional object outputFormat</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		[CustomMethod]
		[SupportByVersion("Access", 10,11,12,14,15,16)]
		public virtual void OutputToOld0(NetOffice.AccessApi.Enums.AcOutputObjectType objectType, object objectName, object outputFormat)
		{
			 InvokerService.InvokeInternal.ExecuteMethod(this, "OutputToOld0", objectType, objectName, outputFormat);
		}

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
		public virtual void OutputToOld0(NetOffice.AccessApi.Enums.AcOutputObjectType objectType, object objectName, object outputFormat, object outputFile)
		{
			 InvokerService.InvokeInternal.ExecuteMethod(this, "OutputToOld0", objectType, objectName, outputFormat, outputFile);
		}

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
		public virtual void OutputToOld0(NetOffice.AccessApi.Enums.AcOutputObjectType objectType, object objectName, object outputFormat, object outputFile, object autoStart)
		{
			 InvokerService.InvokeInternal.ExecuteMethod(this, "OutputToOld0", new object[]{ objectType, objectName, outputFormat, outputFile, autoStart });
		}

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
		public virtual void TransferSQLDatabase(object server, object database, object useTrustedConnection, object login, object password, object transferCopyData)
		{
			 InvokerService.InvokeInternal.ExecuteMethod(this, "TransferSQLDatabase", new object[]{ server, database, useTrustedConnection, login, password, transferCopyData });
		}

		/// <summary>
		/// SupportByVersion Access 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff835048.aspx </remarks>
		/// <param name="server">object server</param>
		/// <param name="database">object database</param>
		[CustomMethod]
		[SupportByVersion("Access", 10,11,12,14,15,16)]
		public virtual void TransferSQLDatabase(object server, object database)
		{
			 InvokerService.InvokeInternal.ExecuteMethod(this, "TransferSQLDatabase", server, database);
		}

		/// <summary>
		/// SupportByVersion Access 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff835048.aspx </remarks>
		/// <param name="server">object server</param>
		/// <param name="database">object database</param>
		/// <param name="useTrustedConnection">optional object useTrustedConnection</param>
		[CustomMethod]
		[SupportByVersion("Access", 10,11,12,14,15,16)]
		public virtual void TransferSQLDatabase(object server, object database, object useTrustedConnection)
		{
			 InvokerService.InvokeInternal.ExecuteMethod(this, "TransferSQLDatabase", server, database, useTrustedConnection);
		}

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
		public virtual void TransferSQLDatabase(object server, object database, object useTrustedConnection, object login)
		{
			 InvokerService.InvokeInternal.ExecuteMethod(this, "TransferSQLDatabase", server, database, useTrustedConnection, login);
		}

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
		public virtual void TransferSQLDatabase(object server, object database, object useTrustedConnection, object login, object password)
		{
			 InvokerService.InvokeInternal.ExecuteMethod(this, "TransferSQLDatabase", new object[]{ server, database, useTrustedConnection, login, password });
		}

		/// <summary>
		/// SupportByVersion Access 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff845497.aspx </remarks>
		/// <param name="databaseFileName">object databaseFileName</param>
		/// <param name="overwriteExistingFile">optional object overwriteExistingFile</param>
		/// <param name="disconnectAllUsers">optional object disconnectAllUsers</param>
		[SupportByVersion("Access", 10,11,12,14,15,16)]
		public virtual void CopyDatabaseFile(object databaseFileName, object overwriteExistingFile, object disconnectAllUsers)
		{
			 InvokerService.InvokeInternal.ExecuteMethod(this, "CopyDatabaseFile", databaseFileName, overwriteExistingFile, disconnectAllUsers);
		}

		/// <summary>
		/// SupportByVersion Access 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff845497.aspx </remarks>
		/// <param name="databaseFileName">object databaseFileName</param>
		[CustomMethod]
		[SupportByVersion("Access", 10,11,12,14,15,16)]
		public virtual void CopyDatabaseFile(object databaseFileName)
		{
			 InvokerService.InvokeInternal.ExecuteMethod(this, "CopyDatabaseFile", databaseFileName);
		}

		/// <summary>
		/// SupportByVersion Access 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff845497.aspx </remarks>
		/// <param name="databaseFileName">object databaseFileName</param>
		/// <param name="overwriteExistingFile">optional object overwriteExistingFile</param>
		[CustomMethod]
		[SupportByVersion("Access", 10,11,12,14,15,16)]
		public virtual void CopyDatabaseFile(object databaseFileName, object overwriteExistingFile)
		{
			 InvokerService.InvokeInternal.ExecuteMethod(this, "CopyDatabaseFile", databaseFileName, overwriteExistingFile);
		}

		/// <summary>
		/// SupportByVersion Access 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff194192.aspx </remarks>
		/// <param name="functionName">object functionName</param>
		/// <param name="view">optional NetOffice.AccessApi.Enums.AcView View = 0</param>
		/// <param name="dataMode">optional NetOffice.AccessApi.Enums.AcOpenDataMode DataMode = 1</param>
		[SupportByVersion("Access", 10,11,12,14,15,16)]
		public virtual void OpenFunction(object functionName, object view, object dataMode)
		{
			 InvokerService.InvokeInternal.ExecuteMethod(this, "OpenFunction", functionName, view, dataMode);
		}

		/// <summary>
		/// SupportByVersion Access 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff194192.aspx </remarks>
		/// <param name="functionName">object functionName</param>
		[CustomMethod]
		[SupportByVersion("Access", 10,11,12,14,15,16)]
		public virtual void OpenFunction(object functionName)
		{
			 InvokerService.InvokeInternal.ExecuteMethod(this, "OpenFunction", functionName);
		}

		/// <summary>
		/// SupportByVersion Access 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff194192.aspx </remarks>
		/// <param name="functionName">object functionName</param>
		/// <param name="view">optional NetOffice.AccessApi.Enums.AcView View = 0</param>
		[CustomMethod]
		[SupportByVersion("Access", 10,11,12,14,15,16)]
		public virtual void OpenFunction(object functionName, object view)
		{
			 InvokerService.InvokeInternal.ExecuteMethod(this, "OpenFunction", functionName, view);
		}

		/// <summary>
		/// SupportByVersion Access 12, 14, 15, 16
		/// </summary>
		/// <param name="filterName">optional object filterName</param>
		/// <param name="whereCondition">optional object whereCondition</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		[SupportByVersion("Access", 12,14,15,16)]
		public virtual void ApplyFilterOld0(object filterName, object whereCondition)
		{
			 InvokerService.InvokeInternal.ExecuteMethod(this, "ApplyFilterOld0", filterName, whereCondition);
		}

		/// <summary>
		/// SupportByVersion Access 12, 14, 15, 16
		/// </summary>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		[CustomMethod]
		[SupportByVersion("Access", 12,14,15,16)]
		public virtual void ApplyFilterOld0()
		{
			 InvokerService.InvokeInternal.ExecuteMethod(this, "ApplyFilterOld0");
		}

		/// <summary>
		/// SupportByVersion Access 12, 14, 15, 16
		/// </summary>
		/// <param name="filterName">optional object filterName</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		[CustomMethod]
		[SupportByVersion("Access", 12,14,15,16)]
		public virtual void ApplyFilterOld0(object filterName)
		{
			 InvokerService.InvokeInternal.ExecuteMethod(this, "ApplyFilterOld0", filterName);
		}

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
		public virtual void OutputToOld1(NetOffice.AccessApi.Enums.AcOutputObjectType objectType, object objectName, object outputFormat, object outputFile, object autoStart, object templateFile, object encoding)
		{
			 InvokerService.InvokeInternal.ExecuteMethod(this, "OutputToOld1", new object[]{ objectType, objectName, outputFormat, outputFile, autoStart, templateFile, encoding });
		}

		/// <summary>
		/// SupportByVersion Access 12, 14, 15, 16
		/// </summary>
		/// <param name="objectType">NetOffice.AccessApi.Enums.AcOutputObjectType objectType</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		[CustomMethod]
		[SupportByVersion("Access", 12,14,15,16)]
		public virtual void OutputToOld1(NetOffice.AccessApi.Enums.AcOutputObjectType objectType)
		{
			 InvokerService.InvokeInternal.ExecuteMethod(this, "OutputToOld1", objectType);
		}

		/// <summary>
		/// SupportByVersion Access 12, 14, 15, 16
		/// </summary>
		/// <param name="objectType">NetOffice.AccessApi.Enums.AcOutputObjectType objectType</param>
		/// <param name="objectName">optional object objectName</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		[CustomMethod]
		[SupportByVersion("Access", 12,14,15,16)]
		public virtual void OutputToOld1(NetOffice.AccessApi.Enums.AcOutputObjectType objectType, object objectName)
		{
			 InvokerService.InvokeInternal.ExecuteMethod(this, "OutputToOld1", objectType, objectName);
		}

		/// <summary>
		/// SupportByVersion Access 12, 14, 15, 16
		/// </summary>
		/// <param name="objectType">NetOffice.AccessApi.Enums.AcOutputObjectType objectType</param>
		/// <param name="objectName">optional object objectName</param>
		/// <param name="outputFormat">optional object outputFormat</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		[CustomMethod]
		[SupportByVersion("Access", 12,14,15,16)]
		public virtual void OutputToOld1(NetOffice.AccessApi.Enums.AcOutputObjectType objectType, object objectName, object outputFormat)
		{
			 InvokerService.InvokeInternal.ExecuteMethod(this, "OutputToOld1", objectType, objectName, outputFormat);
		}

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
		public virtual void OutputToOld1(NetOffice.AccessApi.Enums.AcOutputObjectType objectType, object objectName, object outputFormat, object outputFile)
		{
			 InvokerService.InvokeInternal.ExecuteMethod(this, "OutputToOld1", objectType, objectName, outputFormat, outputFile);
		}

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
		public virtual void OutputToOld1(NetOffice.AccessApi.Enums.AcOutputObjectType objectType, object objectName, object outputFormat, object outputFile, object autoStart)
		{
			 InvokerService.InvokeInternal.ExecuteMethod(this, "OutputToOld1", new object[]{ objectType, objectName, outputFormat, outputFile, autoStart });
		}

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
		public virtual void OutputToOld1(NetOffice.AccessApi.Enums.AcOutputObjectType objectType, object objectName, object outputFormat, object outputFile, object autoStart, object templateFile)
		{
			 InvokerService.InvokeInternal.ExecuteMethod(this, "OutputToOld1", new object[]{ objectType, objectName, outputFormat, outputFile, autoStart, templateFile });
		}

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
		public virtual void TransferSharePointList(NetOffice.AccessApi.Enums.AcSharePointListTransferType transferType, object siteAddress, object listID, object viewID, object tableName, object getLookupDisplayValues)
		{
			 InvokerService.InvokeInternal.ExecuteMethod(this, "TransferSharePointList", new object[]{ transferType, siteAddress, listID, viewID, tableName, getLookupDisplayValues });
		}

		/// <summary>
		/// SupportByVersion Access 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff198137.aspx </remarks>
		/// <param name="transferType">NetOffice.AccessApi.Enums.AcSharePointListTransferType transferType</param>
		/// <param name="siteAddress">object siteAddress</param>
		/// <param name="listID">object listID</param>
		[CustomMethod]
		[SupportByVersion("Access", 12,14,15,16)]
		public virtual void TransferSharePointList(NetOffice.AccessApi.Enums.AcSharePointListTransferType transferType, object siteAddress, object listID)
		{
			 InvokerService.InvokeInternal.ExecuteMethod(this, "TransferSharePointList", transferType, siteAddress, listID);
		}

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
		public virtual void TransferSharePointList(NetOffice.AccessApi.Enums.AcSharePointListTransferType transferType, object siteAddress, object listID, object viewID)
		{
			 InvokerService.InvokeInternal.ExecuteMethod(this, "TransferSharePointList", transferType, siteAddress, listID, viewID);
		}

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
		public virtual void TransferSharePointList(NetOffice.AccessApi.Enums.AcSharePointListTransferType transferType, object siteAddress, object listID, object viewID, object tableName)
		{
			 InvokerService.InvokeInternal.ExecuteMethod(this, "TransferSharePointList", new object[]{ transferType, siteAddress, listID, viewID, tableName });
		}

		/// <summary>
		/// SupportByVersion Access 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff844747.aspx </remarks>
		[SupportByVersion("Access", 12,14,15,16)]
		public virtual void CloseDatabase()
		{
			 InvokerService.InvokeInternal.ExecuteMethod(this, "CloseDatabase");
		}

		/// <summary>
		/// SupportByVersion Access 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff191916.aspx </remarks>
		/// <param name="category">optional object category</param>
		/// <param name="group">optional object group</param>
		[SupportByVersion("Access", 12,14,15,16)]
		public virtual void NavigateTo(object category, object group)
		{
			 InvokerService.InvokeInternal.ExecuteMethod(this, "NavigateTo", category, group);
		}

		/// <summary>
		/// SupportByVersion Access 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff191916.aspx </remarks>
		[CustomMethod]
		[SupportByVersion("Access", 12,14,15,16)]
		public virtual void NavigateTo()
		{
			 InvokerService.InvokeInternal.ExecuteMethod(this, "NavigateTo");
		}

		/// <summary>
		/// SupportByVersion Access 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff191916.aspx </remarks>
		/// <param name="category">optional object category</param>
		[CustomMethod]
		[SupportByVersion("Access", 12,14,15,16)]
		public virtual void NavigateTo(object category)
		{
			 InvokerService.InvokeInternal.ExecuteMethod(this, "NavigateTo", category);
		}

		/// <summary>
		/// SupportByVersion Access 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff836254.aspx </remarks>
		/// <param name="objectType">optional NetOffice.AccessApi.Enums.AcDataObjectType ObjectType = -1</param>
		/// <param name="objectName">optional object objectName</param>
		/// <param name="record">optional NetOffice.AccessApi.Enums.AcRecord Record = 2</param>
		/// <param name="whereCondition">optional object whereCondition</param>
		[SupportByVersion("Access", 12,14,15,16)]
		public virtual void SearchForRecord(object objectType, object objectName, object record, object whereCondition)
		{
			 InvokerService.InvokeInternal.ExecuteMethod(this, "SearchForRecord", objectType, objectName, record, whereCondition);
		}

		/// <summary>
		/// SupportByVersion Access 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff836254.aspx </remarks>
		[CustomMethod]
		[SupportByVersion("Access", 12,14,15,16)]
		public virtual void SearchForRecord()
		{
			 InvokerService.InvokeInternal.ExecuteMethod(this, "SearchForRecord");
		}

		/// <summary>
		/// SupportByVersion Access 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff836254.aspx </remarks>
		/// <param name="objectType">optional NetOffice.AccessApi.Enums.AcDataObjectType ObjectType = -1</param>
		[CustomMethod]
		[SupportByVersion("Access", 12,14,15,16)]
		public virtual void SearchForRecord(object objectType)
		{
			 InvokerService.InvokeInternal.ExecuteMethod(this, "SearchForRecord", objectType);
		}

		/// <summary>
		/// SupportByVersion Access 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff836254.aspx </remarks>
		/// <param name="objectType">optional NetOffice.AccessApi.Enums.AcDataObjectType ObjectType = -1</param>
		/// <param name="objectName">optional object objectName</param>
		[CustomMethod]
		[SupportByVersion("Access", 12,14,15,16)]
		public virtual void SearchForRecord(object objectType, object objectName)
		{
			 InvokerService.InvokeInternal.ExecuteMethod(this, "SearchForRecord", objectType, objectName);
		}

		/// <summary>
		/// SupportByVersion Access 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff836254.aspx </remarks>
		/// <param name="objectType">optional NetOffice.AccessApi.Enums.AcDataObjectType ObjectType = -1</param>
		/// <param name="objectName">optional object objectName</param>
		/// <param name="record">optional NetOffice.AccessApi.Enums.AcRecord Record = 2</param>
		[CustomMethod]
		[SupportByVersion("Access", 12,14,15,16)]
		public virtual void SearchForRecord(object objectType, object objectName, object record)
		{
			 InvokerService.InvokeInternal.ExecuteMethod(this, "SearchForRecord", objectType, objectName, record);
		}

		/// <summary>
		/// SupportByVersion Access 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff192301.aspx </remarks>
		/// <param name="controlName">object controlName</param>
		/// <param name="property">optional NetOffice.AccessApi.Enums.AcProperty Property = 0</param>
		/// <param name="value">optional object value</param>
		[SupportByVersion("Access", 12,14,15,16)]
		public virtual void SetProperty(object controlName, object property, object value)
		{
			 InvokerService.InvokeInternal.ExecuteMethod(this, "SetProperty", controlName, property, value);
		}

		/// <summary>
		/// SupportByVersion Access 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff192301.aspx </remarks>
		/// <param name="controlName">object controlName</param>
		[CustomMethod]
		[SupportByVersion("Access", 12,14,15,16)]
		public virtual void SetProperty(object controlName)
		{
			 InvokerService.InvokeInternal.ExecuteMethod(this, "SetProperty", controlName);
		}

		/// <summary>
		/// SupportByVersion Access 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff192301.aspx </remarks>
		/// <param name="controlName">object controlName</param>
		/// <param name="property">optional NetOffice.AccessApi.Enums.AcProperty Property = 0</param>
		[CustomMethod]
		[SupportByVersion("Access", 12,14,15,16)]
		public virtual void SetProperty(object controlName, object property)
		{
			 InvokerService.InvokeInternal.ExecuteMethod(this, "SetProperty", controlName, property);
		}

		/// <summary>
		/// SupportByVersion Access 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff837036.aspx </remarks>
		[SupportByVersion("Access", 12,14,15,16)]
		public virtual void SingleStep()
		{
			 InvokerService.InvokeInternal.ExecuteMethod(this, "SingleStep");
		}

		/// <summary>
		/// SupportByVersion Access 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff191914.aspx </remarks>
		[SupportByVersion("Access", 12,14,15,16)]
		public virtual void ClearMacroError()
		{
			 InvokerService.InvokeInternal.ExecuteMethod(this, "ClearMacroError");
		}

		/// <summary>
		/// SupportByVersion Access 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff821741.aspx </remarks>
		/// <param name="show">object show</param>
		/// <param name="category">optional object category</param>
		[SupportByVersion("Access", 12,14,15,16)]
		public virtual void SetDisplayedCategories(object show, object category)
		{
			 InvokerService.InvokeInternal.ExecuteMethod(this, "SetDisplayedCategories", show, category);
		}

		/// <summary>
		/// SupportByVersion Access 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff821741.aspx </remarks>
		/// <param name="show">object show</param>
		[CustomMethod]
		[SupportByVersion("Access", 12,14,15,16)]
		public virtual void SetDisplayedCategories(object show)
		{
			 InvokerService.InvokeInternal.ExecuteMethod(this, "SetDisplayedCategories", show);
		}

		/// <summary>
		/// SupportByVersion Access 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff195088.aspx </remarks>
		/// <param name="_lock">object lock</param>
		[SupportByVersion("Access", 12,14,15,16)]
		public virtual void LockNavigationPane(object _lock)
		{
			 InvokerService.InvokeInternal.ExecuteMethod(this, "LockNavigationPane", _lock);
		}

		/// <summary>
		/// SupportByVersion Access 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff834375.aspx </remarks>
		/// <param name="savedImportExportName">object savedImportExportName</param>
		[SupportByVersion("Access", 12,14,15,16)]
		public virtual void RunSavedImportExport(object savedImportExportName)
		{
			 InvokerService.InvokeInternal.ExecuteMethod(this, "RunSavedImportExport", savedImportExportName);
		}

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
		public virtual void BrowseTo(NetOffice.AccessApi.Enums.AcBrowseToObjectType objectType, object objectName, object pathtoSubformControl, object whereCondition, object page, object dataMode)
		{
			 InvokerService.InvokeInternal.ExecuteMethod(this, "BrowseTo", new object[]{ objectType, objectName, pathtoSubformControl, whereCondition, page, dataMode });
		}

		/// <summary>
		/// SupportByVersion Access 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff196381.aspx </remarks>
		/// <param name="objectType">NetOffice.AccessApi.Enums.AcBrowseToObjectType objectType</param>
		/// <param name="objectName">object objectName</param>
		[CustomMethod]
		[SupportByVersion("Access", 14,15,16)]
		public virtual void BrowseTo(NetOffice.AccessApi.Enums.AcBrowseToObjectType objectType, object objectName)
		{
			 InvokerService.InvokeInternal.ExecuteMethod(this, "BrowseTo", objectType, objectName);
		}

		/// <summary>
		/// SupportByVersion Access 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff196381.aspx </remarks>
		/// <param name="objectType">NetOffice.AccessApi.Enums.AcBrowseToObjectType objectType</param>
		/// <param name="objectName">object objectName</param>
		/// <param name="pathtoSubformControl">optional object pathtoSubformControl</param>
		[CustomMethod]
		[SupportByVersion("Access", 14,15,16)]
		public virtual void BrowseTo(NetOffice.AccessApi.Enums.AcBrowseToObjectType objectType, object objectName, object pathtoSubformControl)
		{
			 InvokerService.InvokeInternal.ExecuteMethod(this, "BrowseTo", objectType, objectName, pathtoSubformControl);
		}

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
		public virtual void BrowseTo(NetOffice.AccessApi.Enums.AcBrowseToObjectType objectType, object objectName, object pathtoSubformControl, object whereCondition)
		{
			 InvokerService.InvokeInternal.ExecuteMethod(this, "BrowseTo", objectType, objectName, pathtoSubformControl, whereCondition);
		}

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
		public virtual void BrowseTo(NetOffice.AccessApi.Enums.AcBrowseToObjectType objectType, object objectName, object pathtoSubformControl, object whereCondition, object page)
		{
			 InvokerService.InvokeInternal.ExecuteMethod(this, "BrowseTo", new object[]{ objectType, objectName, pathtoSubformControl, whereCondition, page });
		}

		/// <summary>
		/// SupportByVersion Access 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff194182.aspx </remarks>
		/// <param name="name">object name</param>
		/// <param name="expression">object expression</param>
		[SupportByVersion("Access", 14,15,16)]
		public virtual void SetParameter(object name, object expression)
		{
			 InvokerService.InvokeInternal.ExecuteMethod(this, "SetParameter", name, expression);
		}

		/// <summary>
		/// SupportByVersion Access 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff836068.aspx </remarks>
		/// <param name="macroName">object macroName</param>
		[SupportByVersion("Access", 14,15,16)]
		public virtual void RunDataMacro(object macroName)
		{
			 InvokerService.InvokeInternal.ExecuteMethod(this, "RunDataMacro", macroName);
		}

		/// <summary>
		/// SupportByVersion Access 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff844761.aspx </remarks>
		/// <param name="orderBy">object orderBy</param>
		/// <param name="controlName">optional object controlName</param>
		[SupportByVersion("Access", 14,15,16)]
		public virtual void SetOrderBy(object orderBy, object controlName)
		{
			 InvokerService.InvokeInternal.ExecuteMethod(this, "SetOrderBy", orderBy, controlName);
		}

		/// <summary>
		/// SupportByVersion Access 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff844761.aspx </remarks>
		/// <param name="orderBy">object orderBy</param>
		[CustomMethod]
		[SupportByVersion("Access", 14,15,16)]
		public virtual void SetOrderBy(object orderBy)
		{
			 InvokerService.InvokeInternal.ExecuteMethod(this, "SetOrderBy", orderBy);
		}

		/// <summary>
		/// SupportByVersion Access 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff197950.aspx </remarks>
		/// <param name="filterName">optional object filterName</param>
		/// <param name="whereCondition">optional object whereCondition</param>
		/// <param name="controlName">optional object controlName</param>
		[SupportByVersion("Access", 14,15,16)]
		public virtual void SetFilter(object filterName, object whereCondition, object controlName)
		{
			 InvokerService.InvokeInternal.ExecuteMethod(this, "SetFilter", filterName, whereCondition, controlName);
		}

		/// <summary>
		/// SupportByVersion Access 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff197950.aspx </remarks>
		[CustomMethod]
		[SupportByVersion("Access", 14,15,16)]
		public virtual void SetFilter()
		{
			 InvokerService.InvokeInternal.ExecuteMethod(this, "SetFilter");
		}

		/// <summary>
		/// SupportByVersion Access 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff197950.aspx </remarks>
		/// <param name="filterName">optional object filterName</param>
		[CustomMethod]
		[SupportByVersion("Access", 14,15,16)]
		public virtual void SetFilter(object filterName)
		{
			 InvokerService.InvokeInternal.ExecuteMethod(this, "SetFilter", filterName);
		}

		/// <summary>
		/// SupportByVersion Access 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff197950.aspx </remarks>
		/// <param name="filterName">optional object filterName</param>
		/// <param name="whereCondition">optional object whereCondition</param>
		[CustomMethod]
		[SupportByVersion("Access", 14,15,16)]
		public virtual void SetFilter(object filterName, object whereCondition)
		{
			 InvokerService.InvokeInternal.ExecuteMethod(this, "SetFilter", filterName, whereCondition);
		}

		/// <summary>
		/// SupportByVersion Access 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff191907.aspx </remarks>
		[SupportByVersion("Access", 14,15,16)]
		public virtual void RefreshRecord()
		{
			 InvokerService.InvokeInternal.ExecuteMethod(this, "RefreshRecord");
		}

		#endregion

		#pragma warning restore
	}
}

