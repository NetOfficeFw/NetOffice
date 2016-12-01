using System;
using NetRuntimeSystem = System;
using System.Runtime.InteropServices;
using System.Runtime.CompilerServices;
using System.ComponentModel;
using System.Reflection;
using System.Collections.Generic;
using NetOffice;
namespace NetOffice.WordApi
{
	///<summary>
	/// DispatchInterface MailMerge 
	/// SupportByVersion Word, 9,10,11,12,14,15,16
	/// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff836701.aspx
	///</summary>
	[SupportByVersionAttribute("Word", 9,10,11,12,14,15,16)]
	[EntityTypeAttribute(EntityType.IsDispatchInterface)]
	public class MailMerge : COMObject
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
                    _type = typeof(MailMerge);
                    
                return _type;
            }
        }
        
        #endregion
        
		#region Construction

		///<param name="factory">current used factory core</param>
		///<param name="parentObject">object there has created the proxy</param>
        ///<param name="comProxy">inner wrapped COM proxy</param>
		public MailMerge(Core factory, ICOMObject parentObject, object comProxy) : base(factory, parentObject, comProxy)
		{
			
		}

        ///<param name="parentObject">object there has created the proxy</param>
        ///<param name="comProxy">inner wrapped COM proxy</param>
        [EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public MailMerge(ICOMObject parentObject, object comProxy) : base(parentObject, comProxy)
		{
		}
		
		///<param name="factory">current used factory core</param>
		///<param name="parentObject">object there has created the proxy</param>
        ///<param name="comProxy">inner wrapped COM proxy</param>
        ///<param name="comProxyType">Type of inner wrapped COM proxy"</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public MailMerge(Core factory, ICOMObject parentObject, object comProxy, NetRuntimeSystem.Type comProxyType) : base(factory, parentObject, comProxy, comProxyType)
		{

		}

		///<param name="parentObject">object there has created the proxy</param>
        ///<param name="comProxy">inner wrapped COM proxy</param>
        ///<param name="comProxyType">Type of inner wrapped COM proxy"</param>
        [EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public MailMerge(ICOMObject parentObject, object comProxy, NetRuntimeSystem.Type comProxyType) : base(parentObject, comProxy, comProxyType)
		{
		}
		
		///<param name="replacedObject">object to replaced. replacedObject are not usable after this action</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public MailMerge(ICOMObject replacedObject) : base(replacedObject)
		{
		}
		
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public MailMerge() : base()
		{
		}
		
		/// <param name="progId">registered ProgID</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public MailMerge(string progId) : base(progId)
		{
		}
		
		#endregion
		
		#region Properties

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// Get
		/// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff837890.aspx
		/// </summary>
		[SupportByVersionAttribute("Word", 9,10,11,12,14,15,16)]
		public NetOffice.WordApi.Application Application
		{
			get
			{
				object[] paramsArray = null;
				object returnItem = Invoker.PropertyGet(this, "Application", paramsArray);
				NetOffice.WordApi.Application newObject = Factory.CreateKnownObjectFromComProxy(this,returnItem,NetOffice.WordApi.Application.LateBindingApiWrapperType) as NetOffice.WordApi.Application;
				return newObject;
			}
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// Get
		/// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff837179.aspx
		/// </summary>
		[SupportByVersionAttribute("Word", 9,10,11,12,14,15,16)]
		public Int32 Creator
		{
			get
			{
				object[] paramsArray = null;
				object returnItem = Invoker.PropertyGet(this, "Creator", paramsArray);
				return NetRuntimeSystem.Convert.ToInt32(returnItem);
			}
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// Get
		/// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff839335.aspx
		/// Unknown COM Proxy
		/// </summary>
		[SupportByVersionAttribute("Word", 9,10,11,12,14,15,16)]
		public object Parent
		{
			get
			{
				object[] paramsArray = null;
				object returnItem = Invoker.PropertyGet(this, "Parent", paramsArray);
				ICOMObject newObject = Factory.CreateObjectFromComProxy(this,returnItem);
				return newObject;
			}
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// Get/Set
		/// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff197137.aspx
		/// </summary>
		[SupportByVersionAttribute("Word", 9,10,11,12,14,15,16)]
		public NetOffice.WordApi.Enums.WdMailMergeMainDocType MainDocumentType
		{
			get
			{
				object[] paramsArray = null;
				object returnItem = Invoker.PropertyGet(this, "MainDocumentType", paramsArray);
				int intReturnItem = NetRuntimeSystem.Convert.ToInt32(returnItem);
				return (NetOffice.WordApi.Enums.WdMailMergeMainDocType)intReturnItem;
			}
			set
			{
				object[] paramsArray = Invoker.ValidateParamsArray(value);
				Invoker.PropertySet(this, "MainDocumentType", paramsArray);
			}
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// Get
		/// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff840195.aspx
		/// </summary>
		[SupportByVersionAttribute("Word", 9,10,11,12,14,15,16)]
		public NetOffice.WordApi.Enums.WdMailMergeState State
		{
			get
			{
				object[] paramsArray = null;
				object returnItem = Invoker.PropertyGet(this, "State", paramsArray);
				int intReturnItem = NetRuntimeSystem.Convert.ToInt32(returnItem);
				return (NetOffice.WordApi.Enums.WdMailMergeState)intReturnItem;
			}
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// Get/Set
		/// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff845069.aspx
		/// </summary>
		[SupportByVersionAttribute("Word", 9,10,11,12,14,15,16)]
		public NetOffice.WordApi.Enums.WdMailMergeDestination Destination
		{
			get
			{
				object[] paramsArray = null;
				object returnItem = Invoker.PropertyGet(this, "Destination", paramsArray);
				int intReturnItem = NetRuntimeSystem.Convert.ToInt32(returnItem);
				return (NetOffice.WordApi.Enums.WdMailMergeDestination)intReturnItem;
			}
			set
			{
				object[] paramsArray = Invoker.ValidateParamsArray(value);
				Invoker.PropertySet(this, "Destination", paramsArray);
			}
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// Get
		/// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff838518.aspx
		/// </summary>
		[SupportByVersionAttribute("Word", 9,10,11,12,14,15,16)]
		public NetOffice.WordApi.MailMergeDataSource DataSource
		{
			get
			{
				object[] paramsArray = null;
				object returnItem = Invoker.PropertyGet(this, "DataSource", paramsArray);
				NetOffice.WordApi.MailMergeDataSource newObject = Factory.CreateKnownObjectFromComProxy(this,returnItem,NetOffice.WordApi.MailMergeDataSource.LateBindingApiWrapperType) as NetOffice.WordApi.MailMergeDataSource;
				return newObject;
			}
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// Get
		/// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff193129.aspx
		/// </summary>
		[SupportByVersionAttribute("Word", 9,10,11,12,14,15,16)]
		public NetOffice.WordApi.MailMergeFields Fields
		{
			get
			{
				object[] paramsArray = null;
				object returnItem = Invoker.PropertyGet(this, "Fields", paramsArray);
				NetOffice.WordApi.MailMergeFields newObject = Factory.CreateKnownObjectFromComProxy(this,returnItem,NetOffice.WordApi.MailMergeFields.LateBindingApiWrapperType) as NetOffice.WordApi.MailMergeFields;
				return newObject;
			}
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// Get/Set
		/// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff840472.aspx
		/// </summary>
		[SupportByVersionAttribute("Word", 9,10,11,12,14,15,16)]
		public Int32 ViewMailMergeFieldCodes
		{
			get
			{
				object[] paramsArray = null;
				object returnItem = Invoker.PropertyGet(this, "ViewMailMergeFieldCodes", paramsArray);
				return NetRuntimeSystem.Convert.ToInt32(returnItem);
			}
			set
			{
				object[] paramsArray = Invoker.ValidateParamsArray(value);
				Invoker.PropertySet(this, "ViewMailMergeFieldCodes", paramsArray);
			}
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// Get/Set
		/// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff192581.aspx
		/// </summary>
		[SupportByVersionAttribute("Word", 9,10,11,12,14,15,16)]
		public bool SuppressBlankLines
		{
			get
			{
				object[] paramsArray = null;
				object returnItem = Invoker.PropertyGet(this, "SuppressBlankLines", paramsArray);
				return NetRuntimeSystem.Convert.ToBoolean(returnItem);
			}
			set
			{
				object[] paramsArray = Invoker.ValidateParamsArray(value);
				Invoker.PropertySet(this, "SuppressBlankLines", paramsArray);
			}
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// Get/Set
		/// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff841091.aspx
		/// </summary>
		[SupportByVersionAttribute("Word", 9,10,11,12,14,15,16)]
		public bool MailAsAttachment
		{
			get
			{
				object[] paramsArray = null;
				object returnItem = Invoker.PropertyGet(this, "MailAsAttachment", paramsArray);
				return NetRuntimeSystem.Convert.ToBoolean(returnItem);
			}
			set
			{
				object[] paramsArray = Invoker.ValidateParamsArray(value);
				Invoker.PropertySet(this, "MailAsAttachment", paramsArray);
			}
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// Get/Set
		/// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff820768.aspx
		/// </summary>
		[SupportByVersionAttribute("Word", 9,10,11,12,14,15,16)]
		public string MailAddressFieldName
		{
			get
			{
				object[] paramsArray = null;
				object returnItem = Invoker.PropertyGet(this, "MailAddressFieldName", paramsArray);
				return NetRuntimeSystem.Convert.ToString(returnItem);
			}
			set
			{
				object[] paramsArray = Invoker.ValidateParamsArray(value);
				Invoker.PropertySet(this, "MailAddressFieldName", paramsArray);
			}
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// Get/Set
		/// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff820986.aspx
		/// </summary>
		[SupportByVersionAttribute("Word", 9,10,11,12,14,15,16)]
		public string MailSubject
		{
			get
			{
				object[] paramsArray = null;
				object returnItem = Invoker.PropertyGet(this, "MailSubject", paramsArray);
				return NetRuntimeSystem.Convert.ToString(returnItem);
			}
			set
			{
				object[] paramsArray = Invoker.ValidateParamsArray(value);
				Invoker.PropertySet(this, "MailSubject", paramsArray);
			}
		}

		/// <summary>
		/// SupportByVersion Word 10, 11, 12, 14, 15, 16
		/// Get/Set
		/// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff845591.aspx
		/// </summary>
		[SupportByVersionAttribute("Word", 10,11,12,14,15,16)]
		public bool HighlightMergeFields
		{
			get
			{
				object[] paramsArray = null;
				object returnItem = Invoker.PropertyGet(this, "HighlightMergeFields", paramsArray);
				return NetRuntimeSystem.Convert.ToBoolean(returnItem);
			}
			set
			{
				object[] paramsArray = Invoker.ValidateParamsArray(value);
				Invoker.PropertySet(this, "HighlightMergeFields", paramsArray);
			}
		}

		/// <summary>
		/// SupportByVersion Word 10, 11, 12, 14, 15, 16
		/// Get/Set
		/// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff192784.aspx
		/// </summary>
		[SupportByVersionAttribute("Word", 10,11,12,14,15,16)]
		public NetOffice.WordApi.Enums.WdMailMergeMailFormat MailFormat
		{
			get
			{
				object[] paramsArray = null;
				object returnItem = Invoker.PropertyGet(this, "MailFormat", paramsArray);
				int intReturnItem = NetRuntimeSystem.Convert.ToInt32(returnItem);
				return (NetOffice.WordApi.Enums.WdMailMergeMailFormat)intReturnItem;
			}
			set
			{
				object[] paramsArray = Invoker.ValidateParamsArray(value);
				Invoker.PropertySet(this, "MailFormat", paramsArray);
			}
		}

		/// <summary>
		/// SupportByVersion Word 10, 11, 12, 14, 15, 16
		/// Get/Set
		/// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff192539.aspx
		/// </summary>
		[SupportByVersionAttribute("Word", 10,11,12,14,15,16)]
		public string ShowSendToCustom
		{
			get
			{
				object[] paramsArray = null;
				object returnItem = Invoker.PropertyGet(this, "ShowSendToCustom", paramsArray);
				return NetRuntimeSystem.Convert.ToString(returnItem);
			}
			set
			{
				object[] paramsArray = Invoker.ValidateParamsArray(value);
				Invoker.PropertySet(this, "ShowSendToCustom", paramsArray);
			}
		}

		/// <summary>
		/// SupportByVersion Word 10, 11, 12, 14, 15, 16
		/// Get/Set
		/// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff821604.aspx
		/// </summary>
		[SupportByVersionAttribute("Word", 10,11,12,14,15,16)]
		public Int32 WizardState
		{
			get
			{
				object[] paramsArray = null;
				object returnItem = Invoker.PropertyGet(this, "WizardState", paramsArray);
				return NetRuntimeSystem.Convert.ToInt32(returnItem);
			}
			set
			{
				object[] paramsArray = Invoker.ValidateParamsArray(value);
				Invoker.PropertySet(this, "WizardState", paramsArray);
			}
		}

		#endregion

		#region Methods

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff820730.aspx
		/// </summary>
		/// <param name="name">optional object Name</param>
		/// <param name="passwordDocument">optional object PasswordDocument</param>
		/// <param name="writePasswordDocument">optional object WritePasswordDocument</param>
		/// <param name="headerRecord">optional object HeaderRecord</param>
		/// <param name="mSQuery">optional object MSQuery</param>
		/// <param name="sQLStatement">optional object SQLStatement</param>
		/// <param name="sQLStatement1">optional object SQLStatement1</param>
		/// <param name="connection">optional object Connection</param>
		/// <param name="linkToSource">optional object LinkToSource</param>
		[SupportByVersionAttribute("Word", 9,10,11,12,14,15,16)]
		public void CreateDataSource(object name, object passwordDocument, object writePasswordDocument, object headerRecord, object mSQuery, object sQLStatement, object sQLStatement1, object connection, object linkToSource)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(name, passwordDocument, writePasswordDocument, headerRecord, mSQuery, sQLStatement, sQLStatement1, connection, linkToSource);
			Invoker.Method(this, "CreateDataSource", paramsArray);
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff820730.aspx
		/// </summary>
		[CustomMethodAttribute]
		[SupportByVersionAttribute("Word", 9,10,11,12,14,15,16)]
		public void CreateDataSource()
		{
			object[] paramsArray = null;
			Invoker.Method(this, "CreateDataSource", paramsArray);
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff820730.aspx
		/// </summary>
		/// <param name="name">optional object Name</param>
		[CustomMethodAttribute]
		[SupportByVersionAttribute("Word", 9,10,11,12,14,15,16)]
		public void CreateDataSource(object name)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(name);
			Invoker.Method(this, "CreateDataSource", paramsArray);
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff820730.aspx
		/// </summary>
		/// <param name="name">optional object Name</param>
		/// <param name="passwordDocument">optional object PasswordDocument</param>
		[CustomMethodAttribute]
		[SupportByVersionAttribute("Word", 9,10,11,12,14,15,16)]
		public void CreateDataSource(object name, object passwordDocument)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(name, passwordDocument);
			Invoker.Method(this, "CreateDataSource", paramsArray);
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff820730.aspx
		/// </summary>
		/// <param name="name">optional object Name</param>
		/// <param name="passwordDocument">optional object PasswordDocument</param>
		/// <param name="writePasswordDocument">optional object WritePasswordDocument</param>
		[CustomMethodAttribute]
		[SupportByVersionAttribute("Word", 9,10,11,12,14,15,16)]
		public void CreateDataSource(object name, object passwordDocument, object writePasswordDocument)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(name, passwordDocument, writePasswordDocument);
			Invoker.Method(this, "CreateDataSource", paramsArray);
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff820730.aspx
		/// </summary>
		/// <param name="name">optional object Name</param>
		/// <param name="passwordDocument">optional object PasswordDocument</param>
		/// <param name="writePasswordDocument">optional object WritePasswordDocument</param>
		/// <param name="headerRecord">optional object HeaderRecord</param>
		[CustomMethodAttribute]
		[SupportByVersionAttribute("Word", 9,10,11,12,14,15,16)]
		public void CreateDataSource(object name, object passwordDocument, object writePasswordDocument, object headerRecord)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(name, passwordDocument, writePasswordDocument, headerRecord);
			Invoker.Method(this, "CreateDataSource", paramsArray);
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff820730.aspx
		/// </summary>
		/// <param name="name">optional object Name</param>
		/// <param name="passwordDocument">optional object PasswordDocument</param>
		/// <param name="writePasswordDocument">optional object WritePasswordDocument</param>
		/// <param name="headerRecord">optional object HeaderRecord</param>
		/// <param name="mSQuery">optional object MSQuery</param>
		[CustomMethodAttribute]
		[SupportByVersionAttribute("Word", 9,10,11,12,14,15,16)]
		public void CreateDataSource(object name, object passwordDocument, object writePasswordDocument, object headerRecord, object mSQuery)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(name, passwordDocument, writePasswordDocument, headerRecord, mSQuery);
			Invoker.Method(this, "CreateDataSource", paramsArray);
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff820730.aspx
		/// </summary>
		/// <param name="name">optional object Name</param>
		/// <param name="passwordDocument">optional object PasswordDocument</param>
		/// <param name="writePasswordDocument">optional object WritePasswordDocument</param>
		/// <param name="headerRecord">optional object HeaderRecord</param>
		/// <param name="mSQuery">optional object MSQuery</param>
		/// <param name="sQLStatement">optional object SQLStatement</param>
		[CustomMethodAttribute]
		[SupportByVersionAttribute("Word", 9,10,11,12,14,15,16)]
		public void CreateDataSource(object name, object passwordDocument, object writePasswordDocument, object headerRecord, object mSQuery, object sQLStatement)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(name, passwordDocument, writePasswordDocument, headerRecord, mSQuery, sQLStatement);
			Invoker.Method(this, "CreateDataSource", paramsArray);
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff820730.aspx
		/// </summary>
		/// <param name="name">optional object Name</param>
		/// <param name="passwordDocument">optional object PasswordDocument</param>
		/// <param name="writePasswordDocument">optional object WritePasswordDocument</param>
		/// <param name="headerRecord">optional object HeaderRecord</param>
		/// <param name="mSQuery">optional object MSQuery</param>
		/// <param name="sQLStatement">optional object SQLStatement</param>
		/// <param name="sQLStatement1">optional object SQLStatement1</param>
		[CustomMethodAttribute]
		[SupportByVersionAttribute("Word", 9,10,11,12,14,15,16)]
		public void CreateDataSource(object name, object passwordDocument, object writePasswordDocument, object headerRecord, object mSQuery, object sQLStatement, object sQLStatement1)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(name, passwordDocument, writePasswordDocument, headerRecord, mSQuery, sQLStatement, sQLStatement1);
			Invoker.Method(this, "CreateDataSource", paramsArray);
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff820730.aspx
		/// </summary>
		/// <param name="name">optional object Name</param>
		/// <param name="passwordDocument">optional object PasswordDocument</param>
		/// <param name="writePasswordDocument">optional object WritePasswordDocument</param>
		/// <param name="headerRecord">optional object HeaderRecord</param>
		/// <param name="mSQuery">optional object MSQuery</param>
		/// <param name="sQLStatement">optional object SQLStatement</param>
		/// <param name="sQLStatement1">optional object SQLStatement1</param>
		/// <param name="connection">optional object Connection</param>
		[CustomMethodAttribute]
		[SupportByVersionAttribute("Word", 9,10,11,12,14,15,16)]
		public void CreateDataSource(object name, object passwordDocument, object writePasswordDocument, object headerRecord, object mSQuery, object sQLStatement, object sQLStatement1, object connection)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(name, passwordDocument, writePasswordDocument, headerRecord, mSQuery, sQLStatement, sQLStatement1, connection);
			Invoker.Method(this, "CreateDataSource", paramsArray);
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff196953.aspx
		/// </summary>
		/// <param name="name">string Name</param>
		/// <param name="passwordDocument">optional object PasswordDocument</param>
		/// <param name="writePasswordDocument">optional object WritePasswordDocument</param>
		/// <param name="headerRecord">optional object HeaderRecord</param>
		[SupportByVersionAttribute("Word", 9,10,11,12,14,15,16)]
		public void CreateHeaderSource(string name, object passwordDocument, object writePasswordDocument, object headerRecord)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(name, passwordDocument, writePasswordDocument, headerRecord);
			Invoker.Method(this, "CreateHeaderSource", paramsArray);
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff196953.aspx
		/// </summary>
		/// <param name="name">string Name</param>
		[CustomMethodAttribute]
		[SupportByVersionAttribute("Word", 9,10,11,12,14,15,16)]
		public void CreateHeaderSource(string name)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(name);
			Invoker.Method(this, "CreateHeaderSource", paramsArray);
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff196953.aspx
		/// </summary>
		/// <param name="name">string Name</param>
		/// <param name="passwordDocument">optional object PasswordDocument</param>
		[CustomMethodAttribute]
		[SupportByVersionAttribute("Word", 9,10,11,12,14,15,16)]
		public void CreateHeaderSource(string name, object passwordDocument)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(name, passwordDocument);
			Invoker.Method(this, "CreateHeaderSource", paramsArray);
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff196953.aspx
		/// </summary>
		/// <param name="name">string Name</param>
		/// <param name="passwordDocument">optional object PasswordDocument</param>
		/// <param name="writePasswordDocument">optional object WritePasswordDocument</param>
		[CustomMethodAttribute]
		[SupportByVersionAttribute("Word", 9,10,11,12,14,15,16)]
		public void CreateHeaderSource(string name, object passwordDocument, object writePasswordDocument)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(name, passwordDocument, writePasswordDocument);
			Invoker.Method(this, "CreateHeaderSource", paramsArray);
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff841005.aspx
		/// </summary>
		/// <param name="name">string Name</param>
		/// <param name="format">optional object Format</param>
		/// <param name="confirmConversions">optional object ConfirmConversions</param>
		/// <param name="readOnly">optional object ReadOnly</param>
		/// <param name="linkToSource">optional object LinkToSource</param>
		/// <param name="addToRecentFiles">optional object AddToRecentFiles</param>
		/// <param name="passwordDocument">optional object PasswordDocument</param>
		/// <param name="passwordTemplate">optional object PasswordTemplate</param>
		/// <param name="revert">optional object Revert</param>
		/// <param name="writePasswordDocument">optional object WritePasswordDocument</param>
		/// <param name="writePasswordTemplate">optional object WritePasswordTemplate</param>
		/// <param name="connection">optional object Connection</param>
		/// <param name="sQLStatement">optional object SQLStatement</param>
		/// <param name="sQLStatement1">optional object SQLStatement1</param>
		[SupportByVersionAttribute("Word", 9,10,11,12,14,15,16)]
		public void OpenDataSource(string name, object format, object confirmConversions, object readOnly, object linkToSource, object addToRecentFiles, object passwordDocument, object passwordTemplate, object revert, object writePasswordDocument, object writePasswordTemplate, object connection, object sQLStatement, object sQLStatement1)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(name, format, confirmConversions, readOnly, linkToSource, addToRecentFiles, passwordDocument, passwordTemplate, revert, writePasswordDocument, writePasswordTemplate, connection, sQLStatement, sQLStatement1);
			Invoker.Method(this, "OpenDataSource", paramsArray);
		}

		/// <summary>
		/// SupportByVersion Word 10, 11, 12, 14, 15, 16
		/// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff841005.aspx
		/// </summary>
		/// <param name="name">string Name</param>
		/// <param name="format">optional object Format</param>
		/// <param name="confirmConversions">optional object ConfirmConversions</param>
		/// <param name="readOnly">optional object ReadOnly</param>
		/// <param name="linkToSource">optional object LinkToSource</param>
		/// <param name="addToRecentFiles">optional object AddToRecentFiles</param>
		/// <param name="passwordDocument">optional object PasswordDocument</param>
		/// <param name="passwordTemplate">optional object PasswordTemplate</param>
		/// <param name="revert">optional object Revert</param>
		/// <param name="writePasswordDocument">optional object WritePasswordDocument</param>
		/// <param name="writePasswordTemplate">optional object WritePasswordTemplate</param>
		/// <param name="connection">optional object Connection</param>
		/// <param name="sQLStatement">optional object SQLStatement</param>
		/// <param name="sQLStatement1">optional object SQLStatement1</param>
		/// <param name="openExclusive">optional object OpenExclusive</param>
		/// <param name="subType">optional object SubType</param>
		[SupportByVersionAttribute("Word", 10,11,12,14,15,16)]
		public void OpenDataSource(string name, object format, object confirmConversions, object readOnly, object linkToSource, object addToRecentFiles, object passwordDocument, object passwordTemplate, object revert, object writePasswordDocument, object writePasswordTemplate, object connection, object sQLStatement, object sQLStatement1, object openExclusive, object subType)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(name, format, confirmConversions, readOnly, linkToSource, addToRecentFiles, passwordDocument, passwordTemplate, revert, writePasswordDocument, writePasswordTemplate, connection, sQLStatement, sQLStatement1, openExclusive, subType);
			Invoker.Method(this, "OpenDataSource", paramsArray);
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff841005.aspx
		/// </summary>
		/// <param name="name">string Name</param>
		[CustomMethodAttribute]
		[SupportByVersionAttribute("Word", 9,10,11,12,14,15,16)]
		public void OpenDataSource(string name)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(name);
			Invoker.Method(this, "OpenDataSource", paramsArray);
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff841005.aspx
		/// </summary>
		/// <param name="name">string Name</param>
		/// <param name="format">optional object Format</param>
		[CustomMethodAttribute]
		[SupportByVersionAttribute("Word", 9,10,11,12,14,15,16)]
		public void OpenDataSource(string name, object format)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(name, format);
			Invoker.Method(this, "OpenDataSource", paramsArray);
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff841005.aspx
		/// </summary>
		/// <param name="name">string Name</param>
		/// <param name="format">optional object Format</param>
		/// <param name="confirmConversions">optional object ConfirmConversions</param>
		[CustomMethodAttribute]
		[SupportByVersionAttribute("Word", 9,10,11,12,14,15,16)]
		public void OpenDataSource(string name, object format, object confirmConversions)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(name, format, confirmConversions);
			Invoker.Method(this, "OpenDataSource", paramsArray);
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff841005.aspx
		/// </summary>
		/// <param name="name">string Name</param>
		/// <param name="format">optional object Format</param>
		/// <param name="confirmConversions">optional object ConfirmConversions</param>
		/// <param name="readOnly">optional object ReadOnly</param>
		[CustomMethodAttribute]
		[SupportByVersionAttribute("Word", 9,10,11,12,14,15,16)]
		public void OpenDataSource(string name, object format, object confirmConversions, object readOnly)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(name, format, confirmConversions, readOnly);
			Invoker.Method(this, "OpenDataSource", paramsArray);
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff841005.aspx
		/// </summary>
		/// <param name="name">string Name</param>
		/// <param name="format">optional object Format</param>
		/// <param name="confirmConversions">optional object ConfirmConversions</param>
		/// <param name="readOnly">optional object ReadOnly</param>
		/// <param name="linkToSource">optional object LinkToSource</param>
		[CustomMethodAttribute]
		[SupportByVersionAttribute("Word", 9,10,11,12,14,15,16)]
		public void OpenDataSource(string name, object format, object confirmConversions, object readOnly, object linkToSource)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(name, format, confirmConversions, readOnly, linkToSource);
			Invoker.Method(this, "OpenDataSource", paramsArray);
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff841005.aspx
		/// </summary>
		/// <param name="name">string Name</param>
		/// <param name="format">optional object Format</param>
		/// <param name="confirmConversions">optional object ConfirmConversions</param>
		/// <param name="readOnly">optional object ReadOnly</param>
		/// <param name="linkToSource">optional object LinkToSource</param>
		/// <param name="addToRecentFiles">optional object AddToRecentFiles</param>
		[CustomMethodAttribute]
		[SupportByVersionAttribute("Word", 9,10,11,12,14,15,16)]
		public void OpenDataSource(string name, object format, object confirmConversions, object readOnly, object linkToSource, object addToRecentFiles)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(name, format, confirmConversions, readOnly, linkToSource, addToRecentFiles);
			Invoker.Method(this, "OpenDataSource", paramsArray);
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff841005.aspx
		/// </summary>
		/// <param name="name">string Name</param>
		/// <param name="format">optional object Format</param>
		/// <param name="confirmConversions">optional object ConfirmConversions</param>
		/// <param name="readOnly">optional object ReadOnly</param>
		/// <param name="linkToSource">optional object LinkToSource</param>
		/// <param name="addToRecentFiles">optional object AddToRecentFiles</param>
		/// <param name="passwordDocument">optional object PasswordDocument</param>
		[CustomMethodAttribute]
		[SupportByVersionAttribute("Word", 9,10,11,12,14,15,16)]
		public void OpenDataSource(string name, object format, object confirmConversions, object readOnly, object linkToSource, object addToRecentFiles, object passwordDocument)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(name, format, confirmConversions, readOnly, linkToSource, addToRecentFiles, passwordDocument);
			Invoker.Method(this, "OpenDataSource", paramsArray);
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff841005.aspx
		/// </summary>
		/// <param name="name">string Name</param>
		/// <param name="format">optional object Format</param>
		/// <param name="confirmConversions">optional object ConfirmConversions</param>
		/// <param name="readOnly">optional object ReadOnly</param>
		/// <param name="linkToSource">optional object LinkToSource</param>
		/// <param name="addToRecentFiles">optional object AddToRecentFiles</param>
		/// <param name="passwordDocument">optional object PasswordDocument</param>
		/// <param name="passwordTemplate">optional object PasswordTemplate</param>
		[CustomMethodAttribute]
		[SupportByVersionAttribute("Word", 9,10,11,12,14,15,16)]
		public void OpenDataSource(string name, object format, object confirmConversions, object readOnly, object linkToSource, object addToRecentFiles, object passwordDocument, object passwordTemplate)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(name, format, confirmConversions, readOnly, linkToSource, addToRecentFiles, passwordDocument, passwordTemplate);
			Invoker.Method(this, "OpenDataSource", paramsArray);
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff841005.aspx
		/// </summary>
		/// <param name="name">string Name</param>
		/// <param name="format">optional object Format</param>
		/// <param name="confirmConversions">optional object ConfirmConversions</param>
		/// <param name="readOnly">optional object ReadOnly</param>
		/// <param name="linkToSource">optional object LinkToSource</param>
		/// <param name="addToRecentFiles">optional object AddToRecentFiles</param>
		/// <param name="passwordDocument">optional object PasswordDocument</param>
		/// <param name="passwordTemplate">optional object PasswordTemplate</param>
		/// <param name="revert">optional object Revert</param>
		[CustomMethodAttribute]
		[SupportByVersionAttribute("Word", 9,10,11,12,14,15,16)]
		public void OpenDataSource(string name, object format, object confirmConversions, object readOnly, object linkToSource, object addToRecentFiles, object passwordDocument, object passwordTemplate, object revert)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(name, format, confirmConversions, readOnly, linkToSource, addToRecentFiles, passwordDocument, passwordTemplate, revert);
			Invoker.Method(this, "OpenDataSource", paramsArray);
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff841005.aspx
		/// </summary>
		/// <param name="name">string Name</param>
		/// <param name="format">optional object Format</param>
		/// <param name="confirmConversions">optional object ConfirmConversions</param>
		/// <param name="readOnly">optional object ReadOnly</param>
		/// <param name="linkToSource">optional object LinkToSource</param>
		/// <param name="addToRecentFiles">optional object AddToRecentFiles</param>
		/// <param name="passwordDocument">optional object PasswordDocument</param>
		/// <param name="passwordTemplate">optional object PasswordTemplate</param>
		/// <param name="revert">optional object Revert</param>
		/// <param name="writePasswordDocument">optional object WritePasswordDocument</param>
		[CustomMethodAttribute]
		[SupportByVersionAttribute("Word", 9,10,11,12,14,15,16)]
		public void OpenDataSource(string name, object format, object confirmConversions, object readOnly, object linkToSource, object addToRecentFiles, object passwordDocument, object passwordTemplate, object revert, object writePasswordDocument)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(name, format, confirmConversions, readOnly, linkToSource, addToRecentFiles, passwordDocument, passwordTemplate, revert, writePasswordDocument);
			Invoker.Method(this, "OpenDataSource", paramsArray);
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff841005.aspx
		/// </summary>
		/// <param name="name">string Name</param>
		/// <param name="format">optional object Format</param>
		/// <param name="confirmConversions">optional object ConfirmConversions</param>
		/// <param name="readOnly">optional object ReadOnly</param>
		/// <param name="linkToSource">optional object LinkToSource</param>
		/// <param name="addToRecentFiles">optional object AddToRecentFiles</param>
		/// <param name="passwordDocument">optional object PasswordDocument</param>
		/// <param name="passwordTemplate">optional object PasswordTemplate</param>
		/// <param name="revert">optional object Revert</param>
		/// <param name="writePasswordDocument">optional object WritePasswordDocument</param>
		/// <param name="writePasswordTemplate">optional object WritePasswordTemplate</param>
		[CustomMethodAttribute]
		[SupportByVersionAttribute("Word", 9,10,11,12,14,15,16)]
		public void OpenDataSource(string name, object format, object confirmConversions, object readOnly, object linkToSource, object addToRecentFiles, object passwordDocument, object passwordTemplate, object revert, object writePasswordDocument, object writePasswordTemplate)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(name, format, confirmConversions, readOnly, linkToSource, addToRecentFiles, passwordDocument, passwordTemplate, revert, writePasswordDocument, writePasswordTemplate);
			Invoker.Method(this, "OpenDataSource", paramsArray);
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff841005.aspx
		/// </summary>
		/// <param name="name">string Name</param>
		/// <param name="format">optional object Format</param>
		/// <param name="confirmConversions">optional object ConfirmConversions</param>
		/// <param name="readOnly">optional object ReadOnly</param>
		/// <param name="linkToSource">optional object LinkToSource</param>
		/// <param name="addToRecentFiles">optional object AddToRecentFiles</param>
		/// <param name="passwordDocument">optional object PasswordDocument</param>
		/// <param name="passwordTemplate">optional object PasswordTemplate</param>
		/// <param name="revert">optional object Revert</param>
		/// <param name="writePasswordDocument">optional object WritePasswordDocument</param>
		/// <param name="writePasswordTemplate">optional object WritePasswordTemplate</param>
		/// <param name="connection">optional object Connection</param>
		[CustomMethodAttribute]
		[SupportByVersionAttribute("Word", 9,10,11,12,14,15,16)]
		public void OpenDataSource(string name, object format, object confirmConversions, object readOnly, object linkToSource, object addToRecentFiles, object passwordDocument, object passwordTemplate, object revert, object writePasswordDocument, object writePasswordTemplate, object connection)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(name, format, confirmConversions, readOnly, linkToSource, addToRecentFiles, passwordDocument, passwordTemplate, revert, writePasswordDocument, writePasswordTemplate, connection);
			Invoker.Method(this, "OpenDataSource", paramsArray);
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff841005.aspx
		/// </summary>
		/// <param name="name">string Name</param>
		/// <param name="format">optional object Format</param>
		/// <param name="confirmConversions">optional object ConfirmConversions</param>
		/// <param name="readOnly">optional object ReadOnly</param>
		/// <param name="linkToSource">optional object LinkToSource</param>
		/// <param name="addToRecentFiles">optional object AddToRecentFiles</param>
		/// <param name="passwordDocument">optional object PasswordDocument</param>
		/// <param name="passwordTemplate">optional object PasswordTemplate</param>
		/// <param name="revert">optional object Revert</param>
		/// <param name="writePasswordDocument">optional object WritePasswordDocument</param>
		/// <param name="writePasswordTemplate">optional object WritePasswordTemplate</param>
		/// <param name="connection">optional object Connection</param>
		/// <param name="sQLStatement">optional object SQLStatement</param>
		[CustomMethodAttribute]
		[SupportByVersionAttribute("Word", 9,10,11,12,14,15,16)]
		public void OpenDataSource(string name, object format, object confirmConversions, object readOnly, object linkToSource, object addToRecentFiles, object passwordDocument, object passwordTemplate, object revert, object writePasswordDocument, object writePasswordTemplate, object connection, object sQLStatement)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(name, format, confirmConversions, readOnly, linkToSource, addToRecentFiles, passwordDocument, passwordTemplate, revert, writePasswordDocument, writePasswordTemplate, connection, sQLStatement);
			Invoker.Method(this, "OpenDataSource", paramsArray);
		}

		/// <summary>
		/// SupportByVersion Word 10, 11, 12, 14, 15, 16
		/// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff841005.aspx
		/// </summary>
		/// <param name="name">string Name</param>
		/// <param name="format">optional object Format</param>
		/// <param name="confirmConversions">optional object ConfirmConversions</param>
		/// <param name="readOnly">optional object ReadOnly</param>
		/// <param name="linkToSource">optional object LinkToSource</param>
		/// <param name="addToRecentFiles">optional object AddToRecentFiles</param>
		/// <param name="passwordDocument">optional object PasswordDocument</param>
		/// <param name="passwordTemplate">optional object PasswordTemplate</param>
		/// <param name="revert">optional object Revert</param>
		/// <param name="writePasswordDocument">optional object WritePasswordDocument</param>
		/// <param name="writePasswordTemplate">optional object WritePasswordTemplate</param>
		/// <param name="connection">optional object Connection</param>
		/// <param name="sQLStatement">optional object SQLStatement</param>
		/// <param name="sQLStatement1">optional object SQLStatement1</param>
		/// <param name="openExclusive">optional object OpenExclusive</param>
		[CustomMethodAttribute]
		[SupportByVersionAttribute("Word", 10,11,12,14,15,16)]
		public void OpenDataSource(string name, object format, object confirmConversions, object readOnly, object linkToSource, object addToRecentFiles, object passwordDocument, object passwordTemplate, object revert, object writePasswordDocument, object writePasswordTemplate, object connection, object sQLStatement, object sQLStatement1, object openExclusive)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(name, format, confirmConversions, readOnly, linkToSource, addToRecentFiles, passwordDocument, passwordTemplate, revert, writePasswordDocument, writePasswordTemplate, connection, sQLStatement, sQLStatement1, openExclusive);
			Invoker.Method(this, "OpenDataSource", paramsArray);
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff845427.aspx
		/// </summary>
		/// <param name="name">string Name</param>
		/// <param name="format">optional object Format</param>
		/// <param name="confirmConversions">optional object ConfirmConversions</param>
		/// <param name="readOnly">optional object ReadOnly</param>
		/// <param name="addToRecentFiles">optional object AddToRecentFiles</param>
		/// <param name="passwordDocument">optional object PasswordDocument</param>
		/// <param name="passwordTemplate">optional object PasswordTemplate</param>
		/// <param name="revert">optional object Revert</param>
		/// <param name="writePasswordDocument">optional object WritePasswordDocument</param>
		/// <param name="writePasswordTemplate">optional object WritePasswordTemplate</param>
		[SupportByVersionAttribute("Word", 9,10,11,12,14,15,16)]
		public void OpenHeaderSource(string name, object format, object confirmConversions, object readOnly, object addToRecentFiles, object passwordDocument, object passwordTemplate, object revert, object writePasswordDocument, object writePasswordTemplate)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(name, format, confirmConversions, readOnly, addToRecentFiles, passwordDocument, passwordTemplate, revert, writePasswordDocument, writePasswordTemplate);
			Invoker.Method(this, "OpenHeaderSource", paramsArray);
		}

		/// <summary>
		/// SupportByVersion Word 10, 11, 12, 14, 15, 16
		/// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff845427.aspx
		/// </summary>
		/// <param name="name">string Name</param>
		/// <param name="format">optional object Format</param>
		/// <param name="confirmConversions">optional object ConfirmConversions</param>
		/// <param name="readOnly">optional object ReadOnly</param>
		/// <param name="addToRecentFiles">optional object AddToRecentFiles</param>
		/// <param name="passwordDocument">optional object PasswordDocument</param>
		/// <param name="passwordTemplate">optional object PasswordTemplate</param>
		/// <param name="revert">optional object Revert</param>
		/// <param name="writePasswordDocument">optional object WritePasswordDocument</param>
		/// <param name="writePasswordTemplate">optional object WritePasswordTemplate</param>
		/// <param name="openExclusive">optional object OpenExclusive</param>
		[SupportByVersionAttribute("Word", 10,11,12,14,15,16)]
		public void OpenHeaderSource(string name, object format, object confirmConversions, object readOnly, object addToRecentFiles, object passwordDocument, object passwordTemplate, object revert, object writePasswordDocument, object writePasswordTemplate, object openExclusive)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(name, format, confirmConversions, readOnly, addToRecentFiles, passwordDocument, passwordTemplate, revert, writePasswordDocument, writePasswordTemplate, openExclusive);
			Invoker.Method(this, "OpenHeaderSource", paramsArray);
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff845427.aspx
		/// </summary>
		/// <param name="name">string Name</param>
		[CustomMethodAttribute]
		[SupportByVersionAttribute("Word", 9,10,11,12,14,15,16)]
		public void OpenHeaderSource(string name)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(name);
			Invoker.Method(this, "OpenHeaderSource", paramsArray);
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff845427.aspx
		/// </summary>
		/// <param name="name">string Name</param>
		/// <param name="format">optional object Format</param>
		[CustomMethodAttribute]
		[SupportByVersionAttribute("Word", 9,10,11,12,14,15,16)]
		public void OpenHeaderSource(string name, object format)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(name, format);
			Invoker.Method(this, "OpenHeaderSource", paramsArray);
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff845427.aspx
		/// </summary>
		/// <param name="name">string Name</param>
		/// <param name="format">optional object Format</param>
		/// <param name="confirmConversions">optional object ConfirmConversions</param>
		[CustomMethodAttribute]
		[SupportByVersionAttribute("Word", 9,10,11,12,14,15,16)]
		public void OpenHeaderSource(string name, object format, object confirmConversions)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(name, format, confirmConversions);
			Invoker.Method(this, "OpenHeaderSource", paramsArray);
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff845427.aspx
		/// </summary>
		/// <param name="name">string Name</param>
		/// <param name="format">optional object Format</param>
		/// <param name="confirmConversions">optional object ConfirmConversions</param>
		/// <param name="readOnly">optional object ReadOnly</param>
		[CustomMethodAttribute]
		[SupportByVersionAttribute("Word", 9,10,11,12,14,15,16)]
		public void OpenHeaderSource(string name, object format, object confirmConversions, object readOnly)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(name, format, confirmConversions, readOnly);
			Invoker.Method(this, "OpenHeaderSource", paramsArray);
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff845427.aspx
		/// </summary>
		/// <param name="name">string Name</param>
		/// <param name="format">optional object Format</param>
		/// <param name="confirmConversions">optional object ConfirmConversions</param>
		/// <param name="readOnly">optional object ReadOnly</param>
		/// <param name="addToRecentFiles">optional object AddToRecentFiles</param>
		[CustomMethodAttribute]
		[SupportByVersionAttribute("Word", 9,10,11,12,14,15,16)]
		public void OpenHeaderSource(string name, object format, object confirmConversions, object readOnly, object addToRecentFiles)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(name, format, confirmConversions, readOnly, addToRecentFiles);
			Invoker.Method(this, "OpenHeaderSource", paramsArray);
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff845427.aspx
		/// </summary>
		/// <param name="name">string Name</param>
		/// <param name="format">optional object Format</param>
		/// <param name="confirmConversions">optional object ConfirmConversions</param>
		/// <param name="readOnly">optional object ReadOnly</param>
		/// <param name="addToRecentFiles">optional object AddToRecentFiles</param>
		/// <param name="passwordDocument">optional object PasswordDocument</param>
		[CustomMethodAttribute]
		[SupportByVersionAttribute("Word", 9,10,11,12,14,15,16)]
		public void OpenHeaderSource(string name, object format, object confirmConversions, object readOnly, object addToRecentFiles, object passwordDocument)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(name, format, confirmConversions, readOnly, addToRecentFiles, passwordDocument);
			Invoker.Method(this, "OpenHeaderSource", paramsArray);
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff845427.aspx
		/// </summary>
		/// <param name="name">string Name</param>
		/// <param name="format">optional object Format</param>
		/// <param name="confirmConversions">optional object ConfirmConversions</param>
		/// <param name="readOnly">optional object ReadOnly</param>
		/// <param name="addToRecentFiles">optional object AddToRecentFiles</param>
		/// <param name="passwordDocument">optional object PasswordDocument</param>
		/// <param name="passwordTemplate">optional object PasswordTemplate</param>
		[CustomMethodAttribute]
		[SupportByVersionAttribute("Word", 9,10,11,12,14,15,16)]
		public void OpenHeaderSource(string name, object format, object confirmConversions, object readOnly, object addToRecentFiles, object passwordDocument, object passwordTemplate)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(name, format, confirmConversions, readOnly, addToRecentFiles, passwordDocument, passwordTemplate);
			Invoker.Method(this, "OpenHeaderSource", paramsArray);
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff845427.aspx
		/// </summary>
		/// <param name="name">string Name</param>
		/// <param name="format">optional object Format</param>
		/// <param name="confirmConversions">optional object ConfirmConversions</param>
		/// <param name="readOnly">optional object ReadOnly</param>
		/// <param name="addToRecentFiles">optional object AddToRecentFiles</param>
		/// <param name="passwordDocument">optional object PasswordDocument</param>
		/// <param name="passwordTemplate">optional object PasswordTemplate</param>
		/// <param name="revert">optional object Revert</param>
		[CustomMethodAttribute]
		[SupportByVersionAttribute("Word", 9,10,11,12,14,15,16)]
		public void OpenHeaderSource(string name, object format, object confirmConversions, object readOnly, object addToRecentFiles, object passwordDocument, object passwordTemplate, object revert)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(name, format, confirmConversions, readOnly, addToRecentFiles, passwordDocument, passwordTemplate, revert);
			Invoker.Method(this, "OpenHeaderSource", paramsArray);
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff845427.aspx
		/// </summary>
		/// <param name="name">string Name</param>
		/// <param name="format">optional object Format</param>
		/// <param name="confirmConversions">optional object ConfirmConversions</param>
		/// <param name="readOnly">optional object ReadOnly</param>
		/// <param name="addToRecentFiles">optional object AddToRecentFiles</param>
		/// <param name="passwordDocument">optional object PasswordDocument</param>
		/// <param name="passwordTemplate">optional object PasswordTemplate</param>
		/// <param name="revert">optional object Revert</param>
		/// <param name="writePasswordDocument">optional object WritePasswordDocument</param>
		[CustomMethodAttribute]
		[SupportByVersionAttribute("Word", 9,10,11,12,14,15,16)]
		public void OpenHeaderSource(string name, object format, object confirmConversions, object readOnly, object addToRecentFiles, object passwordDocument, object passwordTemplate, object revert, object writePasswordDocument)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(name, format, confirmConversions, readOnly, addToRecentFiles, passwordDocument, passwordTemplate, revert, writePasswordDocument);
			Invoker.Method(this, "OpenHeaderSource", paramsArray);
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff841097.aspx
		/// </summary>
		/// <param name="pause">optional object Pause</param>
		[SupportByVersionAttribute("Word", 9,10,11,12,14,15,16)]
		public void Execute(object pause)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(pause);
			Invoker.Method(this, "Execute", paramsArray);
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff841097.aspx
		/// </summary>
		[CustomMethodAttribute]
		[SupportByVersionAttribute("Word", 9,10,11,12,14,15,16)]
		public void Execute()
		{
			object[] paramsArray = null;
			Invoker.Method(this, "Execute", paramsArray);
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff835814.aspx
		/// </summary>
		[SupportByVersionAttribute("Word", 9,10,11,12,14,15,16)]
		public void Check()
		{
			object[] paramsArray = null;
			Invoker.Method(this, "Check", paramsArray);
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff192805.aspx
		/// </summary>
		[SupportByVersionAttribute("Word", 9,10,11,12,14,15,16)]
		public void EditDataSource()
		{
			object[] paramsArray = null;
			Invoker.Method(this, "EditDataSource", paramsArray);
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff838561.aspx
		/// </summary>
		[SupportByVersionAttribute("Word", 9,10,11,12,14,15,16)]
		public void EditHeaderSource()
		{
			object[] paramsArray = null;
			Invoker.Method(this, "EditHeaderSource", paramsArray);
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff845149.aspx
		/// </summary>
		[SupportByVersionAttribute("Word", 9,10,11,12,14,15,16)]
		public void EditMainDocument()
		{
			object[] paramsArray = null;
			Invoker.Method(this, "EditMainDocument", paramsArray);
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// 
		/// </summary>
		/// <param name="type">string Type</param>
		[SupportByVersionAttribute("Word", 9,10,11,12,14,15,16)]
		public void UseAddressBook(string type)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(type);
			Invoker.Method(this, "UseAddressBook", paramsArray);
		}

		/// <summary>
		/// SupportByVersion Word 10, 11, 12, 14, 15, 16
		/// 
		/// </summary>
		/// <param name="name">string Name</param>
		/// <param name="format">optional object Format</param>
		/// <param name="confirmConversions">optional object ConfirmConversions</param>
		/// <param name="readOnly">optional object ReadOnly</param>
		/// <param name="linkToSource">optional object LinkToSource</param>
		/// <param name="addToRecentFiles">optional object AddToRecentFiles</param>
		/// <param name="passwordDocument">optional object PasswordDocument</param>
		/// <param name="passwordTemplate">optional object PasswordTemplate</param>
		/// <param name="revert">optional object Revert</param>
		/// <param name="writePasswordDocument">optional object WritePasswordDocument</param>
		/// <param name="writePasswordTemplate">optional object WritePasswordTemplate</param>
		/// <param name="connection">optional object Connection</param>
		/// <param name="sQLStatement">optional object SQLStatement</param>
		/// <param name="sQLStatement1">optional object SQLStatement1</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		[SupportByVersionAttribute("Word", 10,11,12,14,15,16)]
		public void OpenDataSource2000(string name, object format, object confirmConversions, object readOnly, object linkToSource, object addToRecentFiles, object passwordDocument, object passwordTemplate, object revert, object writePasswordDocument, object writePasswordTemplate, object connection, object sQLStatement, object sQLStatement1)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(name, format, confirmConversions, readOnly, linkToSource, addToRecentFiles, passwordDocument, passwordTemplate, revert, writePasswordDocument, writePasswordTemplate, connection, sQLStatement, sQLStatement1);
			Invoker.Method(this, "OpenDataSource2000", paramsArray);
		}

		/// <summary>
		/// SupportByVersion Word 10, 11, 12, 14, 15, 16
		/// 
		/// </summary>
		/// <param name="name">string Name</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		[CustomMethodAttribute]
		[SupportByVersionAttribute("Word", 10,11,12,14,15,16)]
		public void OpenDataSource2000(string name)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(name);
			Invoker.Method(this, "OpenDataSource2000", paramsArray);
		}

		/// <summary>
		/// SupportByVersion Word 10, 11, 12, 14, 15, 16
		/// 
		/// </summary>
		/// <param name="name">string Name</param>
		/// <param name="format">optional object Format</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		[CustomMethodAttribute]
		[SupportByVersionAttribute("Word", 10,11,12,14,15,16)]
		public void OpenDataSource2000(string name, object format)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(name, format);
			Invoker.Method(this, "OpenDataSource2000", paramsArray);
		}

		/// <summary>
		/// SupportByVersion Word 10, 11, 12, 14, 15, 16
		/// 
		/// </summary>
		/// <param name="name">string Name</param>
		/// <param name="format">optional object Format</param>
		/// <param name="confirmConversions">optional object ConfirmConversions</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		[CustomMethodAttribute]
		[SupportByVersionAttribute("Word", 10,11,12,14,15,16)]
		public void OpenDataSource2000(string name, object format, object confirmConversions)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(name, format, confirmConversions);
			Invoker.Method(this, "OpenDataSource2000", paramsArray);
		}

		/// <summary>
		/// SupportByVersion Word 10, 11, 12, 14, 15, 16
		/// 
		/// </summary>
		/// <param name="name">string Name</param>
		/// <param name="format">optional object Format</param>
		/// <param name="confirmConversions">optional object ConfirmConversions</param>
		/// <param name="readOnly">optional object ReadOnly</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		[CustomMethodAttribute]
		[SupportByVersionAttribute("Word", 10,11,12,14,15,16)]
		public void OpenDataSource2000(string name, object format, object confirmConversions, object readOnly)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(name, format, confirmConversions, readOnly);
			Invoker.Method(this, "OpenDataSource2000", paramsArray);
		}

		/// <summary>
		/// SupportByVersion Word 10, 11, 12, 14, 15, 16
		/// 
		/// </summary>
		/// <param name="name">string Name</param>
		/// <param name="format">optional object Format</param>
		/// <param name="confirmConversions">optional object ConfirmConversions</param>
		/// <param name="readOnly">optional object ReadOnly</param>
		/// <param name="linkToSource">optional object LinkToSource</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		[CustomMethodAttribute]
		[SupportByVersionAttribute("Word", 10,11,12,14,15,16)]
		public void OpenDataSource2000(string name, object format, object confirmConversions, object readOnly, object linkToSource)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(name, format, confirmConversions, readOnly, linkToSource);
			Invoker.Method(this, "OpenDataSource2000", paramsArray);
		}

		/// <summary>
		/// SupportByVersion Word 10, 11, 12, 14, 15, 16
		/// 
		/// </summary>
		/// <param name="name">string Name</param>
		/// <param name="format">optional object Format</param>
		/// <param name="confirmConversions">optional object ConfirmConversions</param>
		/// <param name="readOnly">optional object ReadOnly</param>
		/// <param name="linkToSource">optional object LinkToSource</param>
		/// <param name="addToRecentFiles">optional object AddToRecentFiles</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		[CustomMethodAttribute]
		[SupportByVersionAttribute("Word", 10,11,12,14,15,16)]
		public void OpenDataSource2000(string name, object format, object confirmConversions, object readOnly, object linkToSource, object addToRecentFiles)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(name, format, confirmConversions, readOnly, linkToSource, addToRecentFiles);
			Invoker.Method(this, "OpenDataSource2000", paramsArray);
		}

		/// <summary>
		/// SupportByVersion Word 10, 11, 12, 14, 15, 16
		/// 
		/// </summary>
		/// <param name="name">string Name</param>
		/// <param name="format">optional object Format</param>
		/// <param name="confirmConversions">optional object ConfirmConversions</param>
		/// <param name="readOnly">optional object ReadOnly</param>
		/// <param name="linkToSource">optional object LinkToSource</param>
		/// <param name="addToRecentFiles">optional object AddToRecentFiles</param>
		/// <param name="passwordDocument">optional object PasswordDocument</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		[CustomMethodAttribute]
		[SupportByVersionAttribute("Word", 10,11,12,14,15,16)]
		public void OpenDataSource2000(string name, object format, object confirmConversions, object readOnly, object linkToSource, object addToRecentFiles, object passwordDocument)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(name, format, confirmConversions, readOnly, linkToSource, addToRecentFiles, passwordDocument);
			Invoker.Method(this, "OpenDataSource2000", paramsArray);
		}

		/// <summary>
		/// SupportByVersion Word 10, 11, 12, 14, 15, 16
		/// 
		/// </summary>
		/// <param name="name">string Name</param>
		/// <param name="format">optional object Format</param>
		/// <param name="confirmConversions">optional object ConfirmConversions</param>
		/// <param name="readOnly">optional object ReadOnly</param>
		/// <param name="linkToSource">optional object LinkToSource</param>
		/// <param name="addToRecentFiles">optional object AddToRecentFiles</param>
		/// <param name="passwordDocument">optional object PasswordDocument</param>
		/// <param name="passwordTemplate">optional object PasswordTemplate</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		[CustomMethodAttribute]
		[SupportByVersionAttribute("Word", 10,11,12,14,15,16)]
		public void OpenDataSource2000(string name, object format, object confirmConversions, object readOnly, object linkToSource, object addToRecentFiles, object passwordDocument, object passwordTemplate)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(name, format, confirmConversions, readOnly, linkToSource, addToRecentFiles, passwordDocument, passwordTemplate);
			Invoker.Method(this, "OpenDataSource2000", paramsArray);
		}

		/// <summary>
		/// SupportByVersion Word 10, 11, 12, 14, 15, 16
		/// 
		/// </summary>
		/// <param name="name">string Name</param>
		/// <param name="format">optional object Format</param>
		/// <param name="confirmConversions">optional object ConfirmConversions</param>
		/// <param name="readOnly">optional object ReadOnly</param>
		/// <param name="linkToSource">optional object LinkToSource</param>
		/// <param name="addToRecentFiles">optional object AddToRecentFiles</param>
		/// <param name="passwordDocument">optional object PasswordDocument</param>
		/// <param name="passwordTemplate">optional object PasswordTemplate</param>
		/// <param name="revert">optional object Revert</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		[CustomMethodAttribute]
		[SupportByVersionAttribute("Word", 10,11,12,14,15,16)]
		public void OpenDataSource2000(string name, object format, object confirmConversions, object readOnly, object linkToSource, object addToRecentFiles, object passwordDocument, object passwordTemplate, object revert)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(name, format, confirmConversions, readOnly, linkToSource, addToRecentFiles, passwordDocument, passwordTemplate, revert);
			Invoker.Method(this, "OpenDataSource2000", paramsArray);
		}

		/// <summary>
		/// SupportByVersion Word 10, 11, 12, 14, 15, 16
		/// 
		/// </summary>
		/// <param name="name">string Name</param>
		/// <param name="format">optional object Format</param>
		/// <param name="confirmConversions">optional object ConfirmConversions</param>
		/// <param name="readOnly">optional object ReadOnly</param>
		/// <param name="linkToSource">optional object LinkToSource</param>
		/// <param name="addToRecentFiles">optional object AddToRecentFiles</param>
		/// <param name="passwordDocument">optional object PasswordDocument</param>
		/// <param name="passwordTemplate">optional object PasswordTemplate</param>
		/// <param name="revert">optional object Revert</param>
		/// <param name="writePasswordDocument">optional object WritePasswordDocument</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		[CustomMethodAttribute]
		[SupportByVersionAttribute("Word", 10,11,12,14,15,16)]
		public void OpenDataSource2000(string name, object format, object confirmConversions, object readOnly, object linkToSource, object addToRecentFiles, object passwordDocument, object passwordTemplate, object revert, object writePasswordDocument)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(name, format, confirmConversions, readOnly, linkToSource, addToRecentFiles, passwordDocument, passwordTemplate, revert, writePasswordDocument);
			Invoker.Method(this, "OpenDataSource2000", paramsArray);
		}

		/// <summary>
		/// SupportByVersion Word 10, 11, 12, 14, 15, 16
		/// 
		/// </summary>
		/// <param name="name">string Name</param>
		/// <param name="format">optional object Format</param>
		/// <param name="confirmConversions">optional object ConfirmConversions</param>
		/// <param name="readOnly">optional object ReadOnly</param>
		/// <param name="linkToSource">optional object LinkToSource</param>
		/// <param name="addToRecentFiles">optional object AddToRecentFiles</param>
		/// <param name="passwordDocument">optional object PasswordDocument</param>
		/// <param name="passwordTemplate">optional object PasswordTemplate</param>
		/// <param name="revert">optional object Revert</param>
		/// <param name="writePasswordDocument">optional object WritePasswordDocument</param>
		/// <param name="writePasswordTemplate">optional object WritePasswordTemplate</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		[CustomMethodAttribute]
		[SupportByVersionAttribute("Word", 10,11,12,14,15,16)]
		public void OpenDataSource2000(string name, object format, object confirmConversions, object readOnly, object linkToSource, object addToRecentFiles, object passwordDocument, object passwordTemplate, object revert, object writePasswordDocument, object writePasswordTemplate)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(name, format, confirmConversions, readOnly, linkToSource, addToRecentFiles, passwordDocument, passwordTemplate, revert, writePasswordDocument, writePasswordTemplate);
			Invoker.Method(this, "OpenDataSource2000", paramsArray);
		}

		/// <summary>
		/// SupportByVersion Word 10, 11, 12, 14, 15, 16
		/// 
		/// </summary>
		/// <param name="name">string Name</param>
		/// <param name="format">optional object Format</param>
		/// <param name="confirmConversions">optional object ConfirmConversions</param>
		/// <param name="readOnly">optional object ReadOnly</param>
		/// <param name="linkToSource">optional object LinkToSource</param>
		/// <param name="addToRecentFiles">optional object AddToRecentFiles</param>
		/// <param name="passwordDocument">optional object PasswordDocument</param>
		/// <param name="passwordTemplate">optional object PasswordTemplate</param>
		/// <param name="revert">optional object Revert</param>
		/// <param name="writePasswordDocument">optional object WritePasswordDocument</param>
		/// <param name="writePasswordTemplate">optional object WritePasswordTemplate</param>
		/// <param name="connection">optional object Connection</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		[CustomMethodAttribute]
		[SupportByVersionAttribute("Word", 10,11,12,14,15,16)]
		public void OpenDataSource2000(string name, object format, object confirmConversions, object readOnly, object linkToSource, object addToRecentFiles, object passwordDocument, object passwordTemplate, object revert, object writePasswordDocument, object writePasswordTemplate, object connection)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(name, format, confirmConversions, readOnly, linkToSource, addToRecentFiles, passwordDocument, passwordTemplate, revert, writePasswordDocument, writePasswordTemplate, connection);
			Invoker.Method(this, "OpenDataSource2000", paramsArray);
		}

		/// <summary>
		/// SupportByVersion Word 10, 11, 12, 14, 15, 16
		/// 
		/// </summary>
		/// <param name="name">string Name</param>
		/// <param name="format">optional object Format</param>
		/// <param name="confirmConversions">optional object ConfirmConversions</param>
		/// <param name="readOnly">optional object ReadOnly</param>
		/// <param name="linkToSource">optional object LinkToSource</param>
		/// <param name="addToRecentFiles">optional object AddToRecentFiles</param>
		/// <param name="passwordDocument">optional object PasswordDocument</param>
		/// <param name="passwordTemplate">optional object PasswordTemplate</param>
		/// <param name="revert">optional object Revert</param>
		/// <param name="writePasswordDocument">optional object WritePasswordDocument</param>
		/// <param name="writePasswordTemplate">optional object WritePasswordTemplate</param>
		/// <param name="connection">optional object Connection</param>
		/// <param name="sQLStatement">optional object SQLStatement</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		[CustomMethodAttribute]
		[SupportByVersionAttribute("Word", 10,11,12,14,15,16)]
		public void OpenDataSource2000(string name, object format, object confirmConversions, object readOnly, object linkToSource, object addToRecentFiles, object passwordDocument, object passwordTemplate, object revert, object writePasswordDocument, object writePasswordTemplate, object connection, object sQLStatement)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(name, format, confirmConversions, readOnly, linkToSource, addToRecentFiles, passwordDocument, passwordTemplate, revert, writePasswordDocument, writePasswordTemplate, connection, sQLStatement);
			Invoker.Method(this, "OpenDataSource2000", paramsArray);
		}

		/// <summary>
		/// SupportByVersion Word 10, 11, 12, 14, 15, 16
		/// 
		/// </summary>
		/// <param name="name">string Name</param>
		/// <param name="format">optional object Format</param>
		/// <param name="confirmConversions">optional object ConfirmConversions</param>
		/// <param name="readOnly">optional object ReadOnly</param>
		/// <param name="addToRecentFiles">optional object AddToRecentFiles</param>
		/// <param name="passwordDocument">optional object PasswordDocument</param>
		/// <param name="passwordTemplate">optional object PasswordTemplate</param>
		/// <param name="revert">optional object Revert</param>
		/// <param name="writePasswordDocument">optional object WritePasswordDocument</param>
		/// <param name="writePasswordTemplate">optional object WritePasswordTemplate</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		[SupportByVersionAttribute("Word", 10,11,12,14,15,16)]
		public void OpenHeaderSource2000(string name, object format, object confirmConversions, object readOnly, object addToRecentFiles, object passwordDocument, object passwordTemplate, object revert, object writePasswordDocument, object writePasswordTemplate)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(name, format, confirmConversions, readOnly, addToRecentFiles, passwordDocument, passwordTemplate, revert, writePasswordDocument, writePasswordTemplate);
			Invoker.Method(this, "OpenHeaderSource2000", paramsArray);
		}

		/// <summary>
		/// SupportByVersion Word 10, 11, 12, 14, 15, 16
		/// 
		/// </summary>
		/// <param name="name">string Name</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		[CustomMethodAttribute]
		[SupportByVersionAttribute("Word", 10,11,12,14,15,16)]
		public void OpenHeaderSource2000(string name)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(name);
			Invoker.Method(this, "OpenHeaderSource2000", paramsArray);
		}

		/// <summary>
		/// SupportByVersion Word 10, 11, 12, 14, 15, 16
		/// 
		/// </summary>
		/// <param name="name">string Name</param>
		/// <param name="format">optional object Format</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		[CustomMethodAttribute]
		[SupportByVersionAttribute("Word", 10,11,12,14,15,16)]
		public void OpenHeaderSource2000(string name, object format)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(name, format);
			Invoker.Method(this, "OpenHeaderSource2000", paramsArray);
		}

		/// <summary>
		/// SupportByVersion Word 10, 11, 12, 14, 15, 16
		/// 
		/// </summary>
		/// <param name="name">string Name</param>
		/// <param name="format">optional object Format</param>
		/// <param name="confirmConversions">optional object ConfirmConversions</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		[CustomMethodAttribute]
		[SupportByVersionAttribute("Word", 10,11,12,14,15,16)]
		public void OpenHeaderSource2000(string name, object format, object confirmConversions)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(name, format, confirmConversions);
			Invoker.Method(this, "OpenHeaderSource2000", paramsArray);
		}

		/// <summary>
		/// SupportByVersion Word 10, 11, 12, 14, 15, 16
		/// 
		/// </summary>
		/// <param name="name">string Name</param>
		/// <param name="format">optional object Format</param>
		/// <param name="confirmConversions">optional object ConfirmConversions</param>
		/// <param name="readOnly">optional object ReadOnly</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		[CustomMethodAttribute]
		[SupportByVersionAttribute("Word", 10,11,12,14,15,16)]
		public void OpenHeaderSource2000(string name, object format, object confirmConversions, object readOnly)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(name, format, confirmConversions, readOnly);
			Invoker.Method(this, "OpenHeaderSource2000", paramsArray);
		}

		/// <summary>
		/// SupportByVersion Word 10, 11, 12, 14, 15, 16
		/// 
		/// </summary>
		/// <param name="name">string Name</param>
		/// <param name="format">optional object Format</param>
		/// <param name="confirmConversions">optional object ConfirmConversions</param>
		/// <param name="readOnly">optional object ReadOnly</param>
		/// <param name="addToRecentFiles">optional object AddToRecentFiles</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		[CustomMethodAttribute]
		[SupportByVersionAttribute("Word", 10,11,12,14,15,16)]
		public void OpenHeaderSource2000(string name, object format, object confirmConversions, object readOnly, object addToRecentFiles)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(name, format, confirmConversions, readOnly, addToRecentFiles);
			Invoker.Method(this, "OpenHeaderSource2000", paramsArray);
		}

		/// <summary>
		/// SupportByVersion Word 10, 11, 12, 14, 15, 16
		/// 
		/// </summary>
		/// <param name="name">string Name</param>
		/// <param name="format">optional object Format</param>
		/// <param name="confirmConversions">optional object ConfirmConversions</param>
		/// <param name="readOnly">optional object ReadOnly</param>
		/// <param name="addToRecentFiles">optional object AddToRecentFiles</param>
		/// <param name="passwordDocument">optional object PasswordDocument</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		[CustomMethodAttribute]
		[SupportByVersionAttribute("Word", 10,11,12,14,15,16)]
		public void OpenHeaderSource2000(string name, object format, object confirmConversions, object readOnly, object addToRecentFiles, object passwordDocument)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(name, format, confirmConversions, readOnly, addToRecentFiles, passwordDocument);
			Invoker.Method(this, "OpenHeaderSource2000", paramsArray);
		}

		/// <summary>
		/// SupportByVersion Word 10, 11, 12, 14, 15, 16
		/// 
		/// </summary>
		/// <param name="name">string Name</param>
		/// <param name="format">optional object Format</param>
		/// <param name="confirmConversions">optional object ConfirmConversions</param>
		/// <param name="readOnly">optional object ReadOnly</param>
		/// <param name="addToRecentFiles">optional object AddToRecentFiles</param>
		/// <param name="passwordDocument">optional object PasswordDocument</param>
		/// <param name="passwordTemplate">optional object PasswordTemplate</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		[CustomMethodAttribute]
		[SupportByVersionAttribute("Word", 10,11,12,14,15,16)]
		public void OpenHeaderSource2000(string name, object format, object confirmConversions, object readOnly, object addToRecentFiles, object passwordDocument, object passwordTemplate)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(name, format, confirmConversions, readOnly, addToRecentFiles, passwordDocument, passwordTemplate);
			Invoker.Method(this, "OpenHeaderSource2000", paramsArray);
		}

		/// <summary>
		/// SupportByVersion Word 10, 11, 12, 14, 15, 16
		/// 
		/// </summary>
		/// <param name="name">string Name</param>
		/// <param name="format">optional object Format</param>
		/// <param name="confirmConversions">optional object ConfirmConversions</param>
		/// <param name="readOnly">optional object ReadOnly</param>
		/// <param name="addToRecentFiles">optional object AddToRecentFiles</param>
		/// <param name="passwordDocument">optional object PasswordDocument</param>
		/// <param name="passwordTemplate">optional object PasswordTemplate</param>
		/// <param name="revert">optional object Revert</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		[CustomMethodAttribute]
		[SupportByVersionAttribute("Word", 10,11,12,14,15,16)]
		public void OpenHeaderSource2000(string name, object format, object confirmConversions, object readOnly, object addToRecentFiles, object passwordDocument, object passwordTemplate, object revert)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(name, format, confirmConversions, readOnly, addToRecentFiles, passwordDocument, passwordTemplate, revert);
			Invoker.Method(this, "OpenHeaderSource2000", paramsArray);
		}

		/// <summary>
		/// SupportByVersion Word 10, 11, 12, 14, 15, 16
		/// 
		/// </summary>
		/// <param name="name">string Name</param>
		/// <param name="format">optional object Format</param>
		/// <param name="confirmConversions">optional object ConfirmConversions</param>
		/// <param name="readOnly">optional object ReadOnly</param>
		/// <param name="addToRecentFiles">optional object AddToRecentFiles</param>
		/// <param name="passwordDocument">optional object PasswordDocument</param>
		/// <param name="passwordTemplate">optional object PasswordTemplate</param>
		/// <param name="revert">optional object Revert</param>
		/// <param name="writePasswordDocument">optional object WritePasswordDocument</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		[CustomMethodAttribute]
		[SupportByVersionAttribute("Word", 10,11,12,14,15,16)]
		public void OpenHeaderSource2000(string name, object format, object confirmConversions, object readOnly, object addToRecentFiles, object passwordDocument, object passwordTemplate, object revert, object writePasswordDocument)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(name, format, confirmConversions, readOnly, addToRecentFiles, passwordDocument, passwordTemplate, revert, writePasswordDocument);
			Invoker.Method(this, "OpenHeaderSource2000", paramsArray);
		}

		/// <summary>
		/// SupportByVersion Word 10, 11, 12, 14, 15, 16
		/// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff844772.aspx
		/// </summary>
		/// <param name="initialState">object InitialState</param>
		/// <param name="showDocumentStep">optional object ShowDocumentStep</param>
		/// <param name="showTemplateStep">optional object ShowTemplateStep</param>
		/// <param name="showDataStep">optional object ShowDataStep</param>
		/// <param name="showWriteStep">optional object ShowWriteStep</param>
		/// <param name="showPreviewStep">optional object ShowPreviewStep</param>
		/// <param name="showMergeStep">optional object ShowMergeStep</param>
		[SupportByVersionAttribute("Word", 10,11,12,14,15,16)]
		public void ShowWizard(object initialState, object showDocumentStep, object showTemplateStep, object showDataStep, object showWriteStep, object showPreviewStep, object showMergeStep)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(initialState, showDocumentStep, showTemplateStep, showDataStep, showWriteStep, showPreviewStep, showMergeStep);
			Invoker.Method(this, "ShowWizard", paramsArray);
		}

		/// <summary>
		/// SupportByVersion Word 10, 11, 12, 14, 15, 16
		/// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff844772.aspx
		/// </summary>
		/// <param name="initialState">object InitialState</param>
		[CustomMethodAttribute]
		[SupportByVersionAttribute("Word", 10,11,12,14,15,16)]
		public void ShowWizard(object initialState)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(initialState);
			Invoker.Method(this, "ShowWizard", paramsArray);
		}

		/// <summary>
		/// SupportByVersion Word 10, 11, 12, 14, 15, 16
		/// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff844772.aspx
		/// </summary>
		/// <param name="initialState">object InitialState</param>
		/// <param name="showDocumentStep">optional object ShowDocumentStep</param>
		[CustomMethodAttribute]
		[SupportByVersionAttribute("Word", 10,11,12,14,15,16)]
		public void ShowWizard(object initialState, object showDocumentStep)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(initialState, showDocumentStep);
			Invoker.Method(this, "ShowWizard", paramsArray);
		}

		/// <summary>
		/// SupportByVersion Word 10, 11, 12, 14, 15, 16
		/// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff844772.aspx
		/// </summary>
		/// <param name="initialState">object InitialState</param>
		/// <param name="showDocumentStep">optional object ShowDocumentStep</param>
		/// <param name="showTemplateStep">optional object ShowTemplateStep</param>
		[CustomMethodAttribute]
		[SupportByVersionAttribute("Word", 10,11,12,14,15,16)]
		public void ShowWizard(object initialState, object showDocumentStep, object showTemplateStep)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(initialState, showDocumentStep, showTemplateStep);
			Invoker.Method(this, "ShowWizard", paramsArray);
		}

		/// <summary>
		/// SupportByVersion Word 10, 11, 12, 14, 15, 16
		/// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff844772.aspx
		/// </summary>
		/// <param name="initialState">object InitialState</param>
		/// <param name="showDocumentStep">optional object ShowDocumentStep</param>
		/// <param name="showTemplateStep">optional object ShowTemplateStep</param>
		/// <param name="showDataStep">optional object ShowDataStep</param>
		[CustomMethodAttribute]
		[SupportByVersionAttribute("Word", 10,11,12,14,15,16)]
		public void ShowWizard(object initialState, object showDocumentStep, object showTemplateStep, object showDataStep)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(initialState, showDocumentStep, showTemplateStep, showDataStep);
			Invoker.Method(this, "ShowWizard", paramsArray);
		}

		/// <summary>
		/// SupportByVersion Word 10, 11, 12, 14, 15, 16
		/// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff844772.aspx
		/// </summary>
		/// <param name="initialState">object InitialState</param>
		/// <param name="showDocumentStep">optional object ShowDocumentStep</param>
		/// <param name="showTemplateStep">optional object ShowTemplateStep</param>
		/// <param name="showDataStep">optional object ShowDataStep</param>
		/// <param name="showWriteStep">optional object ShowWriteStep</param>
		[CustomMethodAttribute]
		[SupportByVersionAttribute("Word", 10,11,12,14,15,16)]
		public void ShowWizard(object initialState, object showDocumentStep, object showTemplateStep, object showDataStep, object showWriteStep)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(initialState, showDocumentStep, showTemplateStep, showDataStep, showWriteStep);
			Invoker.Method(this, "ShowWizard", paramsArray);
		}

		/// <summary>
		/// SupportByVersion Word 10, 11, 12, 14, 15, 16
		/// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff844772.aspx
		/// </summary>
		/// <param name="initialState">object InitialState</param>
		/// <param name="showDocumentStep">optional object ShowDocumentStep</param>
		/// <param name="showTemplateStep">optional object ShowTemplateStep</param>
		/// <param name="showDataStep">optional object ShowDataStep</param>
		/// <param name="showWriteStep">optional object ShowWriteStep</param>
		/// <param name="showPreviewStep">optional object ShowPreviewStep</param>
		[CustomMethodAttribute]
		[SupportByVersionAttribute("Word", 10,11,12,14,15,16)]
		public void ShowWizard(object initialState, object showDocumentStep, object showTemplateStep, object showDataStep, object showWriteStep, object showPreviewStep)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(initialState, showDocumentStep, showTemplateStep, showDataStep, showWriteStep, showPreviewStep);
			Invoker.Method(this, "ShowWizard", paramsArray);
		}

		#endregion
		#pragma warning restore
	}
}