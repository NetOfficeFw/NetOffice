using System;
using NetRuntimeSystem = System;
using System.ComponentModel;
using NetOffice.Attributes;

namespace NetOffice.ADODBApi
{
	/// <summary>
	/// DispatchInterface Recordset20 
	/// SupportByVersion ADODB, 2.1,2.5
	/// </summary>
	[SupportByVersion("ADODB", 2.1,2.5)]
	[EntityType(EntityType.IsDispatchInterface), BaseType]
 	public class Recordset20 : Recordset15
	{
		#pragma warning disable

		#region Type Information

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
                    _type = typeof(Recordset20);
                return _type;
            }
        }
        
        #endregion
        
		#region Ctor

		/// <param name="factory">current used factory core</param>
		/// <param name="parentObject">object there has created the proxy</param>
		/// <param name="proxyShare">proxy share instead if com proxy</param>
		public Recordset20(Core factory, ICOMObject parentObject, COMProxyShare proxyShare) : base(factory, parentObject, proxyShare)
		{
		}

		///<param name="factory">current used factory core</param>
		///<param name="parentObject">object there has created the proxy</param>
        ///<param name="comProxy">inner wrapped COM proxy</param>
		public Recordset20(Core factory, ICOMObject parentObject, object comProxy) : base(factory, parentObject, comProxy)
		{
			
		}

        ///<param name="parentObject">object there has created the proxy</param>
        ///<param name="comProxy">inner wrapped COM proxy</param>
        [EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public Recordset20(ICOMObject parentObject, object comProxy) : base(parentObject, comProxy)
		{
		}
		
		///<param name="factory">current used factory core</param>
		///<param name="parentObject">object there has created the proxy</param>
        ///<param name="comProxy">inner wrapped COM proxy</param>
        ///<param name="comProxyType">Type of inner wrapped COM proxy"</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public Recordset20(Core factory, ICOMObject parentObject, object comProxy, NetRuntimeSystem.Type comProxyType) : base(factory, parentObject, comProxy, comProxyType)
		{

		}

		///<param name="parentObject">object there has created the proxy</param>
        ///<param name="comProxy">inner wrapped COM proxy</param>
        ///<param name="comProxyType">Type of inner wrapped COM proxy"</param>
        [EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public Recordset20(ICOMObject parentObject, object comProxy, NetRuntimeSystem.Type comProxyType) : base(parentObject, comProxy, comProxyType)
		{
		}
		
		///<param name="replacedObject">object to replaced. replacedObject are not usable after this action</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public Recordset20(ICOMObject replacedObject) : base(replacedObject)
		{
		}
		
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public Recordset20() : base()
		{
		}
		
		/// <param name="progId">registered progID</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public Recordset20(string progId) : base(progId)
		{
		}
		
		#endregion
		
		#region Properties

		/// <summary>
		/// SupportByVersion ADODB 2.1, 2.5
		/// Get/Set
		/// Unknown COM Proxy
		/// </summary>
		[SupportByVersion("ADODB", 2.1,2.5), ProxyResult]
		public object DataSource
		{
			get
			{
				return Factory.ExecuteReferencePropertyGet(this, "DataSource");
			}
			set
			{
				Factory.ExecuteReferencePropertySet(this, "DataSource", value);
			}
		}

		/// <summary>
		/// SupportByVersion ADODB 2.1, 2.5
		/// Get
		/// Unknown COM Proxy
		/// </summary>
		[SupportByVersion("ADODB", 2.1,2.5), ProxyResult]
		public object ActiveCommand
		{
			get
			{
				return Factory.ExecuteReferencePropertyGet(this, "ActiveCommand");
			}
		}

		/// <summary>
		/// SupportByVersion ADODB 2.1, 2.5
		/// Get/Set
		/// </summary>
		[SupportByVersion("ADODB", 2.1,2.5)]
		public bool StayInSync
		{
			get
			{
				return Factory.ExecuteBoolPropertyGet(this, "StayInSync");
			}
			set
			{
				Factory.ExecuteValuePropertySet(this, "StayInSync", value);
			}
		}

		/// <summary>
		/// SupportByVersion ADODB 2.1, 2.5
		/// Get/Set
		/// </summary>
		[SupportByVersion("ADODB", 2.1,2.5)]
		public string DataMember
		{
			get
			{
				return Factory.ExecuteStringPropertyGet(this, "DataMember");
			}
			set
			{
				Factory.ExecuteValuePropertySet(this, "DataMember", value);
			}
		}

		#endregion

		#region Methods

		/// <summary>
		/// SupportByVersion ADODB 2.1, 2.5
		/// </summary>
		[SupportByVersion("ADODB", 2.1,2.5)]
		public void Cancel()
		{
			 Factory.ExecuteMethod(this, "Cancel");
		}

		/// <summary>
		/// SupportByVersion ADODB 2.1
		/// </summary>
		/// <param name="fileName">optional string fileName</param>
		/// <param name="persistFormat">optional NetOffice.ADODBApi.Enums.PersistFormatEnum PersistFormat = 0</param>
		[SupportByVersion("ADODB", 2.1)]
		public void Save(object fileName, object persistFormat)
		{
			 Factory.ExecuteMethod(this, "Save", fileName, persistFormat);
		}

		/// <summary>
		/// SupportByVersion ADODB 2.1
		/// </summary>
		[CustomMethod]
		[SupportByVersion("ADODB", 2.1)]
		public void Save()
		{
			 Factory.ExecuteMethod(this, "Save");
		}

		/// <summary>
		/// SupportByVersion ADODB 2.1
		/// </summary>
		/// <param name="fileName">optional string fileName</param>
		[CustomMethod]
		[SupportByVersion("ADODB", 2.1)]
		public void Save(object fileName)
		{
			 Factory.ExecuteMethod(this, "Save", fileName);
		}

		/// <summary>
		/// SupportByVersion ADODB 2.1, 2.5
		/// </summary>
		/// <param name="stringFormat">optional NetOffice.ADODBApi.Enums.StringFormatEnum StringFormat = 2</param>
		/// <param name="numRows">optional Int32 NumRows = -1</param>
		/// <param name="columnDelimeter">optional string columnDelimeter</param>
		/// <param name="rowDelimeter">optional string rowDelimeter</param>
		/// <param name="nullExpr">optional string nullExpr</param>
		[SupportByVersion("ADODB", 2.1,2.5)]
		public string GetString(object stringFormat, object numRows, object columnDelimeter, object rowDelimeter, object nullExpr)
		{
			return Factory.ExecuteStringMethodGet(this, "GetString", new object[]{ stringFormat, numRows, columnDelimeter, rowDelimeter, nullExpr });
		}

		/// <summary>
		/// SupportByVersion ADODB 2.1, 2.5
		/// </summary>
		[CustomMethod]
		[SupportByVersion("ADODB", 2.1,2.5)]
		public string GetString()
		{
			return Factory.ExecuteStringMethodGet(this, "GetString");
		}

		/// <summary>
		/// SupportByVersion ADODB 2.1, 2.5
		/// </summary>
		/// <param name="stringFormat">optional NetOffice.ADODBApi.Enums.StringFormatEnum StringFormat = 2</param>
		[CustomMethod]
		[SupportByVersion("ADODB", 2.1,2.5)]
		public string GetString(object stringFormat)
		{
			return Factory.ExecuteStringMethodGet(this, "GetString", stringFormat);
		}

		/// <summary>
		/// SupportByVersion ADODB 2.1, 2.5
		/// </summary>
		/// <param name="stringFormat">optional NetOffice.ADODBApi.Enums.StringFormatEnum StringFormat = 2</param>
		/// <param name="numRows">optional Int32 NumRows = -1</param>
		[CustomMethod]
		[SupportByVersion("ADODB", 2.1,2.5)]
		public string GetString(object stringFormat, object numRows)
		{
			return Factory.ExecuteStringMethodGet(this, "GetString", stringFormat, numRows);
		}

		/// <summary>
		/// SupportByVersion ADODB 2.1, 2.5
		/// </summary>
		/// <param name="stringFormat">optional NetOffice.ADODBApi.Enums.StringFormatEnum StringFormat = 2</param>
		/// <param name="numRows">optional Int32 NumRows = -1</param>
		/// <param name="columnDelimeter">optional string columnDelimeter</param>
		[CustomMethod]
		[SupportByVersion("ADODB", 2.1,2.5)]
		public string GetString(object stringFormat, object numRows, object columnDelimeter)
		{
			return Factory.ExecuteStringMethodGet(this, "GetString", stringFormat, numRows, columnDelimeter);
		}

		/// <summary>
		/// SupportByVersion ADODB 2.1, 2.5
		/// </summary>
		/// <param name="stringFormat">optional NetOffice.ADODBApi.Enums.StringFormatEnum StringFormat = 2</param>
		/// <param name="numRows">optional Int32 NumRows = -1</param>
		/// <param name="columnDelimeter">optional string columnDelimeter</param>
		/// <param name="rowDelimeter">optional string rowDelimeter</param>
		[CustomMethod]
		[SupportByVersion("ADODB", 2.1,2.5)]
		public string GetString(object stringFormat, object numRows, object columnDelimeter, object rowDelimeter)
		{
			return Factory.ExecuteStringMethodGet(this, "GetString", stringFormat, numRows, columnDelimeter, rowDelimeter);
		}

		/// <summary>
		/// SupportByVersion ADODB 2.1, 2.5
		/// </summary>
		/// <param name="bookmark1">object bookmark1</param>
		/// <param name="bookmark2">object bookmark2</param>
		[SupportByVersion("ADODB", 2.1,2.5)]
		public NetOffice.ADODBApi.Enums.CompareEnum CompareBookmarks(object bookmark1, object bookmark2)
		{
			return Factory.ExecuteEnumMethodGet<NetOffice.ADODBApi.Enums.CompareEnum>(this, "CompareBookmarks", bookmark1, bookmark2);
		}

		/// <summary>
		/// SupportByVersion ADODB 2.1, 2.5
		/// </summary>
		/// <param name="lockType">optional NetOffice.ADODBApi.Enums.LockTypeEnum LockType = -1</param>
		[SupportByVersion("ADODB", 2.1,2.5)]
		[BaseResult]
		public NetOffice.ADODBApi._Recordset Clone(object lockType)
		{
			return Factory.ExecuteBaseReferenceMethodGet<NetOffice.ADODBApi._Recordset>(this, "Clone", lockType);
		}

		/// <summary>
		/// SupportByVersion ADODB 2.1, 2.5
		/// </summary>
		[CustomMethod]
		[BaseResult]
		[SupportByVersion("ADODB", 2.1,2.5)]
		public NetOffice.ADODBApi._Recordset Clone()
		{
			return Factory.ExecuteBaseReferenceMethodGet<NetOffice.ADODBApi._Recordset>(this, "Clone");
		}

		/// <summary>
		/// SupportByVersion ADODB 2.1, 2.5
		/// </summary>
		/// <param name="affectRecords">optional NetOffice.ADODBApi.Enums.AffectEnum AffectRecords = 3</param>
		/// <param name="resyncValues">optional NetOffice.ADODBApi.Enums.ResyncEnum ResyncValues = 2</param>
		[SupportByVersion("ADODB", 2.1,2.5)]
		public void Resync(object affectRecords, object resyncValues)
		{
			 Factory.ExecuteMethod(this, "Resync", affectRecords, resyncValues);
		}

		/// <summary>
		/// SupportByVersion ADODB 2.1, 2.5
		/// </summary>
		[CustomMethod]
		[SupportByVersion("ADODB", 2.1,2.5)]
		public void Resync()
		{
			 Factory.ExecuteMethod(this, "Resync");
		}

		/// <summary>
		/// SupportByVersion ADODB 2.1, 2.5
		/// </summary>
		/// <param name="affectRecords">optional NetOffice.ADODBApi.Enums.AffectEnum AffectRecords = 3</param>
		[CustomMethod]
		[SupportByVersion("ADODB", 2.1,2.5)]
		public void Resync(object affectRecords)
		{
			 Factory.ExecuteMethod(this, "Resync", affectRecords);
		}

		/// <summary>
		/// SupportByVersion ADODB 2.5
		/// </summary>
		/// <param name="fileName">optional string fileName</param>
		/// <param name="persistFormat">optional NetOffice.ADODBApi.Enums.PersistFormatEnum PersistFormat = 0</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		[SupportByVersion("ADODB", 2.5)]
		public void _xSave(object fileName, object persistFormat)
		{
			 Factory.ExecuteMethod(this, "_xSave", fileName, persistFormat);
		}

		/// <summary>
		/// SupportByVersion ADODB 2.5
		/// </summary>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		[CustomMethod]
		[SupportByVersion("ADODB", 2.5)]
		public void _xSave()
		{
			 Factory.ExecuteMethod(this, "_xSave");
		}

		/// <summary>
		/// SupportByVersion ADODB 2.5
		/// </summary>
		/// <param name="fileName">optional string fileName</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		[CustomMethod]
		[SupportByVersion("ADODB", 2.5)]
		public void _xSave(object fileName)
		{
			 Factory.ExecuteMethod(this, "_xSave", fileName);
		}

		#endregion

		#pragma warning restore
	}
}
