using System;
using NetRuntimeSystem = System;
using System.ComponentModel;
using NetOffice.Attributes;
using NetOffice.ADODBApi;

namespace NetOffice.ADODBApi.Behind
{
	/// <summary>
	/// DispatchInterface Recordset20 
	/// SupportByVersion ADODB, 2.1,2.5
	/// </summary>
	[SupportByVersion("ADODB", 2.1,2.5)]
	[EntityType(EntityType.IsDispatchInterface), BaseType]
 	public class Recordset20 : Recordset15, NetOffice.ADODBApi.Recordset20
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
                    _contractType = typeof(NetOffice.ADODBApi.Recordset20);
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
                    _type = typeof(Recordset20);
                return _type;
            }
        }
        
        #endregion
        
		#region Ctor

		/// <summary>
		/// Stub Ctor, not indented to use
		/// </summary>
		public Recordset20() : base()
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
		public virtual object DataSource
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteReferencePropertyGet(this, "DataSource");
			}
			set
			{
				InvokerService.InvokeInternal.ExecuteReferencePropertySet(this, "DataSource", value);
			}
		}

		/// <summary>
		/// SupportByVersion ADODB 2.1, 2.5
		/// Get
		/// Unknown COM Proxy
		/// </summary>
		[SupportByVersion("ADODB", 2.1,2.5), ProxyResult]
		public virtual object ActiveCommand
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteReferencePropertyGet(this, "ActiveCommand");
			}
		}

		/// <summary>
		/// SupportByVersion ADODB 2.1, 2.5
		/// Get/Set
		/// </summary>
		[SupportByVersion("ADODB", 2.1,2.5)]
		public virtual bool StayInSync
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteBoolPropertyGet(this, "StayInSync");
			}
			set
			{
				InvokerService.InvokeInternal.ExecuteValuePropertySet(this, "StayInSync", value);
			}
		}

		/// <summary>
		/// SupportByVersion ADODB 2.1, 2.5
		/// Get/Set
		/// </summary>
		[SupportByVersion("ADODB", 2.1,2.5)]
		public virtual string DataMember
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteStringPropertyGet(this, "DataMember");
			}
			set
			{
				InvokerService.InvokeInternal.ExecuteValuePropertySet(this, "DataMember", value);
			}
		}

		#endregion

		#region Methods

		/// <summary>
		/// SupportByVersion ADODB 2.1, 2.5
		/// </summary>
		[SupportByVersion("ADODB", 2.1,2.5)]
		public virtual void Cancel()
		{
			 InvokerService.InvokeInternal.ExecuteMethod(this, "Cancel");
		}

		/// <summary>
		/// SupportByVersion ADODB 2.1
		/// </summary>
		/// <param name="fileName">optional string fileName</param>
		/// <param name="persistFormat">optional NetOffice.ADODBApi.Enums.PersistFormatEnum PersistFormat = 0</param>
		[SupportByVersion("ADODB", 2.1)]
		public virtual void Save(object fileName, object persistFormat)
		{
			 InvokerService.InvokeInternal.ExecuteMethod(this, "Save", fileName, persistFormat);
		}

		/// <summary>
		/// SupportByVersion ADODB 2.1
		/// </summary>
		[CustomMethod]
		[SupportByVersion("ADODB", 2.1)]
		public virtual void Save()
		{
			 InvokerService.InvokeInternal.ExecuteMethod(this, "Save");
		}

		/// <summary>
		/// SupportByVersion ADODB 2.1
		/// </summary>
		/// <param name="fileName">optional string fileName</param>
		[CustomMethod]
		[SupportByVersion("ADODB", 2.1)]
		public virtual void Save(object fileName)
		{
			 InvokerService.InvokeInternal.ExecuteMethod(this, "Save", fileName);
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
		public virtual string GetString(object stringFormat, object numRows, object columnDelimeter, object rowDelimeter, object nullExpr)
		{
			return InvokerService.InvokeInternal.ExecuteStringMethodGet(this, "GetString", new object[]{ stringFormat, numRows, columnDelimeter, rowDelimeter, nullExpr });
		}

		/// <summary>
		/// SupportByVersion ADODB 2.1, 2.5
		/// </summary>
		[CustomMethod]
		[SupportByVersion("ADODB", 2.1,2.5)]
		public virtual string GetString()
		{
			return InvokerService.InvokeInternal.ExecuteStringMethodGet(this, "GetString");
		}

		/// <summary>
		/// SupportByVersion ADODB 2.1, 2.5
		/// </summary>
		/// <param name="stringFormat">optional NetOffice.ADODBApi.Enums.StringFormatEnum StringFormat = 2</param>
		[CustomMethod]
		[SupportByVersion("ADODB", 2.1,2.5)]
		public virtual string GetString(object stringFormat)
		{
			return InvokerService.InvokeInternal.ExecuteStringMethodGet(this, "GetString", stringFormat);
		}

		/// <summary>
		/// SupportByVersion ADODB 2.1, 2.5
		/// </summary>
		/// <param name="stringFormat">optional NetOffice.ADODBApi.Enums.StringFormatEnum StringFormat = 2</param>
		/// <param name="numRows">optional Int32 NumRows = -1</param>
		[CustomMethod]
		[SupportByVersion("ADODB", 2.1,2.5)]
		public virtual string GetString(object stringFormat, object numRows)
		{
			return InvokerService.InvokeInternal.ExecuteStringMethodGet(this, "GetString", stringFormat, numRows);
		}

		/// <summary>
		/// SupportByVersion ADODB 2.1, 2.5
		/// </summary>
		/// <param name="stringFormat">optional NetOffice.ADODBApi.Enums.StringFormatEnum StringFormat = 2</param>
		/// <param name="numRows">optional Int32 NumRows = -1</param>
		/// <param name="columnDelimeter">optional string columnDelimeter</param>
		[CustomMethod]
		[SupportByVersion("ADODB", 2.1,2.5)]
		public virtual string GetString(object stringFormat, object numRows, object columnDelimeter)
		{
			return InvokerService.InvokeInternal.ExecuteStringMethodGet(this, "GetString", stringFormat, numRows, columnDelimeter);
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
		public virtual string GetString(object stringFormat, object numRows, object columnDelimeter, object rowDelimeter)
		{
			return InvokerService.InvokeInternal.ExecuteStringMethodGet(this, "GetString", stringFormat, numRows, columnDelimeter, rowDelimeter);
		}

		/// <summary>
		/// SupportByVersion ADODB 2.1, 2.5
		/// </summary>
		/// <param name="bookmark1">object bookmark1</param>
		/// <param name="bookmark2">object bookmark2</param>
		[SupportByVersion("ADODB", 2.1,2.5)]
		public virtual NetOffice.ADODBApi.Enums.CompareEnum CompareBookmarks(object bookmark1, object bookmark2)
		{
			return InvokerService.InvokeInternal.ExecuteEnumMethodGet<NetOffice.ADODBApi.Enums.CompareEnum>(this, "CompareBookmarks", bookmark1, bookmark2);
		}

		/// <summary>
		/// SupportByVersion ADODB 2.1, 2.5
		/// </summary>
		/// <param name="lockType">optional NetOffice.ADODBApi.Enums.LockTypeEnum LockType = -1</param>
		[SupportByVersion("ADODB", 2.1,2.5)]
		[BaseResult]
		public virtual NetOffice.ADODBApi._Recordset Clone(object lockType)
		{
			return InvokerService.InvokeInternal.ExecuteBaseReferenceMethodGet<NetOffice.ADODBApi._Recordset>(this, "Clone", lockType);
		}

		/// <summary>
		/// SupportByVersion ADODB 2.1, 2.5
		/// </summary>
		[CustomMethod]
		[BaseResult]
		[SupportByVersion("ADODB", 2.1,2.5)]
		public virtual NetOffice.ADODBApi._Recordset Clone()
		{
			return InvokerService.InvokeInternal.ExecuteBaseReferenceMethodGet<NetOffice.ADODBApi._Recordset>(this, "Clone");
		}

		/// <summary>
		/// SupportByVersion ADODB 2.1, 2.5
		/// </summary>
		/// <param name="affectRecords">optional NetOffice.ADODBApi.Enums.AffectEnum AffectRecords = 3</param>
		/// <param name="resyncValues">optional NetOffice.ADODBApi.Enums.ResyncEnum ResyncValues = 2</param>
		[SupportByVersion("ADODB", 2.1,2.5)]
		public virtual void Resync(object affectRecords, object resyncValues)
		{
			 InvokerService.InvokeInternal.ExecuteMethod(this, "Resync", affectRecords, resyncValues);
		}

		/// <summary>
		/// SupportByVersion ADODB 2.1, 2.5
		/// </summary>
		[CustomMethod]
		[SupportByVersion("ADODB", 2.1,2.5)]
		public virtual void Resync()
		{
			 InvokerService.InvokeInternal.ExecuteMethod(this, "Resync");
		}

		/// <summary>
		/// SupportByVersion ADODB 2.1, 2.5
		/// </summary>
		/// <param name="affectRecords">optional NetOffice.ADODBApi.Enums.AffectEnum AffectRecords = 3</param>
		[CustomMethod]
		[SupportByVersion("ADODB", 2.1,2.5)]
		public virtual void Resync(object affectRecords)
		{
			 InvokerService.InvokeInternal.ExecuteMethod(this, "Resync", affectRecords);
		}

		/// <summary>
		/// SupportByVersion ADODB 2.5
		/// </summary>
		/// <param name="fileName">optional string fileName</param>
		/// <param name="persistFormat">optional NetOffice.ADODBApi.Enums.PersistFormatEnum PersistFormat = 0</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		[SupportByVersion("ADODB", 2.5)]
		public virtual void _xSave(object fileName, object persistFormat)
		{
			 InvokerService.InvokeInternal.ExecuteMethod(this, "_xSave", fileName, persistFormat);
		}

		/// <summary>
		/// SupportByVersion ADODB 2.5
		/// </summary>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		[CustomMethod]
		[SupportByVersion("ADODB", 2.5)]
		public virtual void _xSave()
		{
			 InvokerService.InvokeInternal.ExecuteMethod(this, "_xSave");
		}

		/// <summary>
		/// SupportByVersion ADODB 2.5
		/// </summary>
		/// <param name="fileName">optional string fileName</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		[CustomMethod]
		[SupportByVersion("ADODB", 2.5)]
		public virtual void _xSave(object fileName)
		{
			 InvokerService.InvokeInternal.ExecuteMethod(this, "_xSave", fileName);
		}

		#endregion

		#pragma warning restore
	}
}

