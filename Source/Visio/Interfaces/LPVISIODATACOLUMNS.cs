using System;
using NetRuntimeSystem = System;
using System.ComponentModel;
using System.Runtime.CompilerServices;
using System.Reflection;
using System.Collections.Generic;
using System.Collections;
using NetOffice;
namespace NetOffice.VisioApi
{
	///<summary>
	/// Interface LPVISIODATACOLUMNS 
	/// SupportByVersion Visio, 12,14
	///</summary>
	[SupportByVersionAttribute("Visio", 12,14)]
	[EntityTypeAttribute(EntityType.IsInterface)]
	public class LPVISIODATACOLUMNS : COMObject ,IEnumerable<NetOffice.VisioApi.IVDataColumn>
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
                    _type = typeof(LPVISIODATACOLUMNS);
                    
                return _type;
            }
        }
        
        #endregion
        
		#region Construction

        /// <param name="parentObject">object there has created the proxy</param>
        /// <param name="comProxy">inner wrapped COM proxy</param>
        [EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public LPVISIODATACOLUMNS(COMObject parentObject, object comProxy) : base(parentObject, comProxy)
		{
		}
		
		/// <param name="parentObject">object there has created the proxy</param>
        /// <param name="comProxy">inner wrapped COM proxy</param>
        /// <param name="comProxyType">Type of inner wrapped COM proxy"</param>
        [EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public LPVISIODATACOLUMNS(COMObject parentObject, object comProxy, NetRuntimeSystem.Type comProxyType) : base(parentObject, comProxy, comProxyType)
		{
		}
		
		/// <param name="replacedObject">object to replaced. replacedObject are not usable after this action</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public LPVISIODATACOLUMNS(COMObject replacedObject) : base(replacedObject)
		{
		}
		
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public LPVISIODATACOLUMNS() : base()
		{
		}
		
		/// <param name="progId">registered ProgID</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public LPVISIODATACOLUMNS(string progId) : base(progId)
		{
		}
		
		#endregion
		
		#region Properties

		/// <summary>
		/// SupportByVersion Visio 12, 14
		/// Get
		/// </summary>
		[SupportByVersionAttribute("Visio", 12,14)]
		public NetOffice.VisioApi.IVApplication Application
		{
			get
			{
				object[] paramsArray = null;
				object returnItem = Invoker.PropertyGet(this, "Application", paramsArray);
				NetOffice.VisioApi.IVApplication newObject = NetOffice.Factory.CreateObjectFromComProxy(this,returnItem) as NetOffice.VisioApi.IVApplication;
				return newObject;
			}
		}

		/// <summary>
		/// SupportByVersion Visio 12, 14
		/// Get
		/// </summary>
		[SupportByVersionAttribute("Visio", 12,14)]
		public Int16 Stat
		{
			get
			{
				object[] paramsArray = null;
				object returnItem = Invoker.PropertyGet(this, "Stat", paramsArray);
				return NetRuntimeSystem.Convert.ToInt16(returnItem);
			}
		}

		/// <summary>
		/// SupportByVersion Visio 12, 14
		/// Get
		/// </summary>
		[SupportByVersionAttribute("Visio", 12,14)]
		public NetOffice.VisioApi.IVDocument Document
		{
			get
			{
				object[] paramsArray = null;
				object returnItem = Invoker.PropertyGet(this, "Document", paramsArray);
				NetOffice.VisioApi.IVDocument newObject = NetOffice.Factory.CreateObjectFromComProxy(this,returnItem) as NetOffice.VisioApi.IVDocument;
				return newObject;
			}
		}

		/// <summary>
		/// SupportByVersion Visio 12, 14
		/// Get
		/// </summary>
		[SupportByVersionAttribute("Visio", 12,14)]
		public Int16 ObjectType
		{
			get
			{
				object[] paramsArray = null;
				object returnItem = Invoker.PropertyGet(this, "ObjectType", paramsArray);
				return NetRuntimeSystem.Convert.ToInt16(returnItem);
			}
		}

		/// <summary>
		/// SupportByVersion Visio 12, 14
		/// Get
		/// </summary>
		[SupportByVersionAttribute("Visio", 12,14)]
		public NetOffice.VisioApi.IVDataRecordset DataRecordset
		{
			get
			{
				object[] paramsArray = null;
				object returnItem = Invoker.PropertyGet(this, "DataRecordset", paramsArray);
				NetOffice.VisioApi.IVDataRecordset newObject = NetOffice.Factory.CreateObjectFromComProxy(this,returnItem) as NetOffice.VisioApi.IVDataRecordset;
				return newObject;
			}
		}

		/// <summary>
		/// SupportByVersion Visio 12, 14
		/// Get
		/// </summary>
		[SupportByVersionAttribute("Visio", 12,14)]
		public Int32 Count
		{
			get
			{
				object[] paramsArray = null;
				object returnItem = Invoker.PropertyGet(this, "Count", paramsArray);
				return NetRuntimeSystem.Convert.ToInt32(returnItem);
			}
		}

		/// <summary>
		/// SupportByVersion Visio 12, 14
		/// Get
		/// </summary>
		/// <param name="indexOrName">object IndexOrName</param>
		[SupportByVersionAttribute("Visio", 12,14)]
		[NetRuntimeSystem.Runtime.CompilerServices.IndexerName("Item")]
		public NetOffice.VisioApi.IVDataColumn this[object indexOrName]
		{
			get
{			
			object[] paramsArray = Invoker.ValidateParamsArray(indexOrName);
			object returnItem = Invoker.PropertyGet(this, "Item", paramsArray);
			NetOffice.VisioApi.IVDataColumn newObject = NetOffice.Factory.CreateObjectFromComProxy(this,returnItem) as NetOffice.VisioApi.IVDataColumn;
			return newObject;
			}
		}

		#endregion

		#region Methods

		/// <summary>
		/// SupportByVersion Visio 12, 14
		/// </summary>
		/// <param name="columnNames">String[] ColumnNames</param>
		/// <param name="properties">Int32[] Properties</param>
		/// <param name="values">object[] Values</param>
		[SupportByVersionAttribute("Visio", 12,14)]
		public void SetColumnProperties(String[] columnNames, Int32[] properties, object[] values)
		{
			object[] paramsArray = Invoker.ValidateParamsArray((object)columnNames, (object)properties, (object)values);
			Invoker.Method(this, "SetColumnProperties", paramsArray);
		}

		#endregion

       #region IEnumerable<NetOffice.VisioApi.IVDataColumn> Member
        
        /// <summary>
		/// SupportByVersionAttribute Visio, 12,14
		/// </summary>
		[SupportByVersionAttribute("Visio", 12,14)]
       public IEnumerator<NetOffice.VisioApi.IVDataColumn> GetEnumerator()  
       {
           NetRuntimeSystem.Collections.IEnumerable innerEnumerator = (this as NetRuntimeSystem.Collections.IEnumerable);
           foreach (NetOffice.VisioApi.IVDataColumn item in innerEnumerator)
               yield return item;
       }

       #endregion
          
		#region IEnumerable Members
       
		/// <summary>
		/// SupportByVersionAttribute Visio, 12,14
		/// </summary>
		[SupportByVersionAttribute("Visio", 12,14)]
		IEnumerator NetRuntimeSystem.Collections.IEnumerable.GetEnumerator()
		{
			return NetOffice.Utils.GetProxyEnumeratorAsProperty(this);
		}

		#endregion
		#pragma warning restore
	}
}