using System;
using NetRuntimeSystem = System;
using System.Runtime.InteropServices;
using System.Runtime.CompilerServices;
using System.ComponentModel;
using System.Reflection;
using System.Collections.Generic;
using System.Collections;
using NetOffice;
namespace NetOffice.OWC10Api
{
	///<summary>
	/// DispatchInterface RecordsetDefs 
	/// SupportByVersion OWC10, 1
	///</summary>
	[SupportByVersionAttribute("OWC10", 1)]
	[EntityTypeAttribute(EntityType.IsDispatchInterface)]
	public class RecordsetDefs : COMObject ,IEnumerable<NetOffice.OWC10Api.RecordsetDef>
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
                    _type = typeof(RecordsetDefs);
                    
                return _type;
            }
        }
        
        #endregion
        
		#region Construction

		///<param name="factory">current used factory core</param>
		///<param name="parentObject">object there has created the proxy</param>
        ///<param name="comProxy">inner wrapped COM proxy</param>
		public RecordsetDefs(Core factory, ICOMObject parentObject, object comProxy) : base(factory, parentObject, comProxy)
		{
			
		}

        ///<param name="parentObject">object there has created the proxy</param>
        ///<param name="comProxy">inner wrapped COM proxy</param>
        [EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public RecordsetDefs(ICOMObject parentObject, object comProxy) : base(parentObject, comProxy)
		{
		}
		
		///<param name="factory">current used factory core</param>
		///<param name="parentObject">object there has created the proxy</param>
        ///<param name="comProxy">inner wrapped COM proxy</param>
        ///<param name="comProxyType">Type of inner wrapped COM proxy"</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public RecordsetDefs(Core factory, ICOMObject parentObject, object comProxy, NetRuntimeSystem.Type comProxyType) : base(factory, parentObject, comProxy, comProxyType)
		{

		}

		///<param name="parentObject">object there has created the proxy</param>
        ///<param name="comProxy">inner wrapped COM proxy</param>
        ///<param name="comProxyType">Type of inner wrapped COM proxy"</param>
        [EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public RecordsetDefs(ICOMObject parentObject, object comProxy, NetRuntimeSystem.Type comProxyType) : base(parentObject, comProxy, comProxyType)
		{
		}
		
		///<param name="replacedObject">object to replaced. replacedObject are not usable after this action</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public RecordsetDefs(ICOMObject replacedObject) : base(replacedObject)
		{
		}
		
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public RecordsetDefs() : base()
		{
		}
		
		/// <param name="progId">registered ProgID</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public RecordsetDefs(string progId) : base(progId)
		{
		}
		
		#endregion
		
		#region Properties

		/// <summary>
		/// SupportByVersion OWC10 1
		/// Get
		/// </summary>
		[SupportByVersionAttribute("OWC10", 1)]
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
		/// SupportByVersion OWC10 1
		/// Get
		/// </summary>
		/// <param name="index">object Index</param>
		[SupportByVersionAttribute("OWC10", 1)]
		[NetRuntimeSystem.Runtime.CompilerServices.IndexerName("Item")]
		public NetOffice.OWC10Api.RecordsetDef this[object index]
		{
			get
{			
			object[] paramsArray = Invoker.ValidateParamsArray(index);
			object returnItem = Invoker.PropertyGet(this, "Item", paramsArray);
			NetOffice.OWC10Api.RecordsetDef newObject = Factory.CreateKnownObjectFromComProxy(this,returnItem,NetOffice.OWC10Api.RecordsetDef.LateBindingApiWrapperType) as NetOffice.OWC10Api.RecordsetDef;
			return newObject;
			}
		}

		#endregion

		#region Methods

		/// <summary>
		/// SupportByVersion OWC10 1
		/// 
		/// </summary>
		/// <param name="schemaRowsource">object SchemaRowsource</param>
		/// <param name="name">optional object Name</param>
		[SupportByVersionAttribute("OWC10", 1)]
		public NetOffice.OWC10Api.RecordsetDef Add(object schemaRowsource, object name)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(schemaRowsource, name);
			object returnItem = Invoker.MethodReturn(this, "Add", paramsArray);
			NetOffice.OWC10Api.RecordsetDef newObject = Factory.CreateKnownObjectFromComProxy(this, returnItem,NetOffice.OWC10Api.RecordsetDef.LateBindingApiWrapperType) as NetOffice.OWC10Api.RecordsetDef;
			return newObject;
		}

		/// <summary>
		/// SupportByVersion OWC10 1
		/// 
		/// </summary>
		/// <param name="schemaRowsource">object SchemaRowsource</param>
		[CustomMethodAttribute]
		[SupportByVersionAttribute("OWC10", 1)]
		public NetOffice.OWC10Api.RecordsetDef Add(object schemaRowsource)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(schemaRowsource);
			object returnItem = Invoker.MethodReturn(this, "Add", paramsArray);
			NetOffice.OWC10Api.RecordsetDef newObject = Factory.CreateKnownObjectFromComProxy(this, returnItem,NetOffice.OWC10Api.RecordsetDef.LateBindingApiWrapperType) as NetOffice.OWC10Api.RecordsetDef;
			return newObject;
		}

		/// <summary>
		/// SupportByVersion OWC10 1
		/// 
		/// </summary>
		/// <param name="source">string Source</param>
		/// <param name="rowsourceType">optional object RowsourceType</param>
		/// <param name="name">optional object Name</param>
		[SupportByVersionAttribute("OWC10", 1)]
		public NetOffice.OWC10Api.RecordsetDef AddNew(string source, object rowsourceType, object name)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(source, rowsourceType, name);
			object returnItem = Invoker.MethodReturn(this, "AddNew", paramsArray);
			NetOffice.OWC10Api.RecordsetDef newObject = Factory.CreateKnownObjectFromComProxy(this, returnItem,NetOffice.OWC10Api.RecordsetDef.LateBindingApiWrapperType) as NetOffice.OWC10Api.RecordsetDef;
			return newObject;
		}

		/// <summary>
		/// SupportByVersion OWC10 1
		/// 
		/// </summary>
		/// <param name="source">string Source</param>
		[CustomMethodAttribute]
		[SupportByVersionAttribute("OWC10", 1)]
		public NetOffice.OWC10Api.RecordsetDef AddNew(string source)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(source);
			object returnItem = Invoker.MethodReturn(this, "AddNew", paramsArray);
			NetOffice.OWC10Api.RecordsetDef newObject = Factory.CreateKnownObjectFromComProxy(this, returnItem,NetOffice.OWC10Api.RecordsetDef.LateBindingApiWrapperType) as NetOffice.OWC10Api.RecordsetDef;
			return newObject;
		}

		/// <summary>
		/// SupportByVersion OWC10 1
		/// 
		/// </summary>
		/// <param name="source">string Source</param>
		/// <param name="rowsourceType">optional object RowsourceType</param>
		[CustomMethodAttribute]
		[SupportByVersionAttribute("OWC10", 1)]
		public NetOffice.OWC10Api.RecordsetDef AddNew(string source, object rowsourceType)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(source, rowsourceType);
			object returnItem = Invoker.MethodReturn(this, "AddNew", paramsArray);
			NetOffice.OWC10Api.RecordsetDef newObject = Factory.CreateKnownObjectFromComProxy(this, returnItem,NetOffice.OWC10Api.RecordsetDef.LateBindingApiWrapperType) as NetOffice.OWC10Api.RecordsetDef;
			return newObject;
		}

		/// <summary>
		/// SupportByVersion OWC10 1
		/// 
		/// </summary>
		/// <param name="index">object Index</param>
		[SupportByVersionAttribute("OWC10", 1)]
		public void Delete(object index)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(index);
			Invoker.Method(this, "Delete", paramsArray);
		}

		#endregion

       #region IEnumerable<NetOffice.OWC10Api.RecordsetDef> Member
        
        /// <summary>
		/// SupportByVersionAttribute OWC10, 1
		/// </summary>
		[SupportByVersionAttribute("OWC10", 1)]
       public IEnumerator<NetOffice.OWC10Api.RecordsetDef> GetEnumerator()  
       {
           NetRuntimeSystem.Collections.IEnumerable innerEnumerator = (this as NetRuntimeSystem.Collections.IEnumerable);
           foreach (NetOffice.OWC10Api.RecordsetDef item in innerEnumerator)
               yield return item;
       }

       #endregion
          
		#region IEnumerable Members
       
		/// <summary>
		/// SupportByVersionAttribute OWC10, 1
		/// </summary>
		[SupportByVersionAttribute("OWC10", 1)]
		IEnumerator NetRuntimeSystem.Collections.IEnumerable.GetEnumerator()
		{
			return NetOffice.Utils.GetProxyEnumeratorAsProperty(this);
		}

		#endregion
		#pragma warning restore
	}
}