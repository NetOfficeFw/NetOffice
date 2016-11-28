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
	/// DispatchInterface SchemaParameters 
	/// SupportByVersion OWC10, 1
	///</summary>
	[SupportByVersionAttribute("OWC10", 1)]
	[EntityTypeAttribute(EntityType.IsDispatchInterface)]
	public class SchemaParameters : COMObject ,IEnumerable<NetOffice.OWC10Api.SchemaParameter>
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
                    _type = typeof(SchemaParameters);
                    
                return _type;
            }
        }
        
        #endregion
        
		#region Construction

		///<param name="factory">current used factory core</param>
		///<param name="parentObject">object there has created the proxy</param>
        ///<param name="comProxy">inner wrapped COM proxy</param>
		public SchemaParameters(Core factory, ICOMObject parentObject, object comProxy) : base(factory, parentObject, comProxy)
		{
			
		}

        ///<param name="parentObject">object there has created the proxy</param>
        ///<param name="comProxy">inner wrapped COM proxy</param>
        [EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public SchemaParameters(ICOMObject parentObject, object comProxy) : base(parentObject, comProxy)
		{
		}
		
		///<param name="factory">current used factory core</param>
		///<param name="parentObject">object there has created the proxy</param>
        ///<param name="comProxy">inner wrapped COM proxy</param>
        ///<param name="comProxyType">Type of inner wrapped COM proxy"</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public SchemaParameters(Core factory, ICOMObject parentObject, object comProxy, NetRuntimeSystem.Type comProxyType) : base(factory, parentObject, comProxy, comProxyType)
		{

		}

		///<param name="parentObject">object there has created the proxy</param>
        ///<param name="comProxy">inner wrapped COM proxy</param>
        ///<param name="comProxyType">Type of inner wrapped COM proxy"</param>
        [EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public SchemaParameters(ICOMObject parentObject, object comProxy, NetRuntimeSystem.Type comProxyType) : base(parentObject, comProxy, comProxyType)
		{
		}
		
		///<param name="replacedObject">object to replaced. replacedObject are not usable after this action</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public SchemaParameters(ICOMObject replacedObject) : base(replacedObject)
		{
		}
		
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public SchemaParameters() : base()
		{
		}
		
		/// <param name="progId">registered ProgID</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public SchemaParameters(string progId) : base(progId)
		{
		}
		
		#endregion
		
		#region Properties

		/// <summary>
		/// SupportByVersion OWC10 1
		/// Get
		/// </summary>
		/// <param name="index">object Index</param>
		[SupportByVersionAttribute("OWC10", 1)]
		[NetRuntimeSystem.Runtime.CompilerServices.IndexerName("Item")]
		public NetOffice.OWC10Api.SchemaParameter this[object index]
		{
			get
{			
			object[] paramsArray = Invoker.ValidateParamsArray(index);
			object returnItem = Invoker.PropertyGet(this, "Item", paramsArray);
			NetOffice.OWC10Api.SchemaParameter newObject = Factory.CreateKnownObjectFromComProxy(this,returnItem,NetOffice.OWC10Api.SchemaParameter.LateBindingApiWrapperType) as NetOffice.OWC10Api.SchemaParameter;
			return newObject;
			}
		}

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

		#endregion

		#region Methods

		/// <summary>
		/// SupportByVersion OWC10 1
		/// 
		/// </summary>
		/// <param name="name">string Name</param>
		/// <param name="dataType">optional object DataType</param>
		/// <param name="size">optional object Size</param>
		/// <param name="scale">optional object Scale</param>
		/// <param name="precision">optional object Precision</param>
		/// <param name="direction">optional object Direction</param>
		[SupportByVersionAttribute("OWC10", 1)]
		public NetOffice.OWC10Api.SchemaParameter Add(string name, object dataType, object size, object scale, object precision, object direction)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(name, dataType, size, scale, precision, direction);
			object returnItem = Invoker.MethodReturn(this, "Add", paramsArray);
			NetOffice.OWC10Api.SchemaParameter newObject = Factory.CreateKnownObjectFromComProxy(this, returnItem,NetOffice.OWC10Api.SchemaParameter.LateBindingApiWrapperType) as NetOffice.OWC10Api.SchemaParameter;
			return newObject;
		}

		/// <summary>
		/// SupportByVersion OWC10 1
		/// 
		/// </summary>
		/// <param name="name">string Name</param>
		[CustomMethodAttribute]
		[SupportByVersionAttribute("OWC10", 1)]
		public NetOffice.OWC10Api.SchemaParameter Add(string name)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(name);
			object returnItem = Invoker.MethodReturn(this, "Add", paramsArray);
			NetOffice.OWC10Api.SchemaParameter newObject = Factory.CreateKnownObjectFromComProxy(this, returnItem,NetOffice.OWC10Api.SchemaParameter.LateBindingApiWrapperType) as NetOffice.OWC10Api.SchemaParameter;
			return newObject;
		}

		/// <summary>
		/// SupportByVersion OWC10 1
		/// 
		/// </summary>
		/// <param name="name">string Name</param>
		/// <param name="dataType">optional object DataType</param>
		[CustomMethodAttribute]
		[SupportByVersionAttribute("OWC10", 1)]
		public NetOffice.OWC10Api.SchemaParameter Add(string name, object dataType)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(name, dataType);
			object returnItem = Invoker.MethodReturn(this, "Add", paramsArray);
			NetOffice.OWC10Api.SchemaParameter newObject = Factory.CreateKnownObjectFromComProxy(this, returnItem,NetOffice.OWC10Api.SchemaParameter.LateBindingApiWrapperType) as NetOffice.OWC10Api.SchemaParameter;
			return newObject;
		}

		/// <summary>
		/// SupportByVersion OWC10 1
		/// 
		/// </summary>
		/// <param name="name">string Name</param>
		/// <param name="dataType">optional object DataType</param>
		/// <param name="size">optional object Size</param>
		[CustomMethodAttribute]
		[SupportByVersionAttribute("OWC10", 1)]
		public NetOffice.OWC10Api.SchemaParameter Add(string name, object dataType, object size)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(name, dataType, size);
			object returnItem = Invoker.MethodReturn(this, "Add", paramsArray);
			NetOffice.OWC10Api.SchemaParameter newObject = Factory.CreateKnownObjectFromComProxy(this, returnItem,NetOffice.OWC10Api.SchemaParameter.LateBindingApiWrapperType) as NetOffice.OWC10Api.SchemaParameter;
			return newObject;
		}

		/// <summary>
		/// SupportByVersion OWC10 1
		/// 
		/// </summary>
		/// <param name="name">string Name</param>
		/// <param name="dataType">optional object DataType</param>
		/// <param name="size">optional object Size</param>
		/// <param name="scale">optional object Scale</param>
		[CustomMethodAttribute]
		[SupportByVersionAttribute("OWC10", 1)]
		public NetOffice.OWC10Api.SchemaParameter Add(string name, object dataType, object size, object scale)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(name, dataType, size, scale);
			object returnItem = Invoker.MethodReturn(this, "Add", paramsArray);
			NetOffice.OWC10Api.SchemaParameter newObject = Factory.CreateKnownObjectFromComProxy(this, returnItem,NetOffice.OWC10Api.SchemaParameter.LateBindingApiWrapperType) as NetOffice.OWC10Api.SchemaParameter;
			return newObject;
		}

		/// <summary>
		/// SupportByVersion OWC10 1
		/// 
		/// </summary>
		/// <param name="name">string Name</param>
		/// <param name="dataType">optional object DataType</param>
		/// <param name="size">optional object Size</param>
		/// <param name="scale">optional object Scale</param>
		/// <param name="precision">optional object Precision</param>
		[CustomMethodAttribute]
		[SupportByVersionAttribute("OWC10", 1)]
		public NetOffice.OWC10Api.SchemaParameter Add(string name, object dataType, object size, object scale, object precision)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(name, dataType, size, scale, precision);
			object returnItem = Invoker.MethodReturn(this, "Add", paramsArray);
			NetOffice.OWC10Api.SchemaParameter newObject = Factory.CreateKnownObjectFromComProxy(this, returnItem,NetOffice.OWC10Api.SchemaParameter.LateBindingApiWrapperType) as NetOffice.OWC10Api.SchemaParameter;
			return newObject;
		}

		#endregion

       #region IEnumerable<NetOffice.OWC10Api.SchemaParameter> Member
        
        /// <summary>
		/// SupportByVersionAttribute OWC10, 1
		/// </summary>
		[SupportByVersionAttribute("OWC10", 1)]
       public IEnumerator<NetOffice.OWC10Api.SchemaParameter> GetEnumerator()  
       {
           NetRuntimeSystem.Collections.IEnumerable innerEnumerator = (this as NetRuntimeSystem.Collections.IEnumerable);
           foreach (NetOffice.OWC10Api.SchemaParameter item in innerEnumerator)
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