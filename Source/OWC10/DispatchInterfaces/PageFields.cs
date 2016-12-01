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
	/// DispatchInterface PageFields 
	/// SupportByVersion OWC10, 1
	///</summary>
	[SupportByVersionAttribute("OWC10", 1)]
	[EntityTypeAttribute(EntityType.IsDispatchInterface)]
	public class PageFields : COMObject ,IEnumerable<NetOffice.OWC10Api.PageField>
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
                    _type = typeof(PageFields);
                    
                return _type;
            }
        }
        
        #endregion
        
		#region Construction

		///<param name="factory">current used factory core</param>
		///<param name="parentObject">object there has created the proxy</param>
        ///<param name="comProxy">inner wrapped COM proxy</param>
		public PageFields(Core factory, ICOMObject parentObject, object comProxy) : base(factory, parentObject, comProxy)
		{
			
		}

        ///<param name="parentObject">object there has created the proxy</param>
        ///<param name="comProxy">inner wrapped COM proxy</param>
        [EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public PageFields(ICOMObject parentObject, object comProxy) : base(parentObject, comProxy)
		{
		}
		
		///<param name="factory">current used factory core</param>
		///<param name="parentObject">object there has created the proxy</param>
        ///<param name="comProxy">inner wrapped COM proxy</param>
        ///<param name="comProxyType">Type of inner wrapped COM proxy"</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public PageFields(Core factory, ICOMObject parentObject, object comProxy, NetRuntimeSystem.Type comProxyType) : base(factory, parentObject, comProxy, comProxyType)
		{

		}

		///<param name="parentObject">object there has created the proxy</param>
        ///<param name="comProxy">inner wrapped COM proxy</param>
        ///<param name="comProxyType">Type of inner wrapped COM proxy"</param>
        [EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public PageFields(ICOMObject parentObject, object comProxy, NetRuntimeSystem.Type comProxyType) : base(parentObject, comProxy, comProxyType)
		{
		}
		
		///<param name="replacedObject">object to replaced. replacedObject are not usable after this action</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public PageFields(ICOMObject replacedObject) : base(replacedObject)
		{
		}
		
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public PageFields() : base()
		{
		}
		
		/// <param name="progId">registered ProgID</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public PageFields(string progId) : base(progId)
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
		public NetOffice.OWC10Api.PageField this[object index]
		{
			get
{			
			object[] paramsArray = Invoker.ValidateParamsArray(index);
			object returnItem = Invoker.PropertyGet(this, "Item", paramsArray);
			NetOffice.OWC10Api.PageField newObject = Factory.CreateKnownObjectFromComProxy(this,returnItem,NetOffice.OWC10Api.PageField.LateBindingApiWrapperType) as NetOffice.OWC10Api.PageField;
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
		/// <param name="index">object Index</param>
		[SupportByVersionAttribute("OWC10", 1)]
		public void Delete(object index)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(index);
			Invoker.Method(this, "Delete", paramsArray);
		}

		/// <summary>
		/// SupportByVersion OWC10 1
		/// 
		/// </summary>
		/// <param name="source">object Source</param>
		/// <param name="fieldType">optional object FieldType</param>
		/// <param name="name">optional object Name</param>
		/// <param name="totalType">optional NetOffice.OWC10Api.Enums.DscTotalTypeEnum TotalType = 0</param>
		/// <param name="index">optional object Index</param>
		[SupportByVersionAttribute("OWC10", 1)]
		public NetOffice.OWC10Api.PageField Add(object source, object fieldType, object name, object totalType, object index)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(source, fieldType, name, totalType, index);
			object returnItem = Invoker.MethodReturn(this, "Add", paramsArray);
			NetOffice.OWC10Api.PageField newObject = Factory.CreateKnownObjectFromComProxy(this, returnItem,NetOffice.OWC10Api.PageField.LateBindingApiWrapperType) as NetOffice.OWC10Api.PageField;
			return newObject;
		}

		/// <summary>
		/// SupportByVersion OWC10 1
		/// 
		/// </summary>
		/// <param name="source">object Source</param>
		[CustomMethodAttribute]
		[SupportByVersionAttribute("OWC10", 1)]
		public NetOffice.OWC10Api.PageField Add(object source)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(source);
			object returnItem = Invoker.MethodReturn(this, "Add", paramsArray);
			NetOffice.OWC10Api.PageField newObject = Factory.CreateKnownObjectFromComProxy(this, returnItem,NetOffice.OWC10Api.PageField.LateBindingApiWrapperType) as NetOffice.OWC10Api.PageField;
			return newObject;
		}

		/// <summary>
		/// SupportByVersion OWC10 1
		/// 
		/// </summary>
		/// <param name="source">object Source</param>
		/// <param name="fieldType">optional object FieldType</param>
		[CustomMethodAttribute]
		[SupportByVersionAttribute("OWC10", 1)]
		public NetOffice.OWC10Api.PageField Add(object source, object fieldType)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(source, fieldType);
			object returnItem = Invoker.MethodReturn(this, "Add", paramsArray);
			NetOffice.OWC10Api.PageField newObject = Factory.CreateKnownObjectFromComProxy(this, returnItem,NetOffice.OWC10Api.PageField.LateBindingApiWrapperType) as NetOffice.OWC10Api.PageField;
			return newObject;
		}

		/// <summary>
		/// SupportByVersion OWC10 1
		/// 
		/// </summary>
		/// <param name="source">object Source</param>
		/// <param name="fieldType">optional object FieldType</param>
		/// <param name="name">optional object Name</param>
		[CustomMethodAttribute]
		[SupportByVersionAttribute("OWC10", 1)]
		public NetOffice.OWC10Api.PageField Add(object source, object fieldType, object name)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(source, fieldType, name);
			object returnItem = Invoker.MethodReturn(this, "Add", paramsArray);
			NetOffice.OWC10Api.PageField newObject = Factory.CreateKnownObjectFromComProxy(this, returnItem,NetOffice.OWC10Api.PageField.LateBindingApiWrapperType) as NetOffice.OWC10Api.PageField;
			return newObject;
		}

		/// <summary>
		/// SupportByVersion OWC10 1
		/// 
		/// </summary>
		/// <param name="source">object Source</param>
		/// <param name="fieldType">optional object FieldType</param>
		/// <param name="name">optional object Name</param>
		/// <param name="totalType">optional NetOffice.OWC10Api.Enums.DscTotalTypeEnum TotalType = 0</param>
		[CustomMethodAttribute]
		[SupportByVersionAttribute("OWC10", 1)]
		public NetOffice.OWC10Api.PageField Add(object source, object fieldType, object name, object totalType)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(source, fieldType, name, totalType);
			object returnItem = Invoker.MethodReturn(this, "Add", paramsArray);
			NetOffice.OWC10Api.PageField newObject = Factory.CreateKnownObjectFromComProxy(this, returnItem,NetOffice.OWC10Api.PageField.LateBindingApiWrapperType) as NetOffice.OWC10Api.PageField;
			return newObject;
		}

		/// <summary>
		/// SupportByVersion OWC10 1
		/// 
		/// </summary>
		/// <param name="source">object Source</param>
		/// <param name="fieldType">optional object FieldType</param>
		/// <param name="name">optional object Name</param>
		/// <param name="totalType">optional NetOffice.OWC10Api.Enums.DscTotalTypeEnum TotalType = 0</param>
		/// <param name="index">optional object Index</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		[SupportByVersionAttribute("OWC10", 1)]
		public NetOffice.OWC10Api.PageField AddBroken(object source, object fieldType, object name, object totalType, object index)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(source, fieldType, name, totalType, index);
			object returnItem = Invoker.MethodReturn(this, "AddBroken", paramsArray);
			NetOffice.OWC10Api.PageField newObject = Factory.CreateKnownObjectFromComProxy(this, returnItem,NetOffice.OWC10Api.PageField.LateBindingApiWrapperType) as NetOffice.OWC10Api.PageField;
			return newObject;
		}

		/// <summary>
		/// SupportByVersion OWC10 1
		/// 
		/// </summary>
		/// <param name="source">object Source</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		[CustomMethodAttribute]
		[SupportByVersionAttribute("OWC10", 1)]
		public NetOffice.OWC10Api.PageField AddBroken(object source)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(source);
			object returnItem = Invoker.MethodReturn(this, "AddBroken", paramsArray);
			NetOffice.OWC10Api.PageField newObject = Factory.CreateKnownObjectFromComProxy(this, returnItem,NetOffice.OWC10Api.PageField.LateBindingApiWrapperType) as NetOffice.OWC10Api.PageField;
			return newObject;
		}

		/// <summary>
		/// SupportByVersion OWC10 1
		/// 
		/// </summary>
		/// <param name="source">object Source</param>
		/// <param name="fieldType">optional object FieldType</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		[CustomMethodAttribute]
		[SupportByVersionAttribute("OWC10", 1)]
		public NetOffice.OWC10Api.PageField AddBroken(object source, object fieldType)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(source, fieldType);
			object returnItem = Invoker.MethodReturn(this, "AddBroken", paramsArray);
			NetOffice.OWC10Api.PageField newObject = Factory.CreateKnownObjectFromComProxy(this, returnItem,NetOffice.OWC10Api.PageField.LateBindingApiWrapperType) as NetOffice.OWC10Api.PageField;
			return newObject;
		}

		/// <summary>
		/// SupportByVersion OWC10 1
		/// 
		/// </summary>
		/// <param name="source">object Source</param>
		/// <param name="fieldType">optional object FieldType</param>
		/// <param name="name">optional object Name</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		[CustomMethodAttribute]
		[SupportByVersionAttribute("OWC10", 1)]
		public NetOffice.OWC10Api.PageField AddBroken(object source, object fieldType, object name)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(source, fieldType, name);
			object returnItem = Invoker.MethodReturn(this, "AddBroken", paramsArray);
			NetOffice.OWC10Api.PageField newObject = Factory.CreateKnownObjectFromComProxy(this, returnItem,NetOffice.OWC10Api.PageField.LateBindingApiWrapperType) as NetOffice.OWC10Api.PageField;
			return newObject;
		}

		/// <summary>
		/// SupportByVersion OWC10 1
		/// 
		/// </summary>
		/// <param name="source">object Source</param>
		/// <param name="fieldType">optional object FieldType</param>
		/// <param name="name">optional object Name</param>
		/// <param name="totalType">optional NetOffice.OWC10Api.Enums.DscTotalTypeEnum TotalType = 0</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		[CustomMethodAttribute]
		[SupportByVersionAttribute("OWC10", 1)]
		public NetOffice.OWC10Api.PageField AddBroken(object source, object fieldType, object name, object totalType)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(source, fieldType, name, totalType);
			object returnItem = Invoker.MethodReturn(this, "AddBroken", paramsArray);
			NetOffice.OWC10Api.PageField newObject = Factory.CreateKnownObjectFromComProxy(this, returnItem,NetOffice.OWC10Api.PageField.LateBindingApiWrapperType) as NetOffice.OWC10Api.PageField;
			return newObject;
		}

		#endregion

       #region IEnumerable<NetOffice.OWC10Api.PageField> Member
        
        /// <summary>
		/// SupportByVersionAttribute OWC10, 1
		/// </summary>
		[SupportByVersionAttribute("OWC10", 1)]
       public IEnumerator<NetOffice.OWC10Api.PageField> GetEnumerator()  
       {
           NetRuntimeSystem.Collections.IEnumerable innerEnumerator = (this as NetRuntimeSystem.Collections.IEnumerable);
           foreach (NetOffice.OWC10Api.PageField item in innerEnumerator)
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