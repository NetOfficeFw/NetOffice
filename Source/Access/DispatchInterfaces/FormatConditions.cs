using System;
using NetRuntimeSystem = System;
using System.Runtime.InteropServices;
using System.Runtime.CompilerServices;
using System.ComponentModel;
using System.Reflection;
using System.Collections.Generic;
using System.Collections;
using NetOffice;
namespace NetOffice.AccessApi
{
	///<summary>
	/// DispatchInterface FormatConditions 
	/// SupportByVersion Access, 9,10,11,12,14
	///</summary>
	[SupportByVersionAttribute("Access", 9,10,11,12,14)]
	[EntityTypeAttribute(EntityType.IsDispatchInterface)]
	public class FormatConditions : COMObject ,IEnumerable<NetOffice.AccessApi._FormatCondition>
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
                    _type = typeof(FormatConditions);
                    
                return _type;
            }
        }
        
        #endregion
        
		#region Construction

        /// <param name="parentObject">object there has created the proxy</param>
        /// <param name="comProxy">inner wrapped COM proxy</param>
        [EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public FormatConditions(COMObject parentObject, object comProxy) : base(parentObject, comProxy)
		{
		}
		
		/// <param name="parentObject">object there has created the proxy</param>
        /// <param name="comProxy">inner wrapped COM proxy</param>
        /// <param name="comProxyType">Type of inner wrapped COM proxy"</param>
        [EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public FormatConditions(COMObject parentObject, object comProxy, NetRuntimeSystem.Type comProxyType) : base(parentObject, comProxy, comProxyType)
		{
		}
		
		/// <param name="replacedObject">object to replaced. replacedObject are not usable after this action</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public FormatConditions(COMObject replacedObject) : base(replacedObject)
		{
		}
		
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public FormatConditions() : base()
		{
		}
		
		/// <param name="progId">registered ProgID</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public FormatConditions(string progId) : base(progId)
		{
		}
		
		#endregion
		
		#region Properties

		/// <summary>
		/// SupportByVersion Access 9, 10, 11, 12, 14
		/// Get
		/// </summary>
		[SupportByVersionAttribute("Access", 9,10,11,12,14)]
		public NetOffice.AccessApi.Application Application
		{
			get
			{
				object[] paramsArray = null;
				object returnItem = Invoker.PropertyGet(this, "Application", paramsArray);
				NetOffice.AccessApi.Application newObject = NetOffice.Factory.CreateKnownObjectFromComProxy(this,returnItem,NetOffice.AccessApi.Application.LateBindingApiWrapperType) as NetOffice.AccessApi.Application;
				return newObject;
			}
		}

		/// <summary>
		/// SupportByVersion Access 9, 10, 11, 12, 14
		/// Get
		/// Unknown COM Proxy
		/// </summary>
		[SupportByVersionAttribute("Access", 9,10,11,12,14)]
		public object Parent
		{
			get
			{
				object[] paramsArray = null;
				object returnItem = Invoker.PropertyGet(this, "Parent", paramsArray);
				COMObject newObject = NetOffice.Factory.CreateObjectFromComProxy(this,returnItem);
				return newObject;
			}
		}

		/// <summary>
		/// SupportByVersion Access 9, 10, 11, 12, 14
		/// Get
		/// </summary>
		/// <param name="index">object Index</param>
		[SupportByVersionAttribute("Access", 9,10,11,12,14)]
		[NetRuntimeSystem.Runtime.CompilerServices.IndexerName("Item")]
		public NetOffice.AccessApi._FormatCondition this[object index]
		{
			get
{			
			object[] paramsArray = Invoker.ValidateParamsArray(index);
			object returnItem = Invoker.PropertyGet(this, "Item", paramsArray);
			NetOffice.AccessApi._FormatCondition newObject = NetOffice.Factory.CreateObjectFromComProxy(this,returnItem) as NetOffice.AccessApi._FormatCondition;
			return newObject;
			}
		}

		/// <summary>
		/// SupportByVersion Access 9, 10, 11, 12, 14
		/// Get
		/// </summary>
		[SupportByVersionAttribute("Access", 9,10,11,12,14)]
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
		/// SupportByVersion Access 9, 10, 11, 12, 14
		/// </summary>
		/// <param name="type">NetOffice.AccessApi.Enums.AcFormatConditionType Type</param>
		/// <param name="_operator">optional NetOffice.AccessApi.Enums.AcFormatConditionOperator Operator = 0</param>
		/// <param name="expression1">optional object Expression1</param>
		/// <param name="expression2">optional object Expression2</param>
		[SupportByVersionAttribute("Access", 9,10,11,12,14)]
		public NetOffice.AccessApi._FormatCondition Add(NetOffice.AccessApi.Enums.AcFormatConditionType type, NetOffice.AccessApi.Enums.AcFormatConditionOperator _operator, object expression1, object expression2)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(type, _operator, expression1, expression2);
			object returnItem = Invoker.MethodReturn(this, "Add", paramsArray);
			NetOffice.AccessApi._FormatCondition newObject = NetOffice.Factory.CreateObjectFromComProxy(this,returnItem) as NetOffice.AccessApi._FormatCondition;
			return newObject;
		}

		/// <summary>
		/// SupportByVersion Access 9, 10, 11, 12, 14
		/// </summary>
		/// <param name="type">NetOffice.AccessApi.Enums.AcFormatConditionType Type</param>
		[CustomMethodAttribute]
		[SupportByVersionAttribute("Access", 9,10,11,12,14)]
		public NetOffice.AccessApi._FormatCondition Add(NetOffice.AccessApi.Enums.AcFormatConditionType type)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(type);
			object returnItem = Invoker.MethodReturn(this, "Add", paramsArray);
			NetOffice.AccessApi._FormatCondition newObject = NetOffice.Factory.CreateObjectFromComProxy(this,returnItem) as NetOffice.AccessApi._FormatCondition;
			return newObject;
		}

		/// <summary>
		/// SupportByVersion Access 9, 10, 11, 12, 14
		/// </summary>
		/// <param name="type">NetOffice.AccessApi.Enums.AcFormatConditionType Type</param>
		/// <param name="_operator">optional NetOffice.AccessApi.Enums.AcFormatConditionOperator Operator = 0</param>
		[CustomMethodAttribute]
		[SupportByVersionAttribute("Access", 9,10,11,12,14)]
		public NetOffice.AccessApi._FormatCondition Add(NetOffice.AccessApi.Enums.AcFormatConditionType type, NetOffice.AccessApi.Enums.AcFormatConditionOperator _operator)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(type, _operator);
			object returnItem = Invoker.MethodReturn(this, "Add", paramsArray);
			NetOffice.AccessApi._FormatCondition newObject = NetOffice.Factory.CreateObjectFromComProxy(this,returnItem) as NetOffice.AccessApi._FormatCondition;
			return newObject;
		}

		/// <summary>
		/// SupportByVersion Access 9, 10, 11, 12, 14
		/// </summary>
		/// <param name="type">NetOffice.AccessApi.Enums.AcFormatConditionType Type</param>
		/// <param name="_operator">optional NetOffice.AccessApi.Enums.AcFormatConditionOperator Operator = 0</param>
		/// <param name="expression1">optional object Expression1</param>
		[CustomMethodAttribute]
		[SupportByVersionAttribute("Access", 9,10,11,12,14)]
		public NetOffice.AccessApi._FormatCondition Add(NetOffice.AccessApi.Enums.AcFormatConditionType type, NetOffice.AccessApi.Enums.AcFormatConditionOperator _operator, object expression1)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(type, _operator, expression1);
			object returnItem = Invoker.MethodReturn(this, "Add", paramsArray);
			NetOffice.AccessApi._FormatCondition newObject = NetOffice.Factory.CreateObjectFromComProxy(this,returnItem) as NetOffice.AccessApi._FormatCondition;
			return newObject;
		}

		/// <summary>
		/// SupportByVersion Access 9, 10, 11, 12, 14
		/// </summary>
		[SupportByVersionAttribute("Access", 9,10,11,12,14)]
		public void Delete()
		{
			object[] paramsArray = null;
			Invoker.Method(this, "Delete", paramsArray);
		}

		/// <summary>
		/// SupportByVersion Access 11, 12, 14
		/// </summary>
		/// <param name="dispid">Int32 dispid</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		[SupportByVersionAttribute("Access", 11,12,14)]
		public bool IsMemberSafe(Int32 dispid)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(dispid);
			object returnItem = Invoker.MethodReturn(this, "IsMemberSafe", paramsArray);
			return (bool)returnItem;
		}

		#endregion

       #region IEnumerable<NetOffice.AccessApi._FormatCondition> Member
        
        /// <summary>
		/// SupportByVersionAttribute Access, 9,10,11,12,14
		/// </summary>
		[SupportByVersionAttribute("Access", 9,10,11,12,14)]
       public IEnumerator<NetOffice.AccessApi._FormatCondition> GetEnumerator()  
       {
           NetRuntimeSystem.Collections.IEnumerable innerEnumerator = (this as NetRuntimeSystem.Collections.IEnumerable);
           foreach (NetOffice.AccessApi._FormatCondition item in innerEnumerator)
               yield return item;
       }

       #endregion
          
		#region IEnumerable Members
       
		/// <summary>
		/// SupportByVersionAttribute Access, 9,10,11,12,14
		/// </summary>
		[SupportByVersionAttribute("Access", 9,10,11,12,14)]
		IEnumerator NetRuntimeSystem.Collections.IEnumerable.GetEnumerator()
		{
			return NetOffice.Utils.GetProxyEnumeratorAsMethod(this);
		}

		#endregion
		#pragma warning restore
	}
}