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
	/// DispatchInterface GroupingDefs 
	/// SupportByVersion OWC10, 1
	///</summary>
	[SupportByVersionAttribute("OWC10", 1)]
	[EntityTypeAttribute(EntityType.IsDispatchInterface)]
	public class GroupingDefs : COMObject ,IEnumerable<NetOffice.OWC10Api.GroupingDef>
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
                    _type = typeof(GroupingDefs);
                    
                return _type;
            }
        }
        
        #endregion
        
		#region Construction

		///<param name="factory">current used factory core</param>
		///<param name="parentObject">object there has created the proxy</param>
        ///<param name="comProxy">inner wrapped COM proxy</param>
		public GroupingDefs(Core factory, ICOMObject parentObject, object comProxy) : base(factory, parentObject, comProxy)
		{
			
		}

        ///<param name="parentObject">object there has created the proxy</param>
        ///<param name="comProxy">inner wrapped COM proxy</param>
        [EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public GroupingDefs(ICOMObject parentObject, object comProxy) : base(parentObject, comProxy)
		{
		}
		
		///<param name="factory">current used factory core</param>
		///<param name="parentObject">object there has created the proxy</param>
        ///<param name="comProxy">inner wrapped COM proxy</param>
        ///<param name="comProxyType">Type of inner wrapped COM proxy"</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public GroupingDefs(Core factory, ICOMObject parentObject, object comProxy, NetRuntimeSystem.Type comProxyType) : base(factory, parentObject, comProxy, comProxyType)
		{

		}

		///<param name="parentObject">object there has created the proxy</param>
        ///<param name="comProxy">inner wrapped COM proxy</param>
        ///<param name="comProxyType">Type of inner wrapped COM proxy"</param>
        [EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public GroupingDefs(ICOMObject parentObject, object comProxy, NetRuntimeSystem.Type comProxyType) : base(parentObject, comProxy, comProxyType)
		{
		}
		
		///<param name="replacedObject">object to replaced. replacedObject are not usable after this action</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public GroupingDefs(ICOMObject replacedObject) : base(replacedObject)
		{
		}
		
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public GroupingDefs() : base()
		{
		}
		
		/// <param name="progId">registered ProgID</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public GroupingDefs(string progId) : base(progId)
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
		public NetOffice.OWC10Api.GroupingDef this[object index]
		{
			get
{			
			object[] paramsArray = Invoker.ValidateParamsArray(index);
			object returnItem = Invoker.PropertyGet(this, "Item", paramsArray);
			NetOffice.OWC10Api.GroupingDef newObject = Factory.CreateKnownObjectFromComProxy(this,returnItem,NetOffice.OWC10Api.GroupingDef.LateBindingApiWrapperType) as NetOffice.OWC10Api.GroupingDef;
			return newObject;
			}
		}

		#endregion

		#region Methods

		/// <summary>
		/// SupportByVersion OWC10 1
		/// 
		/// </summary>
		/// <param name="groupingDefName">string GroupingDefName</param>
		/// <param name="groupingFieldName">string GroupingFieldName</param>
		/// <param name="pageFieldName">string PageFieldName</param>
		/// <param name="index">optional object Index</param>
		[SupportByVersionAttribute("OWC10", 1)]
		public NetOffice.OWC10Api.GroupingDef Add(string groupingDefName, string groupingFieldName, string pageFieldName, object index)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(groupingDefName, groupingFieldName, pageFieldName, index);
			object returnItem = Invoker.MethodReturn(this, "Add", paramsArray);
			NetOffice.OWC10Api.GroupingDef newObject = Factory.CreateKnownObjectFromComProxy(this, returnItem,NetOffice.OWC10Api.GroupingDef.LateBindingApiWrapperType) as NetOffice.OWC10Api.GroupingDef;
			return newObject;
		}

		/// <summary>
		/// SupportByVersion OWC10 1
		/// 
		/// </summary>
		/// <param name="groupingDefName">string GroupingDefName</param>
		/// <param name="groupingFieldName">string GroupingFieldName</param>
		/// <param name="pageFieldName">string PageFieldName</param>
		[CustomMethodAttribute]
		[SupportByVersionAttribute("OWC10", 1)]
		public NetOffice.OWC10Api.GroupingDef Add(string groupingDefName, string groupingFieldName, string pageFieldName)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(groupingDefName, groupingFieldName, pageFieldName);
			object returnItem = Invoker.MethodReturn(this, "Add", paramsArray);
			NetOffice.OWC10Api.GroupingDef newObject = Factory.CreateKnownObjectFromComProxy(this, returnItem,NetOffice.OWC10Api.GroupingDef.LateBindingApiWrapperType) as NetOffice.OWC10Api.GroupingDef;
			return newObject;
		}

		/// <summary>
		/// SupportByVersion OWC10 1
		/// 
		/// </summary>
		/// <param name="groupingDefName">string GroupingDefName</param>
		/// <param name="groupingFieldName">string GroupingFieldName</param>
		/// <param name="pageFieldName">string PageFieldName</param>
		/// <param name="totalType">NetOffice.OWC10Api.Enums.DscTotalTypeEnum TotalType</param>
		/// <param name="index">optional object Index</param>
		[SupportByVersionAttribute("OWC10", 1)]
		public NetOffice.OWC10Api.GroupingDef AddTotal(string groupingDefName, string groupingFieldName, string pageFieldName, NetOffice.OWC10Api.Enums.DscTotalTypeEnum totalType, object index)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(groupingDefName, groupingFieldName, pageFieldName, totalType, index);
			object returnItem = Invoker.MethodReturn(this, "AddTotal", paramsArray);
			NetOffice.OWC10Api.GroupingDef newObject = Factory.CreateKnownObjectFromComProxy(this, returnItem,NetOffice.OWC10Api.GroupingDef.LateBindingApiWrapperType) as NetOffice.OWC10Api.GroupingDef;
			return newObject;
		}

		/// <summary>
		/// SupportByVersion OWC10 1
		/// 
		/// </summary>
		/// <param name="groupingDefName">string GroupingDefName</param>
		/// <param name="groupingFieldName">string GroupingFieldName</param>
		/// <param name="pageFieldName">string PageFieldName</param>
		/// <param name="totalType">NetOffice.OWC10Api.Enums.DscTotalTypeEnum TotalType</param>
		[CustomMethodAttribute]
		[SupportByVersionAttribute("OWC10", 1)]
		public NetOffice.OWC10Api.GroupingDef AddTotal(string groupingDefName, string groupingFieldName, string pageFieldName, NetOffice.OWC10Api.Enums.DscTotalTypeEnum totalType)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(groupingDefName, groupingFieldName, pageFieldName, totalType);
			object returnItem = Invoker.MethodReturn(this, "AddTotal", paramsArray);
			NetOffice.OWC10Api.GroupingDef newObject = Factory.CreateKnownObjectFromComProxy(this, returnItem,NetOffice.OWC10Api.GroupingDef.LateBindingApiWrapperType) as NetOffice.OWC10Api.GroupingDef;
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

       #region IEnumerable<NetOffice.OWC10Api.GroupingDef> Member
        
        /// <summary>
		/// SupportByVersionAttribute OWC10, 1
		/// </summary>
		[SupportByVersionAttribute("OWC10", 1)]
       public IEnumerator<NetOffice.OWC10Api.GroupingDef> GetEnumerator()  
       {
           NetRuntimeSystem.Collections.IEnumerable innerEnumerator = (this as NetRuntimeSystem.Collections.IEnumerable);
           foreach (NetOffice.OWC10Api.GroupingDef item in innerEnumerator)
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