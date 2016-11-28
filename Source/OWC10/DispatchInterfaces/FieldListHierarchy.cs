using System;
using NetRuntimeSystem = System;
using System.Runtime.InteropServices;
using System.Runtime.CompilerServices;
using System.ComponentModel;
using System.Reflection;
using System.Collections.Generic;
using NetOffice;
namespace NetOffice.OWC10Api
{
	///<summary>
	/// DispatchInterface FieldListHierarchy 
	/// SupportByVersion OWC10, 1
	///</summary>
	[SupportByVersionAttribute("OWC10", 1)]
	[EntityTypeAttribute(EntityType.IsDispatchInterface)]
	public class FieldListHierarchy : COMObject
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
                    _type = typeof(FieldListHierarchy);
                    
                return _type;
            }
        }
        
        #endregion
        
		#region Construction

		///<param name="factory">current used factory core</param>
		///<param name="parentObject">object there has created the proxy</param>
        ///<param name="comProxy">inner wrapped COM proxy</param>
		public FieldListHierarchy(Core factory, ICOMObject parentObject, object comProxy) : base(factory, parentObject, comProxy)
		{
			
		}

        ///<param name="parentObject">object there has created the proxy</param>
        ///<param name="comProxy">inner wrapped COM proxy</param>
        [EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public FieldListHierarchy(ICOMObject parentObject, object comProxy) : base(parentObject, comProxy)
		{
		}
		
		///<param name="factory">current used factory core</param>
		///<param name="parentObject">object there has created the proxy</param>
        ///<param name="comProxy">inner wrapped COM proxy</param>
        ///<param name="comProxyType">Type of inner wrapped COM proxy"</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public FieldListHierarchy(Core factory, ICOMObject parentObject, object comProxy, NetRuntimeSystem.Type comProxyType) : base(factory, parentObject, comProxy, comProxyType)
		{

		}

		///<param name="parentObject">object there has created the proxy</param>
        ///<param name="comProxy">inner wrapped COM proxy</param>
        ///<param name="comProxyType">Type of inner wrapped COM proxy"</param>
        [EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public FieldListHierarchy(ICOMObject parentObject, object comProxy, NetRuntimeSystem.Type comProxyType) : base(parentObject, comProxy, comProxyType)
		{
		}
		
		///<param name="replacedObject">object to replaced. replacedObject are not usable after this action</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public FieldListHierarchy(ICOMObject replacedObject) : base(replacedObject)
		{
		}
		
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public FieldListHierarchy() : base()
		{
		}
		
		/// <param name="progId">registered ProgID</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public FieldListHierarchy(string progId) : base(progId)
		{
		}
		
		#endregion
		
		#region Properties

		/// <summary>
		/// SupportByVersion OWC10 1
		/// Get
		/// </summary>
		[SupportByVersionAttribute("OWC10", 1)]
		public NetOffice.OWC10Api.FieldListNode Root
		{
			get
			{
				object[] paramsArray = null;
				object returnItem = Invoker.PropertyGet(this, "Root", paramsArray);
				NetOffice.OWC10Api.FieldListNode newObject = Factory.CreateKnownObjectFromComProxy(this,returnItem,NetOffice.OWC10Api.FieldListNode.LateBindingApiWrapperType) as NetOffice.OWC10Api.FieldListNode;
				return newObject;
			}
		}

		/// <summary>
		/// SupportByVersion OWC10 1
		/// Get/Set
		/// </summary>
		[SupportByVersionAttribute("OWC10", 1)]
		public bool Visible
		{
			get
			{
				object[] paramsArray = null;
				object returnItem = Invoker.PropertyGet(this, "Visible", paramsArray);
				return NetRuntimeSystem.Convert.ToBoolean(returnItem);
			}
			set
			{
				object[] paramsArray = Invoker.ValidateParamsArray(value);
				Invoker.PropertySet(this, "Visible", paramsArray);
			}
		}

		/// <summary>
		/// SupportByVersion OWC10 1
		/// Get
		/// </summary>
		[SupportByVersionAttribute("OWC10", 1)]
		public NetOffice.OWC10Api.FieldListNode Selection
		{
			get
			{
				object[] paramsArray = null;
				object returnItem = Invoker.PropertyGet(this, "Selection", paramsArray);
				NetOffice.OWC10Api.FieldListNode newObject = Factory.CreateKnownObjectFromComProxy(this,returnItem,NetOffice.OWC10Api.FieldListNode.LateBindingApiWrapperType) as NetOffice.OWC10Api.FieldListNode;
				return newObject;
			}
		}

		/// <summary>
		/// SupportByVersion OWC10 1
		/// Get/Set
		/// </summary>
		[SupportByVersionAttribute("OWC10", 1)]
		public bool ConcatenateData
		{
			get
			{
				object[] paramsArray = null;
				object returnItem = Invoker.PropertyGet(this, "ConcatenateData", paramsArray);
				return NetRuntimeSystem.Convert.ToBoolean(returnItem);
			}
			set
			{
				object[] paramsArray = Invoker.ValidateParamsArray(value);
				Invoker.PropertySet(this, "ConcatenateData", paramsArray);
			}
		}

		/// <summary>
		/// SupportByVersion OWC10 1
		/// Get/Set
		/// </summary>
		[SupportByVersionAttribute("OWC10", 1)]
		public string DataSeparator
		{
			get
			{
				object[] paramsArray = null;
				object returnItem = Invoker.PropertyGet(this, "DataSeparator", paramsArray);
				return NetRuntimeSystem.Convert.ToString(returnItem);
			}
			set
			{
				object[] paramsArray = Invoker.ValidateParamsArray(value);
				Invoker.PropertySet(this, "DataSeparator", paramsArray);
			}
		}

		#endregion

		#region Methods

		/// <summary>
		/// SupportByVersion OWC10 1
		/// 
		/// </summary>
		/// <param name="pflhs">NetOffice.OWC10Api.FieldListHierarchySite pflhs</param>
		[SupportByVersionAttribute("OWC10", 1)]
		public void SetHierarchySite(NetOffice.OWC10Api.FieldListHierarchySite pflhs)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(pflhs);
			Invoker.Method(this, "SetHierarchySite", paramsArray);
		}

		/// <summary>
		/// SupportByVersion OWC10 1
		/// 
		/// </summary>
		/// <param name="pflnParent">NetOffice.OWC10Api.FieldListNode pflnParent</param>
		/// <param name="fInsertFirst">bool fInsertFirst</param>
		/// <param name="nID">Int32 nID</param>
		/// <param name="bstrName">string bstrName</param>
		/// <param name="bstrData">string bstrData</param>
		/// <param name="nType">Int32 nType</param>
		[SupportByVersionAttribute("OWC10", 1)]
		public NetOffice.OWC10Api.FieldListNode AddNode(NetOffice.OWC10Api.FieldListNode pflnParent, bool fInsertFirst, Int32 nID, string bstrName, string bstrData, Int32 nType)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(pflnParent, fInsertFirst, nID, bstrName, bstrData, nType);
			object returnItem = Invoker.MethodReturn(this, "AddNode", paramsArray);
			NetOffice.OWC10Api.FieldListNode newObject = Factory.CreateKnownObjectFromComProxy(this, returnItem,NetOffice.OWC10Api.FieldListNode.LateBindingApiWrapperType) as NetOffice.OWC10Api.FieldListNode;
			return newObject;
		}

		/// <summary>
		/// SupportByVersion OWC10 1
		/// 
		/// </summary>
		/// <param name="nID">Int32 nID</param>
		[SupportByVersionAttribute("OWC10", 1)]
		public NetOffice.OWC10Api.FieldListNode GetNode(Int32 nID)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(nID);
			object returnItem = Invoker.MethodReturn(this, "GetNode", paramsArray);
			NetOffice.OWC10Api.FieldListNode newObject = Factory.CreateKnownObjectFromComProxy(this, returnItem,NetOffice.OWC10Api.FieldListNode.LateBindingApiWrapperType) as NetOffice.OWC10Api.FieldListNode;
			return newObject;
		}

		/// <summary>
		/// SupportByVersion OWC10 1
		/// 
		/// </summary>
		/// <param name="pfln">NetOffice.OWC10Api.FieldListNode pfln</param>
		[SupportByVersionAttribute("OWC10", 1)]
		public void RemoveNode(NetOffice.OWC10Api.FieldListNode pfln)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(pfln);
			Invoker.Method(this, "RemoveNode", paramsArray);
		}

		/// <summary>
		/// SupportByVersion OWC10 1
		/// 
		/// </summary>
		/// <param name="nType">Int32 nType</param>
		[SupportByVersionAttribute("OWC10", 1)]
		public NetOffice.OWC10Api.FieldListType AddType(Int32 nType)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(nType);
			object returnItem = Invoker.MethodReturn(this, "AddType", paramsArray);
			NetOffice.OWC10Api.FieldListType newObject = Factory.CreateKnownObjectFromComProxy(this, returnItem,NetOffice.OWC10Api.FieldListType.LateBindingApiWrapperType) as NetOffice.OWC10Api.FieldListType;
			return newObject;
		}

		/// <summary>
		/// SupportByVersion OWC10 1
		/// 
		/// </summary>
		/// <param name="nTypeId">Int32 nTypeId</param>
		[SupportByVersionAttribute("OWC10", 1)]
		public NetOffice.OWC10Api.FieldListType GetType(Int32 nTypeId)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(nTypeId);
			object returnItem = Invoker.MethodReturn(this, "GetType", paramsArray);
			NetOffice.OWC10Api.FieldListType newObject = Factory.CreateKnownObjectFromComProxy(this, returnItem,NetOffice.OWC10Api.FieldListType.LateBindingApiWrapperType) as NetOffice.OWC10Api.FieldListType;
			return newObject;
		}

		/// <summary>
		/// SupportByVersion OWC10 1
		/// 
		/// </summary>
		/// <param name="pfln">NetOffice.OWC10Api.FieldListNode pfln</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		[SupportByVersionAttribute("OWC10", 1)]
		public NetOffice.OWC10Api.FieldListNode GetNextSelected(NetOffice.OWC10Api.FieldListNode pfln)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(pfln);
			object returnItem = Invoker.MethodReturn(this, "GetNextSelected", paramsArray);
			NetOffice.OWC10Api.FieldListNode newObject = Factory.CreateKnownObjectFromComProxy(this, returnItem,NetOffice.OWC10Api.FieldListNode.LateBindingApiWrapperType) as NetOffice.OWC10Api.FieldListNode;
			return newObject;
		}

		#endregion
		#pragma warning restore
	}
}