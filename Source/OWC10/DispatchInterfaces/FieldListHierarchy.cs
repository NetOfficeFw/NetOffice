using System;
using NetRuntimeSystem = System;
using System.ComponentModel;
using NetOffice.Attributes;

namespace NetOffice.OWC10Api
{
	/// <summary>
	/// DispatchInterface FieldListHierarchy 
	/// SupportByVersion OWC10, 1
	/// </summary>
	[SupportByVersion("OWC10", 1)]
	[EntityType(EntityType.IsDispatchInterface)]
 	public class FieldListHierarchy : COMObject
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
                    _type = typeof(FieldListHierarchy);
                return _type;
            }
        }
        
        #endregion
        
		#region Ctor

		/// <param name="factory">current used factory core</param>
		/// <param name="parentObject">object there has created the proxy</param>
		/// <param name="proxyShare">proxy share instead if com proxy</param>
		public FieldListHierarchy(Core factory, ICOMObject parentObject, COMProxyShare proxyShare) : base(factory, parentObject, proxyShare)
		{
		}

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
		
		/// <param name="progId">registered progID</param>
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
		[SupportByVersion("OWC10", 1)]
		public NetOffice.OWC10Api.FieldListNode Root
		{
			get
			{
				return Factory.ExecuteKnownReferencePropertyGet<NetOffice.OWC10Api.FieldListNode>(this, "Root", NetOffice.OWC10Api.FieldListNode.LateBindingApiWrapperType);
			}
		}

		/// <summary>
		/// SupportByVersion OWC10 1
		/// Get/Set
		/// </summary>
		[SupportByVersion("OWC10", 1)]
		public bool Visible
		{
			get
			{
				return Factory.ExecuteBoolPropertyGet(this, "Visible");
			}
			set
			{
				Factory.ExecuteValuePropertySet(this, "Visible", value);
			}
		}

		/// <summary>
		/// SupportByVersion OWC10 1
		/// Get
		/// </summary>
		[SupportByVersion("OWC10", 1)]
		public NetOffice.OWC10Api.FieldListNode Selection
		{
			get
			{
				return Factory.ExecuteKnownReferencePropertyGet<NetOffice.OWC10Api.FieldListNode>(this, "Selection", NetOffice.OWC10Api.FieldListNode.LateBindingApiWrapperType);
			}
		}

		/// <summary>
		/// SupportByVersion OWC10 1
		/// Get/Set
		/// </summary>
		[SupportByVersion("OWC10", 1)]
		public bool ConcatenateData
		{
			get
			{
				return Factory.ExecuteBoolPropertyGet(this, "ConcatenateData");
			}
			set
			{
				Factory.ExecuteValuePropertySet(this, "ConcatenateData", value);
			}
		}

		/// <summary>
		/// SupportByVersion OWC10 1
		/// Get/Set
		/// </summary>
		[SupportByVersion("OWC10", 1)]
		public string DataSeparator
		{
			get
			{
				return Factory.ExecuteStringPropertyGet(this, "DataSeparator");
			}
			set
			{
				Factory.ExecuteValuePropertySet(this, "DataSeparator", value);
			}
		}

		#endregion

		#region Methods

		/// <summary>
		/// SupportByVersion OWC10 1
		/// </summary>
		/// <param name="pflhs">NetOffice.OWC10Api.FieldListHierarchySite pflhs</param>
		[SupportByVersion("OWC10", 1)]
		public void SetHierarchySite(NetOffice.OWC10Api.FieldListHierarchySite pflhs)
		{
			 Factory.ExecuteMethod(this, "SetHierarchySite", pflhs);
		}

		/// <summary>
		/// SupportByVersion OWC10 1
		/// </summary>
		/// <param name="pflnParent">NetOffice.OWC10Api.FieldListNode pflnParent</param>
		/// <param name="fInsertFirst">bool fInsertFirst</param>
		/// <param name="nID">Int32 nID</param>
		/// <param name="bstrName">string bstrName</param>
		/// <param name="bstrData">string bstrData</param>
		/// <param name="nType">Int32 nType</param>
		[SupportByVersion("OWC10", 1)]
		public NetOffice.OWC10Api.FieldListNode AddNode(NetOffice.OWC10Api.FieldListNode pflnParent, bool fInsertFirst, Int32 nID, string bstrName, string bstrData, Int32 nType)
		{
			return Factory.ExecuteKnownReferenceMethodGet<NetOffice.OWC10Api.FieldListNode>(this, "AddNode", NetOffice.OWC10Api.FieldListNode.LateBindingApiWrapperType, new object[]{ pflnParent, fInsertFirst, nID, bstrName, bstrData, nType });
		}

		/// <summary>
		/// SupportByVersion OWC10 1
		/// </summary>
		/// <param name="nID">Int32 nID</param>
		[SupportByVersion("OWC10", 1)]
		public NetOffice.OWC10Api.FieldListNode GetNode(Int32 nID)
		{
			return Factory.ExecuteKnownReferenceMethodGet<NetOffice.OWC10Api.FieldListNode>(this, "GetNode", NetOffice.OWC10Api.FieldListNode.LateBindingApiWrapperType, nID);
		}

		/// <summary>
		/// SupportByVersion OWC10 1
		/// </summary>
		/// <param name="pfln">NetOffice.OWC10Api.FieldListNode pfln</param>
		[SupportByVersion("OWC10", 1)]
		public void RemoveNode(NetOffice.OWC10Api.FieldListNode pfln)
		{
			 Factory.ExecuteMethod(this, "RemoveNode", pfln);
		}

		/// <summary>
		/// SupportByVersion OWC10 1
		/// </summary>
		/// <param name="nType">Int32 nType</param>
		[SupportByVersion("OWC10", 1)]
		public NetOffice.OWC10Api.FieldListType AddType(Int32 nType)
		{
			return Factory.ExecuteKnownReferenceMethodGet<NetOffice.OWC10Api.FieldListType>(this, "AddType", NetOffice.OWC10Api.FieldListType.LateBindingApiWrapperType, nType);
		}

		/// <summary>
		/// SupportByVersion OWC10 1
		/// </summary>
		/// <param name="nTypeId">Int32 nTypeId</param>
		[SupportByVersion("OWC10", 1)]
		public NetOffice.OWC10Api.FieldListType GetType(Int32 nTypeId)
		{
			return Factory.ExecuteKnownReferenceMethodGet<NetOffice.OWC10Api.FieldListType>(this, "GetType", NetOffice.OWC10Api.FieldListType.LateBindingApiWrapperType, nTypeId);
		}

		/// <summary>
		/// SupportByVersion OWC10 1
		/// </summary>
		/// <param name="pfln">NetOffice.OWC10Api.FieldListNode pfln</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		[SupportByVersion("OWC10", 1)]
		public NetOffice.OWC10Api.FieldListNode GetNextSelected(NetOffice.OWC10Api.FieldListNode pfln)
		{
			return Factory.ExecuteKnownReferenceMethodGet<NetOffice.OWC10Api.FieldListNode>(this, "GetNextSelected", NetOffice.OWC10Api.FieldListNode.LateBindingApiWrapperType, pfln);
		}

		#endregion

		#pragma warning restore
	}
}
