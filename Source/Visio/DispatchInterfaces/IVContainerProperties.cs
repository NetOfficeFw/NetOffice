using System;
using NetRuntimeSystem = System;
using System.ComponentModel;
using NetOffice.Attributes;

namespace NetOffice.VisioApi
{
	/// <summary>
	/// DispatchInterface IVContainerProperties 
	/// SupportByVersion Visio, 14,15,16
	/// </summary>
	[SupportByVersion("Visio", 14,15,16)]
	[EntityType(EntityType.IsDispatchInterface), BaseType]
 	public class IVContainerProperties : COMObject
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
                    _type = typeof(IVContainerProperties);
                return _type;
            }
        }
        
        #endregion
        
		#region Ctor

		/// <param name="factory">current used factory core</param>
		/// <param name="parentObject">object there has created the proxy</param>
		/// <param name="proxyShare">proxy share instead if com proxy</param>
		public IVContainerProperties(Core factory, ICOMObject parentObject, COMProxyShare proxyShare) : base(factory, parentObject, proxyShare)
		{
		}

		///<param name="factory">current used factory core</param>
		///<param name="parentObject">object there has created the proxy</param>
        ///<param name="comProxy">inner wrapped COM proxy</param>
		public IVContainerProperties(Core factory, ICOMObject parentObject, object comProxy) : base(factory, parentObject, comProxy)
		{
			
		}

        ///<param name="parentObject">object there has created the proxy</param>
        ///<param name="comProxy">inner wrapped COM proxy</param>
        [EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public IVContainerProperties(ICOMObject parentObject, object comProxy) : base(parentObject, comProxy)
		{
		}
		
		///<param name="factory">current used factory core</param>
		///<param name="parentObject">object there has created the proxy</param>
        ///<param name="comProxy">inner wrapped COM proxy</param>
        ///<param name="comProxyType">Type of inner wrapped COM proxy"</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public IVContainerProperties(Core factory, ICOMObject parentObject, object comProxy, NetRuntimeSystem.Type comProxyType) : base(factory, parentObject, comProxy, comProxyType)
		{

		}

		///<param name="parentObject">object there has created the proxy</param>
        ///<param name="comProxy">inner wrapped COM proxy</param>
        ///<param name="comProxyType">Type of inner wrapped COM proxy"</param>
        [EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public IVContainerProperties(ICOMObject parentObject, object comProxy, NetRuntimeSystem.Type comProxyType) : base(parentObject, comProxy, comProxyType)
		{
		}
		
		///<param name="replacedObject">object to replaced. replacedObject are not usable after this action</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public IVContainerProperties(ICOMObject replacedObject) : base(replacedObject)
		{
		}
		
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public IVContainerProperties() : base()
		{
		}
		
		/// <param name="progId">registered progID</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public IVContainerProperties(string progId) : base(progId)
		{
		}
		
		#endregion
		
		#region Properties

		/// <summary>
		/// SupportByVersion Visio 14, 15, 16
		/// Get
		/// </summary>
		[SupportByVersion("Visio", 14,15,16)]
		[BaseResult]
		public NetOffice.VisioApi.IVApplication Application
		{
			get
			{
				return Factory.ExecuteBaseReferencePropertyGet<NetOffice.VisioApi.IVApplication>(this, "Application");
			}
		}

		/// <summary>
		/// SupportByVersion Visio 14, 15, 16
		/// Get
		/// </summary>
		[SupportByVersion("Visio", 14,15,16)]
		public Int16 Stat
		{
			get
			{
				return Factory.ExecuteInt16PropertyGet(this, "Stat");
			}
		}

		/// <summary>
		/// SupportByVersion Visio 14, 15, 16
		/// Get
		/// </summary>
		[SupportByVersion("Visio", 14,15,16)]
		public Int16 ObjectType
		{
			get
			{
				return Factory.ExecuteInt16PropertyGet(this, "ObjectType");
			}
		}

		/// <summary>
		/// SupportByVersion Visio 14, 15, 16
		/// Get
		/// </summary>
		[SupportByVersion("Visio", 14,15,16)]
		[BaseResult]
		public NetOffice.VisioApi.IVDocument Document
		{
			get
			{
				return Factory.ExecuteBaseReferencePropertyGet<NetOffice.VisioApi.IVDocument>(this, "Document");
			}
		}

		/// <summary>
		/// SupportByVersion Visio 14, 15, 16
		/// Get
		/// </summary>
		[SupportByVersion("Visio", 14,15,16)]
		[BaseResult]
		public NetOffice.VisioApi.IVShape Shape
		{
			get
			{
				return Factory.ExecuteBaseReferencePropertyGet<NetOffice.VisioApi.IVShape>(this, "Shape");
			}
		}

		/// <summary>
		/// SupportByVersion Visio 14, 15, 16
		/// Get
		/// </summary>
		[SupportByVersion("Visio", 14,15,16)]
		public NetOffice.VisioApi.Enums.VisContainerTypes ContainerType
		{
			get
			{
				return Factory.ExecuteEnumPropertyGet<NetOffice.VisioApi.Enums.VisContainerTypes>(this, "ContainerType");
			}
		}

		/// <summary>
		/// SupportByVersion Visio 14, 15, 16
		/// Get/Set
		/// </summary>
		[SupportByVersion("Visio", 14,15,16)]
		public NetOffice.VisioApi.Enums.VisListAlignment ListAlignment
		{
			get
			{
				return Factory.ExecuteEnumPropertyGet<NetOffice.VisioApi.Enums.VisListAlignment>(this, "ListAlignment");
			}
			set
			{
				Factory.ExecuteEnumPropertySet(this, "ListAlignment", value);
			}
		}

		/// <summary>
		/// SupportByVersion Visio 14, 15, 16
		/// Get/Set
		/// </summary>
		[SupportByVersion("Visio", 14,15,16)]
		public NetOffice.VisioApi.Enums.VisListDirection ListDirection
		{
			get
			{
				return Factory.ExecuteEnumPropertyGet<NetOffice.VisioApi.Enums.VisListDirection>(this, "ListDirection");
			}
			set
			{
				Factory.ExecuteEnumPropertySet(this, "ListDirection", value);
			}
		}

		/// <summary>
		/// SupportByVersion Visio 14, 15, 16
		/// Get/Set
		/// </summary>
		[SupportByVersion("Visio", 14,15,16)]
		public bool LockMembership
		{
			get
			{
				return Factory.ExecuteBoolPropertyGet(this, "LockMembership");
			}
			set
			{
				Factory.ExecuteValuePropertySet(this, "LockMembership", value);
			}
		}

		/// <summary>
		/// SupportByVersion Visio 14, 15, 16
		/// Get/Set
		/// </summary>
		[SupportByVersion("Visio", 14,15,16)]
		public NetOffice.VisioApi.Enums.VisContainerAutoResize ResizeAsNeeded
		{
			get
			{
				return Factory.ExecuteEnumPropertyGet<NetOffice.VisioApi.Enums.VisContainerAutoResize>(this, "ResizeAsNeeded");
			}
			set
			{
				Factory.ExecuteEnumPropertySet(this, "ResizeAsNeeded", value);
			}
		}

		/// <summary>
		/// SupportByVersion Visio 14, 15, 16
		/// Get/Set
		/// </summary>
		[SupportByVersion("Visio", 14,15,16)]
		[BaseResult]
		public NetOffice.VisioApi.IVShape OverlappedList
		{
			get
			{
				return Factory.ExecuteBaseReferencePropertyGet<NetOffice.VisioApi.IVShape>(this, "OverlappedList");
			}
			set
			{
				Factory.ExecuteReferencePropertySet(this, "OverlappedList", value);
			}
		}

		/// <summary>
		/// SupportByVersion Visio 14, 15, 16
		/// Get/Set
		/// </summary>
		[SupportByVersion("Visio", 14,15,16)]
		public Int32 ContainerStyle
		{
			get
			{
				return Factory.ExecuteInt32PropertyGet(this, "ContainerStyle");
			}
			set
			{
				Factory.ExecuteValuePropertySet(this, "ContainerStyle", value);
			}
		}

		/// <summary>
		/// SupportByVersion Visio 14, 15, 16
		/// Get/Set
		/// </summary>
		[SupportByVersion("Visio", 14,15,16)]
		public Int32 HeadingStyle
		{
			get
			{
				return Factory.ExecuteInt32PropertyGet(this, "HeadingStyle");
			}
			set
			{
				Factory.ExecuteValuePropertySet(this, "HeadingStyle", value);
			}
		}

		#endregion

		#region Methods

		/// <summary>
		/// SupportByVersion Visio 14, 15, 16
		/// </summary>
		[SupportByVersion("Visio", 14,15,16)]
		public void Disband()
		{
			 Factory.ExecuteMethod(this, "Disband");
		}

		/// <summary>
		/// SupportByVersion Visio 14, 15, 16
		/// </summary>
		[SupportByVersion("Visio", 14,15,16)]
		public void FitToContents()
		{
			 Factory.ExecuteMethod(this, "FitToContents");
		}

		/// <summary>
		/// SupportByVersion Visio 14, 15, 16
		/// </summary>
		/// <param name="marginUnits">NetOffice.VisioApi.Enums.VisUnitCodes marginUnits</param>
		[SupportByVersion("Visio", 14,15,16)]
		public Double GetMargin(NetOffice.VisioApi.Enums.VisUnitCodes marginUnits)
		{
			return Factory.ExecuteDoubleMethodGet(this, "GetMargin", marginUnits);
		}

		/// <summary>
		/// SupportByVersion Visio 14, 15, 16
		/// </summary>
		/// <param name="marginUnits">NetOffice.VisioApi.Enums.VisUnitCodes marginUnits</param>
		/// <param name="marginSize">Double marginSize</param>
		[SupportByVersion("Visio", 14,15,16)]
		public void SetMargin(NetOffice.VisioApi.Enums.VisUnitCodes marginUnits, Double marginSize)
		{
			 Factory.ExecuteMethod(this, "SetMargin", marginUnits, marginSize);
		}

		/// <summary>
		/// SupportByVersion Visio 14, 15, 16
		/// </summary>
		/// <param name="spacingUnits">NetOffice.VisioApi.Enums.VisUnitCodes spacingUnits</param>
		[SupportByVersion("Visio", 14,15,16)]
		public Double GetListSpacing(NetOffice.VisioApi.Enums.VisUnitCodes spacingUnits)
		{
			return Factory.ExecuteDoubleMethodGet(this, "GetListSpacing", spacingUnits);
		}

		/// <summary>
		/// SupportByVersion Visio 14, 15, 16
		/// </summary>
		/// <param name="spacingUnits">NetOffice.VisioApi.Enums.VisUnitCodes spacingUnits</param>
		/// <param name="spacingSize">Double spacingSize</param>
		[SupportByVersion("Visio", 14,15,16)]
		public void SetListSpacing(NetOffice.VisioApi.Enums.VisUnitCodes spacingUnits, Double spacingSize)
		{
			 Factory.ExecuteMethod(this, "SetListSpacing", spacingUnits, spacingSize);
		}

		/// <summary>
		/// SupportByVersion Visio 14, 15, 16
		/// </summary>
		/// <param name="objectToInsert">object objectToInsert</param>
		/// <param name="position">Int32 position</param>
		[SupportByVersion("Visio", 14,15,16)]
		public void InsertListMember(object objectToInsert, Int32 position)
		{
			 Factory.ExecuteMethod(this, "InsertListMember", objectToInsert, position);
		}

		/// <summary>
		/// SupportByVersion Visio 14, 15, 16
		/// </summary>
		/// <param name="shapeMember">NetOffice.VisioApi.IVShape shapeMember</param>
		[SupportByVersion("Visio", 14,15,16)]
		public Int32 GetListMemberPosition(NetOffice.VisioApi.IVShape shapeMember)
		{
			return Factory.ExecuteInt32MethodGet(this, "GetListMemberPosition", shapeMember);
		}

		/// <summary>
		/// SupportByVersion Visio 14, 15, 16
		/// </summary>
		/// <param name="shape">NetOffice.VisioApi.IVShape shape</param>
		[SupportByVersion("Visio", 14,15,16)]
		public NetOffice.VisioApi.Enums.VisContainerMemberState GetMemberState(NetOffice.VisioApi.IVShape shape)
		{
			return Factory.ExecuteEnumMethodGet<NetOffice.VisioApi.Enums.VisContainerMemberState>(this, "GetMemberState", shape);
		}

		/// <summary>
		/// SupportByVersion Visio 14, 15, 16
		/// </summary>
		/// <param name="objectToRemove">object objectToRemove</param>
		[SupportByVersion("Visio", 14,15,16)]
		public void RemoveMember(object objectToRemove)
		{
			 Factory.ExecuteMethod(this, "RemoveMember", objectToRemove);
		}

		/// <summary>
		/// SupportByVersion Visio 14, 15, 16
		/// </summary>
		/// <param name="objectToReorder">object objectToReorder</param>
		/// <param name="position">Int32 position</param>
		[SupportByVersion("Visio", 14,15,16)]
		public void ReorderListMember(object objectToReorder, Int32 position)
		{
			 Factory.ExecuteMethod(this, "ReorderListMember", objectToReorder, position);
		}

		/// <summary>
		/// SupportByVersion Visio 14, 15, 16
		/// </summary>
		[SupportByVersion("Visio", 14,15,16)]
		public Int32[] GetListMembers()
		{
			object[] paramsArray = null;
			object returnItem = (object)Invoker.MethodReturn(this, "GetListMembers", paramsArray);
			return (Int32[])returnItem;
		}

		/// <summary>
		/// SupportByVersion Visio 14, 15, 16
		/// </summary>
		/// <param name="containerFlags">Int32 containerFlags</param>
		[SupportByVersion("Visio", 14,15,16)]
		public Int32[] GetMemberShapes(Int32 containerFlags)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(containerFlags);
			object returnItem = (object)Invoker.MethodReturn(this, "GetMemberShapes", paramsArray);
			return (Int32[])returnItem;
		}

		/// <summary>
		/// SupportByVersion Visio 14, 15, 16
		/// </summary>
		/// <param name="pObjectToAdd">object pObjectToAdd</param>
		/// <param name="addOptions">NetOffice.VisioApi.Enums.VisMemberAddOptions addOptions</param>
		[SupportByVersion("Visio", 14,15,16)]
		public void AddMember(object pObjectToAdd, NetOffice.VisioApi.Enums.VisMemberAddOptions addOptions)
		{
			 Factory.ExecuteMethod(this, "AddMember", pObjectToAdd, addOptions);
		}

		/// <summary>
		/// SupportByVersion Visio 14, 15, 16
		/// </summary>
		/// <param name="direction">NetOffice.VisioApi.Enums.VisLayoutDirection direction</param>
		[SupportByVersion("Visio", 14,15,16)]
		public void RotateFlipList(NetOffice.VisioApi.Enums.VisLayoutDirection direction)
		{
			 Factory.ExecuteMethod(this, "RotateFlipList", direction);
		}

		#endregion

		#pragma warning restore
	}
}
