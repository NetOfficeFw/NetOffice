using System;
using NetRuntimeSystem = System;
using System.ComponentModel;
using NetOffice.Attributes;
using NetOffice.VisioApi;

namespace NetOffice.VisioApi.Behind
{
	/// <summary>
	/// DispatchInterface IVContainerProperties 
	/// SupportByVersion Visio, 14,15,16
	/// </summary>
	[SupportByVersion("Visio", 14,15,16)]
	[EntityType(EntityType.IsDispatchInterface), BaseType]
 	public class IVContainerProperties : COMObject, NetOffice.VisioApi.IVContainerProperties
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
                    _contractType = typeof(NetOffice.VisioApi.IVContainerProperties);
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
                    _type = typeof(IVContainerProperties);
                return _type;
            }
        }
        
        #endregion
        
		#region Ctor

		/// <summary>
		/// Stub Ctor, not indented to use
		/// </summary>
		public IVContainerProperties() : base()
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
		public virtual NetOffice.VisioApi.IVApplication Application
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteBaseReferencePropertyGet<NetOffice.VisioApi.IVApplication>(this, "Application");
			}
		}

		/// <summary>
		/// SupportByVersion Visio 14, 15, 16
		/// Get
		/// </summary>
		[SupportByVersion("Visio", 14,15,16)]
		public virtual Int16 Stat
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteInt16PropertyGet(this, "Stat");
			}
		}

		/// <summary>
		/// SupportByVersion Visio 14, 15, 16
		/// Get
		/// </summary>
		[SupportByVersion("Visio", 14,15,16)]
		public virtual Int16 ObjectType
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteInt16PropertyGet(this, "ObjectType");
			}
		}

		/// <summary>
		/// SupportByVersion Visio 14, 15, 16
		/// Get
		/// </summary>
		[SupportByVersion("Visio", 14,15,16)]
		[BaseResult]
		public virtual NetOffice.VisioApi.IVDocument Document
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteBaseReferencePropertyGet<NetOffice.VisioApi.IVDocument>(this, "Document");
			}
		}

		/// <summary>
		/// SupportByVersion Visio 14, 15, 16
		/// Get
		/// </summary>
		[SupportByVersion("Visio", 14,15,16)]
		[BaseResult]
		public virtual NetOffice.VisioApi.IVShape Shape
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteBaseReferencePropertyGet<NetOffice.VisioApi.IVShape>(this, "Shape");
			}
		}

		/// <summary>
		/// SupportByVersion Visio 14, 15, 16
		/// Get
		/// </summary>
		[SupportByVersion("Visio", 14,15,16)]
		public virtual NetOffice.VisioApi.Enums.VisContainerTypes ContainerType
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteEnumPropertyGet<NetOffice.VisioApi.Enums.VisContainerTypes>(this, "ContainerType");
			}
		}

		/// <summary>
		/// SupportByVersion Visio 14, 15, 16
		/// Get/Set
		/// </summary>
		[SupportByVersion("Visio", 14,15,16)]
		public virtual NetOffice.VisioApi.Enums.VisListAlignment ListAlignment
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteEnumPropertyGet<NetOffice.VisioApi.Enums.VisListAlignment>(this, "ListAlignment");
			}
			set
			{
				InvokerService.InvokeInternal.ExecuteEnumPropertySet(this, "ListAlignment", value);
			}
		}

		/// <summary>
		/// SupportByVersion Visio 14, 15, 16
		/// Get/Set
		/// </summary>
		[SupportByVersion("Visio", 14,15,16)]
		public virtual NetOffice.VisioApi.Enums.VisListDirection ListDirection
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteEnumPropertyGet<NetOffice.VisioApi.Enums.VisListDirection>(this, "ListDirection");
			}
			set
			{
				InvokerService.InvokeInternal.ExecuteEnumPropertySet(this, "ListDirection", value);
			}
		}

		/// <summary>
		/// SupportByVersion Visio 14, 15, 16
		/// Get/Set
		/// </summary>
		[SupportByVersion("Visio", 14,15,16)]
		public virtual bool LockMembership
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteBoolPropertyGet(this, "LockMembership");
			}
			set
			{
				InvokerService.InvokeInternal.ExecuteValuePropertySet(this, "LockMembership", value);
			}
		}

		/// <summary>
		/// SupportByVersion Visio 14, 15, 16
		/// Get/Set
		/// </summary>
		[SupportByVersion("Visio", 14,15,16)]
		public virtual NetOffice.VisioApi.Enums.VisContainerAutoResize ResizeAsNeeded
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteEnumPropertyGet<NetOffice.VisioApi.Enums.VisContainerAutoResize>(this, "ResizeAsNeeded");
			}
			set
			{
				InvokerService.InvokeInternal.ExecuteEnumPropertySet(this, "ResizeAsNeeded", value);
			}
		}

		/// <summary>
		/// SupportByVersion Visio 14, 15, 16
		/// Get/Set
		/// </summary>
		[SupportByVersion("Visio", 14,15,16)]
		[BaseResult]
		public virtual NetOffice.VisioApi.IVShape OverlappedList
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteBaseReferencePropertyGet<NetOffice.VisioApi.IVShape>(this, "OverlappedList");
			}
			set
			{
				InvokerService.InvokeInternal.ExecuteReferencePropertySet(this, "OverlappedList", value);
			}
		}

		/// <summary>
		/// SupportByVersion Visio 14, 15, 16
		/// Get/Set
		/// </summary>
		[SupportByVersion("Visio", 14,15,16)]
		public virtual Int32 ContainerStyle
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteInt32PropertyGet(this, "ContainerStyle");
			}
			set
			{
				InvokerService.InvokeInternal.ExecuteValuePropertySet(this, "ContainerStyle", value);
			}
		}

		/// <summary>
		/// SupportByVersion Visio 14, 15, 16
		/// Get/Set
		/// </summary>
		[SupportByVersion("Visio", 14,15,16)]
		public virtual Int32 HeadingStyle
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteInt32PropertyGet(this, "HeadingStyle");
			}
			set
			{
				InvokerService.InvokeInternal.ExecuteValuePropertySet(this, "HeadingStyle", value);
			}
		}

		#endregion

		#region Methods

		/// <summary>
		/// SupportByVersion Visio 14, 15, 16
		/// </summary>
		[SupportByVersion("Visio", 14,15,16)]
		public virtual void Disband()
		{
			 InvokerService.InvokeInternal.ExecuteMethod(this, "Disband");
		}

		/// <summary>
		/// SupportByVersion Visio 14, 15, 16
		/// </summary>
		[SupportByVersion("Visio", 14,15,16)]
		public virtual void FitToContents()
		{
			 InvokerService.InvokeInternal.ExecuteMethod(this, "FitToContents");
		}

		/// <summary>
		/// SupportByVersion Visio 14, 15, 16
		/// </summary>
		/// <param name="marginUnits">NetOffice.VisioApi.Enums.VisUnitCodes marginUnits</param>
		[SupportByVersion("Visio", 14,15,16)]
		public virtual Double GetMargin(NetOffice.VisioApi.Enums.VisUnitCodes marginUnits)
		{
			return InvokerService.InvokeInternal.ExecuteDoubleMethodGet(this, "GetMargin", marginUnits);
		}

		/// <summary>
		/// SupportByVersion Visio 14, 15, 16
		/// </summary>
		/// <param name="marginUnits">NetOffice.VisioApi.Enums.VisUnitCodes marginUnits</param>
		/// <param name="marginSize">Double marginSize</param>
		[SupportByVersion("Visio", 14,15,16)]
		public virtual void SetMargin(NetOffice.VisioApi.Enums.VisUnitCodes marginUnits, Double marginSize)
		{
			 InvokerService.InvokeInternal.ExecuteMethod(this, "SetMargin", marginUnits, marginSize);
		}

		/// <summary>
		/// SupportByVersion Visio 14, 15, 16
		/// </summary>
		/// <param name="spacingUnits">NetOffice.VisioApi.Enums.VisUnitCodes spacingUnits</param>
		[SupportByVersion("Visio", 14,15,16)]
		public virtual Double GetListSpacing(NetOffice.VisioApi.Enums.VisUnitCodes spacingUnits)
		{
			return InvokerService.InvokeInternal.ExecuteDoubleMethodGet(this, "GetListSpacing", spacingUnits);
		}

		/// <summary>
		/// SupportByVersion Visio 14, 15, 16
		/// </summary>
		/// <param name="spacingUnits">NetOffice.VisioApi.Enums.VisUnitCodes spacingUnits</param>
		/// <param name="spacingSize">Double spacingSize</param>
		[SupportByVersion("Visio", 14,15,16)]
		public virtual void SetListSpacing(NetOffice.VisioApi.Enums.VisUnitCodes spacingUnits, Double spacingSize)
		{
			 InvokerService.InvokeInternal.ExecuteMethod(this, "SetListSpacing", spacingUnits, spacingSize);
		}

		/// <summary>
		/// SupportByVersion Visio 14, 15, 16
		/// </summary>
		/// <param name="objectToInsert">object objectToInsert</param>
		/// <param name="position">Int32 position</param>
		[SupportByVersion("Visio", 14,15,16)]
		public virtual void InsertListMember(object objectToInsert, Int32 position)
		{
			 InvokerService.InvokeInternal.ExecuteMethod(this, "InsertListMember", objectToInsert, position);
		}

		/// <summary>
		/// SupportByVersion Visio 14, 15, 16
		/// </summary>
		/// <param name="shapeMember">NetOffice.VisioApi.IVShape shapeMember</param>
		[SupportByVersion("Visio", 14,15,16)]
		public virtual Int32 GetListMemberPosition(NetOffice.VisioApi.IVShape shapeMember)
		{
			return InvokerService.InvokeInternal.ExecuteInt32MethodGet(this, "GetListMemberPosition", shapeMember);
		}

		/// <summary>
		/// SupportByVersion Visio 14, 15, 16
		/// </summary>
		/// <param name="shape">NetOffice.VisioApi.IVShape shape</param>
		[SupportByVersion("Visio", 14,15,16)]
		public virtual NetOffice.VisioApi.Enums.VisContainerMemberState GetMemberState(NetOffice.VisioApi.IVShape shape)
		{
			return InvokerService.InvokeInternal.ExecuteEnumMethodGet<NetOffice.VisioApi.Enums.VisContainerMemberState>(this, "GetMemberState", shape);
		}

		/// <summary>
		/// SupportByVersion Visio 14, 15, 16
		/// </summary>
		/// <param name="objectToRemove">object objectToRemove</param>
		[SupportByVersion("Visio", 14,15,16)]
		public virtual void RemoveMember(object objectToRemove)
		{
			 InvokerService.InvokeInternal.ExecuteMethod(this, "RemoveMember", objectToRemove);
		}

		/// <summary>
		/// SupportByVersion Visio 14, 15, 16
		/// </summary>
		/// <param name="objectToReorder">object objectToReorder</param>
		/// <param name="position">Int32 position</param>
		[SupportByVersion("Visio", 14,15,16)]
		public virtual void ReorderListMember(object objectToReorder, Int32 position)
		{
			 InvokerService.InvokeInternal.ExecuteMethod(this, "ReorderListMember", objectToReorder, position);
		}

		/// <summary>
		/// SupportByVersion Visio 14, 15, 16
		/// </summary>
		[SupportByVersion("Visio", 14,15,16)]
		public virtual Int32[] GetListMembers()
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
		public virtual Int32[] GetMemberShapes(Int32 containerFlags)
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
		public virtual void AddMember(object pObjectToAdd, NetOffice.VisioApi.Enums.VisMemberAddOptions addOptions)
		{
			 InvokerService.InvokeInternal.ExecuteMethod(this, "AddMember", pObjectToAdd, addOptions);
		}

		/// <summary>
		/// SupportByVersion Visio 14, 15, 16
		/// </summary>
		/// <param name="direction">NetOffice.VisioApi.Enums.VisLayoutDirection direction</param>
		[SupportByVersion("Visio", 14,15,16)]
		public virtual void RotateFlipList(NetOffice.VisioApi.Enums.VisLayoutDirection direction)
		{
			 InvokerService.InvokeInternal.ExecuteMethod(this, "RotateFlipList", direction);
		}

		#endregion

		#pragma warning restore
	}
}

