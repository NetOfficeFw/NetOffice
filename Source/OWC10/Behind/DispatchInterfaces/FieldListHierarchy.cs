using System;
using NetRuntimeSystem = System;
using System.ComponentModel;
using NetOffice.Attributes;
using NetOffice.OWC10Api;

namespace NetOffice.OWC10Api.Behind
{
	/// <summary>
	/// DispatchInterface FieldListHierarchy 
	/// SupportByVersion OWC10, 1
	/// </summary>
	[SupportByVersion("OWC10", 1)]
	[EntityType(EntityType.IsDispatchInterface)]
 	public class FieldListHierarchy : COMObject, NetOffice.OWC10Api.FieldListHierarchy
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
                    _contractType = typeof(NetOffice.OWC10Api.FieldListHierarchy);
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
                    _type = typeof(FieldListHierarchy);
                return _type;
            }
        }
        
        #endregion
        
		#region Ctor

		/// <summary>
		/// Stub Ctor, not indented to use
		/// </summary>
		public FieldListHierarchy() : base()
		{

		}

		#endregion
		
		#region Properties

		/// <summary>
		/// SupportByVersion OWC10 1
		/// Get
		/// </summary>
		[SupportByVersion("OWC10", 1)]
		public virtual NetOffice.OWC10Api.FieldListNode Root
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteKnownReferencePropertyGet<NetOffice.OWC10Api.FieldListNode>(this, "Root", typeof(NetOffice.OWC10Api.FieldListNode));
			}
		}

		/// <summary>
		/// SupportByVersion OWC10 1
		/// Get/Set
		/// </summary>
		[SupportByVersion("OWC10", 1)]
		public virtual bool Visible
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteBoolPropertyGet(this, "Visible");
			}
			set
			{
				InvokerService.InvokeInternal.ExecuteValuePropertySet(this, "Visible", value);
			}
		}

		/// <summary>
		/// SupportByVersion OWC10 1
		/// Get
		/// </summary>
		[SupportByVersion("OWC10", 1)]
		public virtual NetOffice.OWC10Api.FieldListNode Selection
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteKnownReferencePropertyGet<NetOffice.OWC10Api.FieldListNode>(this, "Selection", typeof(NetOffice.OWC10Api.FieldListNode));
			}
		}

		/// <summary>
		/// SupportByVersion OWC10 1
		/// Get/Set
		/// </summary>
		[SupportByVersion("OWC10", 1)]
		public virtual bool ConcatenateData
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteBoolPropertyGet(this, "ConcatenateData");
			}
			set
			{
				InvokerService.InvokeInternal.ExecuteValuePropertySet(this, "ConcatenateData", value);
			}
		}

		/// <summary>
		/// SupportByVersion OWC10 1
		/// Get/Set
		/// </summary>
		[SupportByVersion("OWC10", 1)]
		public virtual string DataSeparator
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteStringPropertyGet(this, "DataSeparator");
			}
			set
			{
				InvokerService.InvokeInternal.ExecuteValuePropertySet(this, "DataSeparator", value);
			}
		}

		#endregion

		#region Methods

		/// <summary>
		/// SupportByVersion OWC10 1
		/// </summary>
		/// <param name="pflhs">NetOffice.OWC10Api.FieldListHierarchySite pflhs</param>
		[SupportByVersion("OWC10", 1)]
		public virtual void SetHierarchySite(NetOffice.OWC10Api.FieldListHierarchySite pflhs)
		{
			 InvokerService.InvokeInternal.ExecuteMethod(this, "SetHierarchySite", pflhs);
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
		public virtual NetOffice.OWC10Api.FieldListNode AddNode(NetOffice.OWC10Api.FieldListNode pflnParent, bool fInsertFirst, Int32 nID, string bstrName, string bstrData, Int32 nType)
		{
			return InvokerService.InvokeInternal.ExecuteKnownReferenceMethodGet<NetOffice.OWC10Api.FieldListNode>(this, "AddNode", typeof(NetOffice.OWC10Api.FieldListNode), new object[]{ pflnParent, fInsertFirst, nID, bstrName, bstrData, nType });
		}

		/// <summary>
		/// SupportByVersion OWC10 1
		/// </summary>
		/// <param name="nID">Int32 nID</param>
		[SupportByVersion("OWC10", 1)]
		public virtual NetOffice.OWC10Api.FieldListNode GetNode(Int32 nID)
		{
			return InvokerService.InvokeInternal.ExecuteKnownReferenceMethodGet<NetOffice.OWC10Api.FieldListNode>(this, "GetNode", typeof(NetOffice.OWC10Api.FieldListNode), nID);
		}

		/// <summary>
		/// SupportByVersion OWC10 1
		/// </summary>
		/// <param name="pfln">NetOffice.OWC10Api.FieldListNode pfln</param>
		[SupportByVersion("OWC10", 1)]
		public virtual void RemoveNode(NetOffice.OWC10Api.FieldListNode pfln)
		{
			 InvokerService.InvokeInternal.ExecuteMethod(this, "RemoveNode", pfln);
		}

		/// <summary>
		/// SupportByVersion OWC10 1
		/// </summary>
		/// <param name="nType">Int32 nType</param>
		[SupportByVersion("OWC10", 1)]
		public virtual NetOffice.OWC10Api.FieldListType AddType(Int32 nType)
		{
			return InvokerService.InvokeInternal.ExecuteKnownReferenceMethodGet<NetOffice.OWC10Api.FieldListType>(this, "AddType", typeof(NetOffice.OWC10Api.FieldListType), nType);
		}

		/// <summary>
		/// SupportByVersion OWC10 1
		/// </summary>
		/// <param name="nTypeId">Int32 nTypeId</param>
		[SupportByVersion("OWC10", 1)]
		public virtual NetOffice.OWC10Api.FieldListType GetType(Int32 nTypeId)
		{
			return InvokerService.InvokeInternal.ExecuteKnownReferenceMethodGet<NetOffice.OWC10Api.FieldListType>(this, "GetType", typeof(NetOffice.OWC10Api.FieldListType), nTypeId);
		}

		/// <summary>
		/// SupportByVersion OWC10 1
		/// </summary>
		/// <param name="pfln">NetOffice.OWC10Api.FieldListNode pfln</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		[SupportByVersion("OWC10", 1)]
		public virtual NetOffice.OWC10Api.FieldListNode GetNextSelected(NetOffice.OWC10Api.FieldListNode pfln)
		{
			return InvokerService.InvokeInternal.ExecuteKnownReferenceMethodGet<NetOffice.OWC10Api.FieldListNode>(this, "GetNextSelected", typeof(NetOffice.OWC10Api.FieldListNode), pfln);
		}

		#endregion

		#pragma warning restore
	}
}


