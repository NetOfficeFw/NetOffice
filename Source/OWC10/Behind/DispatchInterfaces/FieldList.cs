using System;
using NetRuntimeSystem = System;
using System.ComponentModel;
using NetOffice.Attributes;
using NetOffice.OWC10Api;

namespace NetOffice.OWC10Api.Behind
{
	/// <summary>
	/// DispatchInterface FieldList 
	/// SupportByVersion OWC10, 1
	/// </summary>
	[SupportByVersion("OWC10", 1)]
	[EntityType(EntityType.IsDispatchInterface), BaseType]
 	public class FieldList : COMObject, NetOffice.OWC10Api.FieldList
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
                    _contractType = typeof(NetOffice.OWC10Api.FieldList);
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
                    _type = typeof(FieldList);
                return _type;
            }
        }
        
        #endregion
        
		#region Ctor

		/// <summary>
		/// Stub Ctor, not indented to use
		/// </summary>
		public FieldList() : base()
		{

		}

		#endregion
		
		#region Properties

		/// <summary>
		/// SupportByVersion OWC10 1
		/// Get
		/// </summary>
		[SupportByVersion("OWC10", 1)]
		public virtual Int32 ClipboardFormat
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteInt32PropertyGet(this, "ClipboardFormat");
			}
		}

		/// <summary>
		/// SupportByVersion OWC10 1
		/// Get
		/// </summary>
		[SupportByVersion("OWC10", 1)]
		public virtual string InstanceID
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteStringPropertyGet(this, "InstanceID");
			}
		}

		/// <summary>
		/// SupportByVersion OWC10 1
		/// Get/Set
		/// </summary>
		[SupportByVersion("OWC10", 1), NativeResult]
		public virtual stdole.IFont Font
		{
			get
			{
				object[] paramsArray = null;
				object returnItem = Invoker.PropertyGet(this, "Font", paramsArray);
                return returnItem as stdole.IFont;
			}
			set
			{
				object[] paramsArray = Invoker.ValidateParamsArray(value);
				Invoker.PropertySet(this, "Font", paramsArray);
			}
		}

		/// <summary>
		/// SupportByVersion OWC10 1
		/// Get/Set
		/// </summary>
		[SupportByVersion("OWC10", 1)]
		public virtual bool MultiSelect
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteBoolPropertyGet(this, "MultiSelect");
			}
			set
			{
				InvokerService.InvokeInternal.ExecuteValuePropertySet(this, "MultiSelect", value);
			}
		}

		/// <summary>
		/// SupportByVersion OWC10 1
		/// Get/Set
		/// </summary>
		[SupportByVersion("OWC10", 1)]
		public virtual NetOffice.OWC10Api.Enums.FieldListSelectRestriction SelectRestriction
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteEnumPropertyGet<NetOffice.OWC10Api.Enums.FieldListSelectRestriction>(this, "SelectRestriction");
			}
			set
			{
				InvokerService.InvokeInternal.ExecuteEnumPropertySet(this, "SelectRestriction", value);
			}
		}

		#endregion

		#region Methods

		/// <summary>
		/// SupportByVersion OWC10 1
		/// </summary>
		/// <param name="bVisible">bool bVisible</param>
		[SupportByVersion("OWC10", 1)]
		public virtual NetOffice.OWC10Api.FieldListHierarchy CreateHierarchy(bool bVisible)
		{
			return InvokerService.InvokeInternal.ExecuteKnownReferenceMethodGet<NetOffice.OWC10Api.FieldListHierarchy>(this, "CreateHierarchy", typeof(NetOffice.OWC10Api.FieldListHierarchy), bVisible);
		}

		/// <summary>
		/// SupportByVersion OWC10 1
		/// </summary>
		/// <param name="iWidth">Int32 iWidth</param>
		/// <param name="iHeight">Int32 iHeight</param>
		/// <param name="pip">stdole.IPicture pip</param>
		/// <param name="crMask">Int32 crMask</param>
		[SupportByVersion("OWC10", 1)]
		public virtual Int32 AddBitmap(Int32 iWidth, Int32 iHeight, stdole.IPicture pip, Int32 crMask)
		{
			return InvokerService.InvokeInternal.ExecuteInt32MethodGet(this, "AddBitmap", iWidth, iHeight, pip, crMask);
		}

		/// <summary>
		/// SupportByVersion OWC10 1
		/// </summary>
		/// <param name="pfln">NetOffice.OWC10Api.FieldListNode pfln</param>
		[SupportByVersion("OWC10", 1)]
		public virtual NetOffice.OWC10Api.FieldListNode GetNextSelected(NetOffice.OWC10Api.FieldListNode pfln)
		{
			return InvokerService.InvokeInternal.ExecuteKnownReferenceMethodGet<NetOffice.OWC10Api.FieldListNode>(this, "GetNextSelected", typeof(NetOffice.OWC10Api.FieldListNode), pfln);
		}

		/// <summary>
		/// SupportByVersion OWC10 1
		/// </summary>
		[SupportByVersion("OWC10", 1)]
		public virtual void ClearSelection()
		{
			 InvokerService.InvokeInternal.ExecuteMethod(this, "ClearSelection");
		}

		/// <summary>
		/// SupportByVersion OWC10 1
		/// </summary>
		/// <param name="iImage">Int32 iImage</param>
		/// <param name="iOverlay">Int32 iOverlay</param>
		[SupportByVersion("OWC10", 1)]
		public virtual void SetOverlayImage(Int32 iImage, Int32 iOverlay)
		{
			 InvokerService.InvokeInternal.ExecuteMethod(this, "SetOverlayImage", iImage, iOverlay);
		}

		#endregion

		#pragma warning restore
	}
}


