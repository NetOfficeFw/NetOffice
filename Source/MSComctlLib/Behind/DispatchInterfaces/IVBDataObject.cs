using System;
using NetRuntimeSystem = System;
using System.ComponentModel;
using NetOffice.Attributes;
using NetOffice.MSComctlLibApi;

namespace NetOffice.MSComctlLibApi.Behind
{
	/// <summary>
	/// DispatchInterface IVBDataObject 
	/// SupportByVersion MSComctlLib, 6
	/// </summary>
	[SupportByVersion("MSComctlLib", 6)]
	[EntityType(EntityType.IsDispatchInterface), BaseType]
 	public class IVBDataObject : COMObject, NetOffice.MSComctlLibApi.IVBDataObject
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
                    _contractType = typeof(NetOffice.MSComctlLibApi.IVBDataObject);
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
                    _type = typeof(IVBDataObject);
                return _type;
            }
        }
        
        #endregion
        
		#region Ctor

		/// <summary>
		/// Stub Ctor, not indented to use
		/// </summary>
		public IVBDataObject() : base()
		{

		}

		#endregion
		
		#region Properties

		/// <summary>
		/// SupportByVersion MSComctlLib 6
		/// Get
		/// </summary>
		[SupportByVersion("MSComctlLib", 6)]
		[BaseResult]
		public virtual NetOffice.MSComctlLibApi.IVBDataObjectFiles Files
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteBaseReferencePropertyGet<NetOffice.MSComctlLibApi.IVBDataObjectFiles>(this, "Files");
			}
		}

		#endregion

		#region Methods

		/// <summary>
		/// SupportByVersion MSComctlLib 6
		/// </summary>
		[SupportByVersion("MSComctlLib", 6)]
		public virtual void Clear()
		{
			 InvokerService.InvokeInternal.ExecuteMethod(this, "Clear");
		}

		/// <summary>
		/// SupportByVersion MSComctlLib 6
		/// </summary>
		/// <param name="sFormat">Int16 sFormat</param>
		[SupportByVersion("MSComctlLib", 6)]
		public virtual object GetData(Int16 sFormat)
		{
			return InvokerService.InvokeInternal.ExecuteVariantMethodGet(this, "GetData", sFormat);
		}

		/// <summary>
		/// SupportByVersion MSComctlLib 6
		/// </summary>
		/// <param name="sFormat">Int16 sFormat</param>
		[SupportByVersion("MSComctlLib", 6)]
		public virtual bool GetFormat(Int16 sFormat)
		{
			return InvokerService.InvokeInternal.ExecuteBoolMethodGet(this, "GetFormat", sFormat);
		}

		/// <summary>
		/// SupportByVersion MSComctlLib 6
		/// </summary>
		/// <param name="vValue">optional object vValue</param>
		/// <param name="vFormat">optional object vFormat</param>
		[SupportByVersion("MSComctlLib", 6)]
		public virtual void SetData(object vValue, object vFormat)
		{
			 InvokerService.InvokeInternal.ExecuteMethod(this, "SetData", vValue, vFormat);
		}

		/// <summary>
		/// SupportByVersion MSComctlLib 6
		/// </summary>
		[CustomMethod]
		[SupportByVersion("MSComctlLib", 6)]
		public virtual void SetData()
		{
			 InvokerService.InvokeInternal.ExecuteMethod(this, "SetData");
		}

		/// <summary>
		/// SupportByVersion MSComctlLib 6
		/// </summary>
		/// <param name="vValue">optional object vValue</param>
		[CustomMethod]
		[SupportByVersion("MSComctlLib", 6)]
		public virtual void SetData(object vValue)
		{
			 InvokerService.InvokeInternal.ExecuteMethod(this, "SetData", vValue);
		}

		#endregion

		#pragma warning restore
	}
}

