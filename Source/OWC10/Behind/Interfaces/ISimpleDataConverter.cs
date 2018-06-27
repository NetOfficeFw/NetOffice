using System;
using NetRuntimeSystem = System;
using System.ComponentModel;
using NetOffice.Attributes;
using NetOffice.OWC10Api;

namespace NetOffice.OWC10Api.Behind
{
	/// <summary>
	/// Interface ISimpleDataConverter 
	/// SupportByVersion OWC10, 1
	/// </summary>
	[SupportByVersion("OWC10", 1)]
	[EntityType(EntityType.IsInterface)]
 	public class ISimpleDataConverter : COMObject, NetOffice.OWC10Api.ISimpleDataConverter
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
                    _contractType = typeof(NetOffice.OWC10Api.ISimpleDataConverter);
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
                    _type = typeof(ISimpleDataConverter);
                return _type;
            }
        }
        
        #endregion
        
		#region Ctor

		/// <summary>
		/// Stub Ctor, not indented to use
		/// </summary>
		public ISimpleDataConverter() : base()
		{

		}

		#endregion
		
		#region Properties

		#endregion

		#region Methods

		/// <summary>
		/// SupportByVersion OWC10 1
		/// </summary>
		/// <param name="varSrc">object varSrc</param>
		/// <param name="vtDest">Int32 vtDest</param>
		/// <param name="pUnknownElement">object pUnknownElement</param>
		/// <param name="pvarDest">object pvarDest</param>
		[SupportByVersion("OWC10", 1)]
		public virtual Int32 ConvertData(object varSrc, Int32 vtDest, object pUnknownElement, object pvarDest)
		{
			return InvokerService.InvokeInternal.ExecuteInt32MethodGet(this, "ConvertData", varSrc, vtDest, pUnknownElement, pvarDest);
		}

		/// <summary>
		/// SupportByVersion OWC10 1
		/// </summary>
		/// <param name="vt1">Int32 vt1</param>
		/// <param name="vt2">Int32 vt2</param>
		[SupportByVersion("OWC10", 1)]
		public virtual Int32 CanConvertData(Int32 vt1, Int32 vt2)
		{
			return InvokerService.InvokeInternal.ExecuteInt32MethodGet(this, "CanConvertData", vt1, vt2);
		}

		#endregion

		#pragma warning restore
	}
}

