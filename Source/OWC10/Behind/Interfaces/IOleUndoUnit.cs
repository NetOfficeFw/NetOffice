using System.Reflection;
using System;
using NetRuntimeSystem = System;
using System.ComponentModel;
using NetOffice.Attributes;
using NetOffice.OWC10Api;

namespace NetOffice.OWC10Api.Behind
{
	/// <summary>
	/// Interface IOleUndoUnit 
	/// SupportByVersion OWC10, 1
	/// </summary>
	[SupportByVersion("OWC10", 1)]
	[EntityType(EntityType.IsInterface), BaseType]
 	public class IOleUndoUnit : COMObject, NetOffice.OWC10Api.IOleUndoUnit
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
                    _contractType = typeof(NetOffice.OWC10Api.IOleUndoUnit);
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
                    _type = typeof(IOleUndoUnit);
                return _type;
            }
        }
        
        #endregion
        
		#region Ctor

		/// <summary>
		/// Stub Ctor, not indented to use
		/// </summary>
		public IOleUndoUnit() : base()
		{

		}

		#endregion
		
		#region Properties

		#endregion

		#region Methods

		/// <summary>
		/// SupportByVersion OWC10 1
		/// </summary>
		/// <param name="pUndoManager">NetOffice.OWC10Api.IOleUndoManager pUndoManager</param>
		[SupportByVersion("OWC10", 1)]
		public virtual Int32 Do(NetOffice.OWC10Api.IOleUndoManager pUndoManager)
		{
			return InvokerService.InvokeInternal.ExecuteInt32MethodGet(this, "Do", pUndoManager);
		}

		/// <summary>
		/// SupportByVersion OWC10 1
		/// </summary>
		/// <param name="pbstr">string pbstr</param>
		[SupportByVersion("OWC10", 1)]
		public virtual Int32 GetDescription(out string pbstr)
		{
			ParameterModifier[] modifiers = Invoker.CreateParamModifiers(true);
			pbstr = string.Empty;
			object[] paramsArray = Invoker.ValidateParamsArray(pbstr);
			object returnItem = Invoker.MethodReturn(this, "GetDescription", paramsArray, modifiers);
			pbstr = paramsArray[0] as string;
			return NetRuntimeSystem.Convert.ToInt32(returnItem);
		}

		/// <summary>
		/// SupportByVersion OWC10 1
		/// </summary>
		/// <param name="pClsid">Guid pClsid</param>
		/// <param name="plID">Int32 plID</param>
		[SupportByVersion("OWC10", 1)]
		public virtual Int32 GetUnitType(out Guid pClsid, out Int32 plID)
		{
			ParameterModifier[] modifiers = Invoker.CreateParamModifiers(true,true);
			pClsid = Guid.Empty;
			plID = 0;
			object[] paramsArray = Invoker.ValidateParamsArray(pClsid, plID);
			object returnItem = Invoker.MethodReturn(this, "GetUnitType", paramsArray, modifiers);
			pClsid = (Guid)paramsArray[0];
			plID = (Int32)paramsArray[1];
			return NetRuntimeSystem.Convert.ToInt32(returnItem);
		}

		/// <summary>
		/// SupportByVersion OWC10 1
		/// </summary>
		[SupportByVersion("OWC10", 1)]
		public virtual Int32 OnNextAdd()
		{
			return InvokerService.InvokeInternal.ExecuteInt32MethodGet(this, "OnNextAdd");
		}

		#endregion

		#pragma warning restore
	}
}

