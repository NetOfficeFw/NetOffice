﻿using System.Reflection;
using System;
using NetRuntimeSystem = System;
using System.ComponentModel;
using NetOffice.Attributes;

namespace NetOffice.OfficeApi
{
	/// <summary>
	/// Interface IConverterPreferences 
	/// SupportByVersion Office, 14,15,16
	/// </summary>
	/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/Office.IConverterPreferences"/> </remarks>
	[SupportByVersion("Office", 14,15,16)]
	[EntityType(EntityType.IsInterface)]
 	public class IConverterPreferences : COMObject
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
                    _type = typeof(IConverterPreferences);
                return _type;
            }
        }
        
        #endregion
        
		#region Ctor

		/// <param name="factory">current used factory core</param>
		/// <param name="parentObject">object there has created the proxy</param>
		/// <param name="proxyShare">proxy share instead if com proxy</param>
		public IConverterPreferences(Core factory, ICOMObject parentObject, COMProxyShare proxyShare) : base(factory, parentObject, proxyShare)
		{
		}

		///<param name="factory">current used factory core</param>
		///<param name="parentObject">object there has created the proxy</param>
        ///<param name="comProxy">inner wrapped COM proxy</param>
		public IConverterPreferences(Core factory, ICOMObject parentObject, object comProxy) : base(factory, parentObject, comProxy)
		{
			
		}

        ///<param name="parentObject">object there has created the proxy</param>
        ///<param name="comProxy">inner wrapped COM proxy</param>
        [EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public IConverterPreferences(ICOMObject parentObject, object comProxy) : base(parentObject, comProxy)
		{
		}
		
		///<param name="factory">current used factory core</param>
		///<param name="parentObject">object there has created the proxy</param>
        ///<param name="comProxy">inner wrapped COM proxy</param>
        ///<param name="comProxyType">Type of inner wrapped COM proxy"</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public IConverterPreferences(Core factory, ICOMObject parentObject, object comProxy, NetRuntimeSystem.Type comProxyType) : base(factory, parentObject, comProxy, comProxyType)
		{

		}

		///<param name="parentObject">object there has created the proxy</param>
        ///<param name="comProxy">inner wrapped COM proxy</param>
        ///<param name="comProxyType">Type of inner wrapped COM proxy"</param>
        [EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public IConverterPreferences(ICOMObject parentObject, object comProxy, NetRuntimeSystem.Type comProxyType) : base(parentObject, comProxy, comProxyType)
		{
		}
		
		///<param name="replacedObject">object to replaced. replacedObject are not usable after this action</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public IConverterPreferences(ICOMObject replacedObject) : base(replacedObject)
		{
		}
		
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public IConverterPreferences() : base()
		{
		}
		
		/// <param name="progId">registered progID</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public IConverterPreferences(string progId) : base(progId)
		{
		}
		
		#endregion
		
		#region Properties

		#endregion

		#region Methods

		/// <summary>
		/// SupportByVersion Office 14, 15, 16
		/// </summary>
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/Office.IConverterPreferences.HrGetMacroEnabled"/> </remarks>
		/// <param name="pfMacroEnabled">Int32 pfMacroEnabled</param>
		[SupportByVersion("Office", 14,15,16)]
		public Int32 HrGetMacroEnabled(out Int32 pfMacroEnabled)
		{
			ParameterModifier[] modifiers = Invoker.CreateParamModifiers(true);
			pfMacroEnabled = 0;
			object[] paramsArray = Invoker.ValidateParamsArray(pfMacroEnabled);
			object returnItem = Invoker.MethodReturn(this, "HrGetMacroEnabled", paramsArray, modifiers);
			pfMacroEnabled = (Int32)paramsArray[0];
			return NetRuntimeSystem.Convert.ToInt32(returnItem);
		}

		/// <summary>
		/// SupportByVersion Office 14, 15, 16
		/// </summary>
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/Office.IConverterPreferences.HrCheckFormat"/> </remarks>
		/// <param name="pFormat">Int32 pFormat</param>
		[SupportByVersion("Office", 14,15,16)]
		public Int32 HrCheckFormat(out Int32 pFormat)
		{
			ParameterModifier[] modifiers = Invoker.CreateParamModifiers(true);
			pFormat = 0;
			object[] paramsArray = Invoker.ValidateParamsArray(pFormat);
			object returnItem = Invoker.MethodReturn(this, "HrCheckFormat", paramsArray, modifiers);
			pFormat = (Int32)paramsArray[0];
			return NetRuntimeSystem.Convert.ToInt32(returnItem);
		}

		/// <summary>
		/// SupportByVersion Office 14, 15, 16
		/// </summary>
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/Office.IConverterPreferences.HrGetLossySave"/> </remarks>
		/// <param name="pfLossySave">Int32 pfLossySave</param>
		[SupportByVersion("Office", 14,15,16)]
		public Int32 HrGetLossySave(out Int32 pfLossySave)
		{
			ParameterModifier[] modifiers = Invoker.CreateParamModifiers(true);
			pfLossySave = 0;
			object[] paramsArray = Invoker.ValidateParamsArray(pfLossySave);
			object returnItem = Invoker.MethodReturn(this, "HrGetLossySave", paramsArray, modifiers);
			pfLossySave = (Int32)paramsArray[0];
			return NetRuntimeSystem.Convert.ToInt32(returnItem);
		}

		#endregion

		#pragma warning restore
	}
}
