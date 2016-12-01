using System;
using NetRuntimeSystem = System;
using System.Runtime.InteropServices;
using System.Runtime.CompilerServices;
using System.ComponentModel;
using System.Reflection;
using System.Collections.Generic;
using NetOffice;
namespace NetOffice.OfficeApi
{
	///<summary>
	/// DispatchInterface EncryptionProvider 
	/// SupportByVersion Office, 12,14,15,16
	/// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff863389.aspx
	///</summary>
	[SupportByVersionAttribute("Office", 12,14,15,16)]
	[EntityTypeAttribute(EntityType.IsDispatchInterface)]
	public class EncryptionProvider : COMObject
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
                    _type = typeof(EncryptionProvider);
                    
                return _type;
            }
        }
        
        #endregion
        
		#region Construction

		///<param name="factory">current used factory core</param>
		///<param name="parentObject">object there has created the proxy</param>
        ///<param name="comProxy">inner wrapped COM proxy</param>
		public EncryptionProvider(Core factory, ICOMObject parentObject, object comProxy) : base(factory, parentObject, comProxy)
		{
			
		}

        ///<param name="parentObject">object there has created the proxy</param>
        ///<param name="comProxy">inner wrapped COM proxy</param>
        [EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public EncryptionProvider(ICOMObject parentObject, object comProxy) : base(parentObject, comProxy)
		{
		}
		
		///<param name="factory">current used factory core</param>
		///<param name="parentObject">object there has created the proxy</param>
        ///<param name="comProxy">inner wrapped COM proxy</param>
        ///<param name="comProxyType">Type of inner wrapped COM proxy"</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public EncryptionProvider(Core factory, ICOMObject parentObject, object comProxy, NetRuntimeSystem.Type comProxyType) : base(factory, parentObject, comProxy, comProxyType)
		{

		}

		///<param name="parentObject">object there has created the proxy</param>
        ///<param name="comProxy">inner wrapped COM proxy</param>
        ///<param name="comProxyType">Type of inner wrapped COM proxy"</param>
        [EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public EncryptionProvider(ICOMObject parentObject, object comProxy, NetRuntimeSystem.Type comProxyType) : base(parentObject, comProxy, comProxyType)
		{
		}
		
		///<param name="replacedObject">object to replaced. replacedObject are not usable after this action</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public EncryptionProvider(ICOMObject replacedObject) : base(replacedObject)
		{
		}
		
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public EncryptionProvider() : base()
		{
		}
		
		/// <param name="progId">registered ProgID</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public EncryptionProvider(string progId) : base(progId)
		{
		}
		
		#endregion
		
		#region Properties

		#endregion

		#region Methods

		/// <summary>
		/// SupportByVersion Office 12, 14, 15, 16
		/// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff864896.aspx
		/// </summary>
		/// <param name="encprovdet">NetOffice.OfficeApi.Enums.EncryptionProviderDetail encprovdet</param>
		[SupportByVersionAttribute("Office", 12,14,15,16)]
		public object GetProviderDetail(NetOffice.OfficeApi.Enums.EncryptionProviderDetail encprovdet)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(encprovdet);
			object returnItem = Invoker.MethodReturn(this, "GetProviderDetail", paramsArray);
			if((null != returnItem) && (returnItem is MarshalByRefObject))
			{
				ICOMObject newObject = Factory.CreateObjectFromComProxy(this, returnItem);
				return newObject;
			}
			else
			{
				return  returnItem;
			}
		}

		/// <summary>
		/// SupportByVersion Office 12, 14, 15, 16
		/// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff864027.aspx
		/// </summary>
		/// <param name="parentWindow">object ParentWindow</param>
		[SupportByVersionAttribute("Office", 12,14,15,16)]
		public Int32 NewSession(object parentWindow)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(parentWindow);
			object returnItem = Invoker.MethodReturn(this, "NewSession", paramsArray);
			return NetRuntimeSystem.Convert.ToInt32(returnItem);
		}

		/// <summary>
		/// SupportByVersion Office 12, 14, 15, 16
		/// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff864627.aspx
		/// </summary>
		/// <param name="parentWindow">object ParentWindow</param>
		/// <param name="encryptionData">object EncryptionData</param>
		/// <param name="permissionsMask">UIntPtr PermissionsMask</param>
		[SupportByVersionAttribute("Office", 12,14,15,16)]
		public Int32 Authenticate(object parentWindow, object encryptionData, out UIntPtr permissionsMask)
		{
			ParameterModifier[] modifiers = Invoker.CreateParamModifiers(false,false,true);
			permissionsMask = UIntPtr.Zero;
			object[] paramsArray = Invoker.ValidateParamsArray(parentWindow, encryptionData, permissionsMask);
			object returnItem = Invoker.MethodReturn(this, "Authenticate", paramsArray);
			permissionsMask = (UIntPtr)paramsArray[2];
			return NetRuntimeSystem.Convert.ToInt32(returnItem);
		}

		/// <summary>
		/// SupportByVersion Office 12, 14, 15, 16
		/// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff864902.aspx
		/// </summary>
		/// <param name="sessionHandle">Int32 SessionHandle</param>
		[SupportByVersionAttribute("Office", 12,14,15,16)]
		public Int32 CloneSession(Int32 sessionHandle)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(sessionHandle);
			object returnItem = Invoker.MethodReturn(this, "CloneSession", paramsArray);
			return NetRuntimeSystem.Convert.ToInt32(returnItem);
		}

		/// <summary>
		/// SupportByVersion Office 12, 14, 15, 16
		/// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff864652.aspx
		/// </summary>
		/// <param name="sessionHandle">Int32 SessionHandle</param>
		[SupportByVersionAttribute("Office", 12,14,15,16)]
		public void EndSession(Int32 sessionHandle)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(sessionHandle);
			Invoker.Method(this, "EndSession", paramsArray);
		}

		/// <summary>
		/// SupportByVersion Office 12, 14, 15, 16
		/// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff862766.aspx
		/// </summary>
		/// <param name="sessionHandle">Int32 SessionHandle</param>
		/// <param name="encryptionData">object EncryptionData</param>
		[SupportByVersionAttribute("Office", 12,14,15,16)]
		public Int32 Save(Int32 sessionHandle, object encryptionData)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(sessionHandle, encryptionData);
			object returnItem = Invoker.MethodReturn(this, "Save", paramsArray);
			return NetRuntimeSystem.Convert.ToInt32(returnItem);
		}

		/// <summary>
		/// SupportByVersion Office 12, 14, 15, 16
		/// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff861839.aspx
		/// </summary>
		/// <param name="sessionHandle">Int32 SessionHandle</param>
		/// <param name="streamName">string StreamName</param>
		/// <param name="unencryptedStream">object UnencryptedStream</param>
		/// <param name="encryptedStream">object EncryptedStream</param>
		[SupportByVersionAttribute("Office", 12,14,15,16)]
		public void EncryptStream(Int32 sessionHandle, string streamName, object unencryptedStream, object encryptedStream)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(sessionHandle, streamName, unencryptedStream, encryptedStream);
			Invoker.Method(this, "EncryptStream", paramsArray);
		}

		/// <summary>
		/// SupportByVersion Office 12, 14, 15, 16
		/// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff864940.aspx
		/// </summary>
		/// <param name="sessionHandle">Int32 SessionHandle</param>
		/// <param name="streamName">string StreamName</param>
		/// <param name="encryptedStream">object EncryptedStream</param>
		/// <param name="unencryptedStream">object UnencryptedStream</param>
		[SupportByVersionAttribute("Office", 12,14,15,16)]
		public void DecryptStream(Int32 sessionHandle, string streamName, object encryptedStream, object unencryptedStream)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(sessionHandle, streamName, encryptedStream, unencryptedStream);
			Invoker.Method(this, "DecryptStream", paramsArray);
		}

		/// <summary>
		/// SupportByVersion Office 12, 14, 15, 16
		/// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff863378.aspx
		/// </summary>
		/// <param name="sessionHandle">Int32 SessionHandle</param>
		/// <param name="parentWindow">object ParentWindow</param>
		/// <param name="readOnly">bool ReadOnly</param>
		/// <param name="remove">bool Remove</param>
		[SupportByVersionAttribute("Office", 12,14,15,16)]
		public void ShowSettings(Int32 sessionHandle, object parentWindow, bool readOnly, out bool remove)
		{
			ParameterModifier[] modifiers = Invoker.CreateParamModifiers(false,false,false,true);
			remove = false;
			object[] paramsArray = Invoker.ValidateParamsArray(sessionHandle, parentWindow, readOnly, remove);
			Invoker.Method(this, "ShowSettings", paramsArray, modifiers);
			remove = (bool)paramsArray[3];
		}

		#endregion
		#pragma warning restore
	}
}