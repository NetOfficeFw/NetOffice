using System;
using NetRuntimeSystem = System;
using System.ComponentModel;
using NetOffice.Attributes;

namespace NetOffice.OfficeApi
{
	/// <summary>
	/// DispatchInterface ContactCard 
	/// SupportByVersion Office, 14,15,16
	/// </summary>
	/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff860545.aspx </remarks>
	[SupportByVersion("Office", 14,15,16)]
	[EntityType(EntityType.IsDispatchInterface)]
 	public class ContactCard : _IMsoDispObj
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
                    _type = typeof(ContactCard);
                return _type;
            }
        }
        
        #endregion
        
		#region Ctor

		/// <param name="factory">current used factory core</param>
		/// <param name="parentObject">object there has created the proxy</param>
		/// <param name="proxyShare">proxy share instead if com proxy</param>
		public ContactCard(Core factory, ICOMObject parentObject, COMProxyShare proxyShare) : base(factory, parentObject, proxyShare)
		{
		}

		///<param name="factory">current used factory core</param>
		///<param name="parentObject">object there has created the proxy</param>
        ///<param name="comProxy">inner wrapped COM proxy</param>
		public ContactCard(Core factory, ICOMObject parentObject, object comProxy) : base(factory, parentObject, comProxy)
		{
			
		}

        ///<param name="parentObject">object there has created the proxy</param>
        ///<param name="comProxy">inner wrapped COM proxy</param>
        [EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public ContactCard(ICOMObject parentObject, object comProxy) : base(parentObject, comProxy)
		{
		}
		
		///<param name="factory">current used factory core</param>
		///<param name="parentObject">object there has created the proxy</param>
        ///<param name="comProxy">inner wrapped COM proxy</param>
        ///<param name="comProxyType">Type of inner wrapped COM proxy"</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public ContactCard(Core factory, ICOMObject parentObject, object comProxy, NetRuntimeSystem.Type comProxyType) : base(factory, parentObject, comProxy, comProxyType)
		{

		}

		///<param name="parentObject">object there has created the proxy</param>
        ///<param name="comProxy">inner wrapped COM proxy</param>
        ///<param name="comProxyType">Type of inner wrapped COM proxy"</param>
        [EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public ContactCard(ICOMObject parentObject, object comProxy, NetRuntimeSystem.Type comProxyType) : base(parentObject, comProxy, comProxyType)
		{
		}
		
		///<param name="replacedObject">object to replaced. replacedObject are not usable after this action</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public ContactCard(ICOMObject replacedObject) : base(replacedObject)
		{
		}
		
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public ContactCard() : base()
		{
		}
		
		/// <param name="progId">registered progID</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public ContactCard(string progId) : base(progId)
		{
		}
		
		#endregion
		
		#region Properties

		#endregion

		#region Methods

		/// <summary>
		/// SupportByVersion Office 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff863157.aspx </remarks>
		[SupportByVersion("Office", 14,15,16)]
		public void Close()
		{
			 Factory.ExecuteMethod(this, "Close");
		}

		/// <summary>
		/// SupportByVersion Office 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff861819.aspx </remarks>
		/// <param name="cardStyle">NetOffice.OfficeApi.Enums.MsoContactCardStyle cardStyle</param>
		/// <param name="rectangleLeft">Int32 rectangleLeft</param>
		/// <param name="rectangleRight">Int32 rectangleRight</param>
		/// <param name="rectangleTop">Int32 rectangleTop</param>
		/// <param name="rectangleBottom">Int32 rectangleBottom</param>
		/// <param name="horizontalPosition">Int32 horizontalPosition</param>
		/// <param name="showWithDelay">optional bool ShowWithDelay = false</param>
		[SupportByVersion("Office", 14,15,16)]
		public void Show(NetOffice.OfficeApi.Enums.MsoContactCardStyle cardStyle, Int32 rectangleLeft, Int32 rectangleRight, Int32 rectangleTop, Int32 rectangleBottom, Int32 horizontalPosition, object showWithDelay)
		{
			 Factory.ExecuteMethod(this, "Show", new object[]{ cardStyle, rectangleLeft, rectangleRight, rectangleTop, rectangleBottom, horizontalPosition, showWithDelay });
		}

		/// <summary>
		/// SupportByVersion Office 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff861819.aspx </remarks>
		/// <param name="cardStyle">NetOffice.OfficeApi.Enums.MsoContactCardStyle cardStyle</param>
		/// <param name="rectangleLeft">Int32 rectangleLeft</param>
		/// <param name="rectangleRight">Int32 rectangleRight</param>
		/// <param name="rectangleTop">Int32 rectangleTop</param>
		/// <param name="rectangleBottom">Int32 rectangleBottom</param>
		/// <param name="horizontalPosition">Int32 horizontalPosition</param>
		[CustomMethod]
		[SupportByVersion("Office", 14,15,16)]
		public void Show(NetOffice.OfficeApi.Enums.MsoContactCardStyle cardStyle, Int32 rectangleLeft, Int32 rectangleRight, Int32 rectangleTop, Int32 rectangleBottom, Int32 horizontalPosition)
		{
			 Factory.ExecuteMethod(this, "Show", new object[]{ cardStyle, rectangleLeft, rectangleRight, rectangleTop, rectangleBottom, horizontalPosition });
		}

		#endregion

		#pragma warning restore
	}
}
