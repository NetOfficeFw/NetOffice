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
	/// DispatchInterface ContactCard 
	/// SupportByVersion Office, 14,15,16
	/// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff860545.aspx
	///</summary>
	[SupportByVersionAttribute("Office", 14,15,16)]
	[EntityTypeAttribute(EntityType.IsDispatchInterface)]
	public class ContactCard : _IMsoDispObj
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
                    _type = typeof(ContactCard);
                    
                return _type;
            }
        }
        
        #endregion
        
		#region Construction

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
		
		/// <param name="progId">registered ProgID</param>
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
		/// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff863157.aspx
		/// </summary>
		[SupportByVersionAttribute("Office", 14,15,16)]
		public void Close()
		{
			object[] paramsArray = null;
			Invoker.Method(this, "Close", paramsArray);
		}

		/// <summary>
		/// SupportByVersion Office 14, 15, 16
		/// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff861819.aspx
		/// </summary>
		/// <param name="cardStyle">NetOffice.OfficeApi.Enums.MsoContactCardStyle CardStyle</param>
		/// <param name="rectangleLeft">Int32 RectangleLeft</param>
		/// <param name="rectangleRight">Int32 RectangleRight</param>
		/// <param name="rectangleTop">Int32 RectangleTop</param>
		/// <param name="rectangleBottom">Int32 RectangleBottom</param>
		/// <param name="horizontalPosition">Int32 HorizontalPosition</param>
		/// <param name="showWithDelay">optional bool ShowWithDelay = false</param>
		[SupportByVersionAttribute("Office", 14,15,16)]
		public void Show(NetOffice.OfficeApi.Enums.MsoContactCardStyle cardStyle, Int32 rectangleLeft, Int32 rectangleRight, Int32 rectangleTop, Int32 rectangleBottom, Int32 horizontalPosition, object showWithDelay)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(cardStyle, rectangleLeft, rectangleRight, rectangleTop, rectangleBottom, horizontalPosition, showWithDelay);
			Invoker.Method(this, "Show", paramsArray);
		}

		/// <summary>
		/// SupportByVersion Office 14, 15, 16
		/// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff861819.aspx
		/// </summary>
		/// <param name="cardStyle">NetOffice.OfficeApi.Enums.MsoContactCardStyle CardStyle</param>
		/// <param name="rectangleLeft">Int32 RectangleLeft</param>
		/// <param name="rectangleRight">Int32 RectangleRight</param>
		/// <param name="rectangleTop">Int32 RectangleTop</param>
		/// <param name="rectangleBottom">Int32 RectangleBottom</param>
		/// <param name="horizontalPosition">Int32 HorizontalPosition</param>
		[CustomMethodAttribute]
		[SupportByVersionAttribute("Office", 14,15,16)]
		public void Show(NetOffice.OfficeApi.Enums.MsoContactCardStyle cardStyle, Int32 rectangleLeft, Int32 rectangleRight, Int32 rectangleTop, Int32 rectangleBottom, Int32 horizontalPosition)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(cardStyle, rectangleLeft, rectangleRight, rectangleTop, rectangleBottom, horizontalPosition);
			Invoker.Method(this, "Show", paramsArray);
		}

		#endregion
		#pragma warning restore
	}
}