using System;
using NetRuntimeSystem = System;
using System.ComponentModel;
using NetOffice.Attributes;

namespace NetOffice.ExcelApi
{
	/// <summary>
	/// DispatchInterface Speech 
	/// SupportByVersion Excel, 10,11,12,14,15,16
	/// </summary>
	/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff194865.aspx </remarks>
	[SupportByVersion("Excel", 10,11,12,14,15,16)]
	[EntityType(EntityType.IsDispatchInterface)]
 	public class Speech : COMObject
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
                    _type = typeof(Speech);
                return _type;
            }
        }
        
        #endregion
        
		#region Ctor

		/// <param name="factory">current used factory core</param>
		/// <param name="parentObject">object there has created the proxy</param>
		/// <param name="proxyShare">proxy share instead if com proxy</param>
		public Speech(Core factory, ICOMObject parentObject, COMProxyShare proxyShare) : base(factory, parentObject, proxyShare)
		{
		}

		///<param name="factory">current used factory core</param>
		///<param name="parentObject">object there has created the proxy</param>
        ///<param name="comProxy">inner wrapped COM proxy</param>
		public Speech(Core factory, ICOMObject parentObject, object comProxy) : base(factory, parentObject, comProxy)
		{
			
		}

        ///<param name="parentObject">object there has created the proxy</param>
        ///<param name="comProxy">inner wrapped COM proxy</param>
        [EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public Speech(ICOMObject parentObject, object comProxy) : base(parentObject, comProxy)
		{
		}
		
		///<param name="factory">current used factory core</param>
		///<param name="parentObject">object there has created the proxy</param>
        ///<param name="comProxy">inner wrapped COM proxy</param>
        ///<param name="comProxyType">Type of inner wrapped COM proxy"</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public Speech(Core factory, ICOMObject parentObject, object comProxy, NetRuntimeSystem.Type comProxyType) : base(factory, parentObject, comProxy, comProxyType)
		{

		}

		///<param name="parentObject">object there has created the proxy</param>
        ///<param name="comProxy">inner wrapped COM proxy</param>
        ///<param name="comProxyType">Type of inner wrapped COM proxy"</param>
        [EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public Speech(ICOMObject parentObject, object comProxy, NetRuntimeSystem.Type comProxyType) : base(parentObject, comProxy, comProxyType)
		{
		}
		
		///<param name="replacedObject">object to replaced. replacedObject are not usable after this action</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public Speech(ICOMObject replacedObject) : base(replacedObject)
		{
		}
		
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public Speech() : base()
		{
		}
		
		/// <param name="progId">registered progID</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public Speech(string progId) : base(progId)
		{
		}
		
		#endregion
		
		#region Properties

		/// <summary>
		/// SupportByVersion Excel 10, 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff835598.aspx </remarks>
		[SupportByVersion("Excel", 10,11,12,14,15,16)]
		public NetOffice.ExcelApi.Enums.XlSpeakDirection Direction
		{
			get
			{
				return Factory.ExecuteEnumPropertyGet<NetOffice.ExcelApi.Enums.XlSpeakDirection>(this, "Direction");
			}
			set
			{
				Factory.ExecuteEnumPropertySet(this, "Direction", value);
			}
		}

		/// <summary>
		/// SupportByVersion Excel 10, 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff837122.aspx </remarks>
		[SupportByVersion("Excel", 10,11,12,14,15,16)]
		public bool SpeakCellOnEnter
		{
			get
			{
				return Factory.ExecuteBoolPropertyGet(this, "SpeakCellOnEnter");
			}
			set
			{
				Factory.ExecuteValuePropertySet(this, "SpeakCellOnEnter", value);
			}
		}

		#endregion

		#region Methods

		/// <summary>
		/// SupportByVersion Excel 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff839393.aspx </remarks>
		/// <param name="text">string text</param>
		/// <param name="speakAsync">optional object speakAsync</param>
		/// <param name="speakXML">optional object speakXML</param>
		/// <param name="purge">optional object purge</param>
		[SupportByVersion("Excel", 10,11,12,14,15,16)]
		public void Speak(string text, object speakAsync, object speakXML, object purge)
		{
			 Factory.ExecuteMethod(this, "Speak", text, speakAsync, speakXML, purge);
		}

		/// <summary>
		/// SupportByVersion Excel 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff839393.aspx </remarks>
		/// <param name="text">string text</param>
		[CustomMethod]
		[SupportByVersion("Excel", 10,11,12,14,15,16)]
		public void Speak(string text)
		{
			 Factory.ExecuteMethod(this, "Speak", text);
		}

		/// <summary>
		/// SupportByVersion Excel 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff839393.aspx </remarks>
		/// <param name="text">string text</param>
		/// <param name="speakAsync">optional object speakAsync</param>
		[CustomMethod]
		[SupportByVersion("Excel", 10,11,12,14,15,16)]
		public void Speak(string text, object speakAsync)
		{
			 Factory.ExecuteMethod(this, "Speak", text, speakAsync);
		}

		/// <summary>
		/// SupportByVersion Excel 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff839393.aspx </remarks>
		/// <param name="text">string text</param>
		/// <param name="speakAsync">optional object speakAsync</param>
		/// <param name="speakXML">optional object speakXML</param>
		[CustomMethod]
		[SupportByVersion("Excel", 10,11,12,14,15,16)]
		public void Speak(string text, object speakAsync, object speakXML)
		{
			 Factory.ExecuteMethod(this, "Speak", text, speakAsync, speakXML);
		}

		#endregion

		#pragma warning restore
	}
}
