using System;
using NetRuntimeSystem = System;
using System.ComponentModel;
using NetOffice.Attributes;

namespace NetOffice.WordApi
{
	/// <summary>
	/// DispatchInterface RevisionsFilter 
	/// SupportByVersion Word, 15, 16
	/// </summary>
	/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/jj228191.aspx </remarks>
	[SupportByVersion("Word", 15, 16)]
	[EntityType(EntityType.IsDispatchInterface)]
 	public class RevisionsFilter : COMObject
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
                    _type = typeof(RevisionsFilter);
                return _type;
            }
        }
        
        #endregion
        
		#region Ctor

		/// <param name="factory">current used factory core</param>
		/// <param name="parentObject">object there has created the proxy</param>
		/// <param name="proxyShare">proxy share instead if com proxy</param>
		public RevisionsFilter(Core factory, ICOMObject parentObject, COMProxyShare proxyShare) : base(factory, parentObject, proxyShare)
		{
		}

		///<param name="factory">current used factory core</param>
		///<param name="parentObject">object there has created the proxy</param>
        ///<param name="comProxy">inner wrapped COM proxy</param>
		public RevisionsFilter(Core factory, ICOMObject parentObject, object comProxy) : base(factory, parentObject, comProxy)
		{
			
		}

        ///<param name="parentObject">object there has created the proxy</param>
        ///<param name="comProxy">inner wrapped COM proxy</param>
        [EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public RevisionsFilter(ICOMObject parentObject, object comProxy) : base(parentObject, comProxy)
		{
		}
		
		///<param name="factory">current used factory core</param>
		///<param name="parentObject">object there has created the proxy</param>
        ///<param name="comProxy">inner wrapped COM proxy</param>
        ///<param name="comProxyType">Type of inner wrapped COM proxy"</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public RevisionsFilter(Core factory, ICOMObject parentObject, object comProxy, NetRuntimeSystem.Type comProxyType) : base(factory, parentObject, comProxy, comProxyType)
		{

		}

		///<param name="parentObject">object there has created the proxy</param>
        ///<param name="comProxy">inner wrapped COM proxy</param>
        ///<param name="comProxyType">Type of inner wrapped COM proxy"</param>
        [EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public RevisionsFilter(ICOMObject parentObject, object comProxy, NetRuntimeSystem.Type comProxyType) : base(parentObject, comProxy, comProxyType)
		{
		}
		
		///<param name="replacedObject">object to replaced. replacedObject are not usable after this action</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public RevisionsFilter(ICOMObject replacedObject) : base(replacedObject)
		{
		}
		
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public RevisionsFilter() : base()
		{
		}
		
		/// <param name="progId">registered progID</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public RevisionsFilter(string progId) : base(progId)
		{
		}
		
		#endregion
		
		#region Properties

		/// <summary>
		/// SupportByVersion Word 15,16
		/// Get/Set
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/jj231521.aspx </remarks>
		[SupportByVersion("Word", 15, 16)]
		public NetOffice.WordApi.Enums.WdRevisionsView View
		{
			get
			{
				return Factory.ExecuteEnumPropertyGet<NetOffice.WordApi.Enums.WdRevisionsView>(this, "View");
			}
			set
			{
				Factory.ExecuteEnumPropertySet(this, "View", value);
			}
		}

		/// <summary>
		/// SupportByVersion Word 15,16
		/// Get/Set
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/jj230964.aspx </remarks>
		[SupportByVersion("Word", 15, 16)]
		public NetOffice.WordApi.Enums.WdRevisionsMarkup Markup
		{
			get
			{
				return Factory.ExecuteEnumPropertyGet<NetOffice.WordApi.Enums.WdRevisionsMarkup>(this, "Markup");
			}
			set
			{
				Factory.ExecuteEnumPropertySet(this, "Markup", value);
			}
		}

		/// <summary>
		/// SupportByVersion Word 15,16
		/// Get
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/jj231708.aspx </remarks>
		[SupportByVersion("Word", 15, 16)]
		public NetOffice.WordApi.Reviewers Reviewers
		{
			get
			{
				return Factory.ExecuteKnownReferencePropertyGet<NetOffice.WordApi.Reviewers>(this, "Reviewers", NetOffice.WordApi.Reviewers.LateBindingApiWrapperType);
			}
		}

		#endregion

		#region Methods

		/// <summary>
		/// SupportByVersion Word 15,16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/jj227337.aspx </remarks>
		[SupportByVersion("Word", 15, 16)]
		public void ToggleShowAllReviewers()
		{
			 Factory.ExecuteMethod(this, "ToggleShowAllReviewers");
		}

		#endregion

		#pragma warning restore
	}
}
