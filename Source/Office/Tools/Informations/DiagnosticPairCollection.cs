using System;
using System.ComponentModel;
using System.Collections.Generic;
using System.Reflection;
using System.Text;

namespace NetOffice.OfficeApi.Tools.Informations
{
    /// <summary>
    /// Analyze the current environment in detail and save the result in its collection
    /// </summary>
    public class DiagnosticPairCollection : List<DiagnosticPair>, ITypedList
    {
        #region Ctor

        /// <summary>
        /// Creates an instance of the class
        /// </summary>
        /// <param name="owner">owner instance</param>
        public DiagnosticPairCollection(Utils.CommonUtils owner)
        {
            if (null == owner)
                throw new ArgumentNullException("owner");
            Owner = owner;
            foreach (DiagnosticPair item in GetSummary())
                Add(item);
        }

        #endregion

        #region Properties

        /// <summary>
        /// Utils Owner Instance
        /// </summary>
        internal Utils.CommonUtils Owner { get; set; }

        #endregion

        #region Methods

        /// <summary>
        /// Analyze current environment
        /// </summary>
        /// <returns>analyze result</returns>
        public IEnumerable<DiagnosticPair> GetSummary()
        {
            List<DiagnosticPair> list = new List<DiagnosticPair>();

            if (null != Owner.Infos.Assembly)
            { 
                foreach (KeyValuePair<string, string> item in Owner.Infos.Assembly)
                    list.Add(new DiagnosticPair(item.Key, item.Value));
            }

            if (null != Owner.Infos.Environment)
            { 
                foreach (KeyValuePair<string, string> item in Owner.Infos.Environment)
                    list.Add(new DiagnosticPair(item.Key, item.Value));
            }

            if (null != Owner.Infos.AppDomain)
            { 
                foreach (KeyValuePair<string, string> item in Owner.Infos.AppDomain)
                    list.Add(new DiagnosticPair(item.Key, item.Value));
            }

            if (null != Owner.Infos.Host)
            {
                foreach (KeyValuePair<string, string> item in Owner.Infos.Host)
                    list.Add(new DiagnosticPair(item.Key, item.Value));            
            }

            Owner.Infos.GetCustomInformations(this);

            return list;
        }

        #endregion

        #region ITypedList

        /// <summary>
        /// <see cref="ITypedList.GetItemProperties"/>
        /// </summary>
        /// <param name="listAccessors">attributes for the type</param>
        /// <returns>property descriptor collection for DiagnosticPair</returns>
        public PropertyDescriptorCollection GetItemProperties(PropertyDescriptor[] listAccessors)
        {
            return TypeDescriptor.GetProperties(typeof(DiagnosticPair));
        }

        /// <summary>
        /// <see cref="ITypedList.GetListName"/>
        /// </summary>
        /// <param name="listAccessors">attributes for the type</param>
        /// <returns>not implemented</returns>
        public string GetListName(PropertyDescriptor[] listAccessors)
        {
            throw new NotImplementedException();
        }

        #endregion
    }
}
