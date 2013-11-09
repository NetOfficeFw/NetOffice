using System;
using System.ComponentModel;
using System.Reflection;
using System.Collections.Generic;
using System.Linq;
using System.Xml.Linq;
using System.Text;

using CM = NOTools.ComponentModel;

namespace NOToolsTests.CSharpTextEditor2.DataLayer
{
    /// <summary>
    /// Represents a top-level root table
    /// </summary>
    public class RootList : DataList<RootItem>
    {
        #region Ctor
        
        /// <summary>
        /// Creates an instance of the class
        /// </summary>
        /// <param name="listName">unique name of the list</param>
        public RootList(string listName) : base(listName)
        {
            Schema = XDocument.Parse(ReadXmlString(listName + ".xsd"));
        }

        #endregion

        #region Properties
        
        /// <summary>
        /// Schema Definition
        /// </summary>
        private XDocument Schema { get; set; }

        #endregion

        #region Methods

        /// <summary>
        /// Reads xml data as string from ebebedded ressource
        /// </summary>
        /// <param name="ressourcePath">path to ressource</param>
        /// <returns>xml string</returns>
        protected internal string ReadXmlString(string ressourcePath)
        {
            System.IO.Stream ressourceStream = null;
            System.IO.StreamReader textStreamReader = null;
            try
            {
                Assembly assembly = this.GetType().Assembly;
                ressourceStream = assembly.GetManifestResourceStream(assembly.GetName().Name + ".Data." + ressourcePath);
                if (ressourceStream == null)
                    throw (new System.IO.IOException("Error accessing resource Stream."));

                textStreamReader = new System.IO.StreamReader(ressourceStream);
                if (textStreamReader == null)
                    throw (new System.IO.IOException("Error accessing resource File."));

                string text = textStreamReader.ReadToEnd();
                return text;
            }
            catch (Exception exception)
            {
                throw (exception);
            }
            finally
            {
                if (null != textStreamReader)
                    textStreamReader.Close();
                if (null != ressourceStream)
                    ressourceStream.Close();
            }
        }

        #endregion

        #region Overrides

        public override PropertyDescriptorCollection GetItemProperties(PropertyDescriptor[] listAccessors)
        {
            List<PropertyDescriptor> list = new List<PropertyDescriptor>();

            XElement rootNode = ((Schema.FirstNode as XElement).FirstNode as XElement);
            foreach (var item in rootNode.Element("sequence").Elements("element"))
                list.Add(new RootPropertyDescriptor(item));

            return new PropertyDescriptorCollection(list.ToArray());
        }

        public override RootItem OnGetThisIndexerItem(int index)
        {
            if (!IsLoaded)
                LoadFromDatabase();
            return null;
        }

        public override IEnumerator<RootItem> GetEnumerator()
        {
            if (!IsLoaded)
                LoadFromDatabase();
            return base.GetEnumerator();
        }

        public override void LoadFromDatabase()
        {
            Clear();
            XDocument dataDocument = XDocument.Parse(ReadXmlString(Name + ".xml"));
            foreach (var item in (dataDocument.FirstNode as XElement).Elements("Item"))
                Add(new RootItem(item));
            IsLoaded = true;
        }
        
        protected override void OnGetCount()
        {
            if (!IsLoaded)
                LoadFromDatabase();
            base.OnGetCount();
        }

        public override string ToString()
        {
            return String.Format("{0} Items", Count);
        }

        #endregion
    }
}
