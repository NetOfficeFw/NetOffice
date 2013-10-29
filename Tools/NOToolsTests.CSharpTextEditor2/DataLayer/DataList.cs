using System;
using System.IO;
using System.Reflection;
using System.ComponentModel;
using System.Collections.Generic;
using System.Linq;
using System.Xml;
using System.Xml.Linq;
using System.Text;

namespace NOToolsTests.CSharpTextEditor2.DataLayer
{
    public abstract class DataList<T>: List<T> , ITypedList where T: DataItem
    {
        public DataList(string listName)
        {
            ListName = listName;
            Schema = XDocument.Parse(ReadXmlString(listName + ".xsd"));
        }

        private XDocument Schema { get; set; }

        internal string ListName { get; set; }

        public abstract void LoadFromDatabase();

        /// <summary>
        /// 
        /// </summary>
        /// <param name="listAccessors"></param>
        /// <returns></returns>
        public PropertyDescriptorCollection GetItemProperties(PropertyDescriptor[] listAccessors)
        {
            List<PropertyDescriptor> list = new List<PropertyDescriptor>();

            XElement rootNode = ((Schema.FirstNode as XElement).FirstNode as XElement);
            foreach (var item in rootNode.Element("sequence").Elements("element"))
                list.Add(new RootPropertyDescriptor(item));

            return new PropertyDescriptorCollection(list.ToArray());
        }

        public string GetListName(PropertyDescriptor[] listAccessors)
        {
            return ListName;
        }

        public override string ToString()
        {
            return "DataList " + ListName;
        }

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
    }
}
