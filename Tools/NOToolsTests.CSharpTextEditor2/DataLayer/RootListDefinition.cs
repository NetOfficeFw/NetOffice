using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace NOToolsTests.CSharpTextEditor2.DataLayer
{
    public class RootListDefinition : IEnumerable<AccessContext>
    {
        public RootListDefinition(string tableName)
        {
            Name = tableName;
            List = new List<AccessContext>();
            DataSource = new RootList(tableName);
        }
        
        public string Name { get; private set; }

        internal RootList DataSource { get; private set; }

        private List<AccessContext> List { get; set; }

        public AccessContext Add(string uniqueAccessContextName, bool throwExceptionIfAlreadyExists = false)
        {
            foreach (var item in List)
            {
                if (item.Name.Equals(uniqueAccessContextName, StringComparison.InvariantCultureIgnoreCase))
                {
                    if (throwExceptionIfAlreadyExists)
                        throw new InvalidOperationException("Context name already exists");
                    else
                        return item;
                }
            }

            AccessContext newContext = new AccessContext(this, uniqueAccessContextName);
            List.Add(newContext);

            return newContext;
        }

        public void Remove(string uniqueAccessContextName)
        {
            AccessContext delContext = null;
            foreach (var item in List)
            {
                if (item.Name.Equals(uniqueAccessContextName, StringComparison.InvariantCultureIgnoreCase))
                {
                    delContext = item;
                    break;
                }
            }

            if (null != delContext)
                List.Remove(delContext);
        }

        public AccessContext this[string uniqueAccessContextName]
        {
            get
            {
                foreach (var item in List)
                {
                    if (item.Name.Equals(uniqueAccessContextName, StringComparison.InvariantCultureIgnoreCase))
                            return item;
                }
                throw new ArgumentOutOfRangeException(uniqueAccessContextName);
            }
        }

        public IEnumerator<AccessContext> GetEnumerator()
        {
            return List.GetEnumerator();
        }

        System.Collections.IEnumerator System.Collections.IEnumerable.GetEnumerator()
        {
            return List.GetEnumerator();
        }
    }
}
