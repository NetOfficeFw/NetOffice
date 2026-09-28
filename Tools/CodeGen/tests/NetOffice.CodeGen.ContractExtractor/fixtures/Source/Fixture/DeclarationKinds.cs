namespace Fixture
{
    /// <summary>Declaration fixture.</summary>
    public class DeclarationKinds : System.Collections.IEnumerable, IFoo
    {
        /// <summary>A multiline property.</summary>
        public
        string Name
        {
            get
            {
                return "name";
            }
            set
            {
            }
        }

        public int this[
            string key = "default"
        ]
        {
            get { return 1; }
            set { }
        }

        public event System.EventHandler Changed;
        public const int Field = 7;

        public string Call(
            int count = 3,
            string text = "x")
        {
            return text;
        }

        System.Collections.IEnumerator System.Collections.IEnumerable.GetEnumerator()
        {
            return null;
        }

        object IFoo.Value
        {
            get { return null; }
        }
    }

    public interface IFoo
    {
        object Value { get; }
    }
}
