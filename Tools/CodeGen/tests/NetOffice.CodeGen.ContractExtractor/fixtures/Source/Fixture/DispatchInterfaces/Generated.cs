namespace Fixture
{
    /// <summary>Generated wrapper.</summary>
    [SupportByVersion("Fixture", 1, 2)]
    public class Generated : BaseGenerated, IFoo
    {
        [System.ComponentModel.EditorBrowsable(System.ComponentModel.EditorBrowsableState.Never)]
        public Generated()
        {
        }

        public event System.EventHandler Changed;

        /// <summary>Ignored private member.</summary>
        /// <param name="hidden">Must not leak.</param>
        private void Hidden(int hidden)
        {
        }

        /// <summary>Invokes the generated member.</summary>
        /// <param name="value">The value.</param>
        [SupportByVersion("Fixture", 2)]
        public string Invoke(int value = 4)
        {
            for(;;) Factory.Trace();
            if(false==Factory.Settings.Enabled) return null;
            return Factory.ExecuteStringMethodGet(this, "Invoke", new object[]{ value });;
        }
    }

    public delegate void GeneratedEventHandler(int value);

    public enum GeneratedEnum
    {
        Value = 128
    }
}
