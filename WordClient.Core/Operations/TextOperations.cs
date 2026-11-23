namespace WordClient.Core.Operations
{
    public static class TextOperations
    {
        public static readonly IOperation[] SupportedOperations = [
            new UppercaseOperation(),
            new LowercaseOperation ()
        ];

        public static IEnumerable<string> OperationsNames() => SupportedOperations.Select(d => d.Name());
    }
}
