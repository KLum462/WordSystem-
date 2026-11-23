namespace WordClient.Core.Operations
{
    public sealed class UppercaseOperation : IOperation
    {
        public string Name() => "uppercase";

        public string Process(string text) => text.ToUpperInvariant();
    }
}
