namespace WordClient.Core.Operations
{
    public sealed class LowercaseOperation : IOperation
    {
        public string Name() => "lowercase";

        public string Process(string text) => text.ToLowerInvariant();
    }
}
