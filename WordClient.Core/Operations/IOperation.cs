namespace WordClient.Core.Operations
{
    public interface IOperation
    {
        string Process(string text);
        string Name();
    }
}
