using Xunit;

namespace Delegates.Tests
{
    // This collection ensures that all tests manipulating Console.Out run serially
    [CollectionDefinition("Console Tests", DisableParallelization = true)]
    public class ConsoleTestCollection
    {
    }
}
