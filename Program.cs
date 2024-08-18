namespace RunPowerShellCSharp;

internal static class Program {
    /// <summary>
    /// Small example showing only single PowerShell method made by Santi
    /// </summary>
    /// <returns></returns>
    private static async Task ExampleSmall() {
        var parameters = new Dictionary<string, object>();
        var executor = new PowerShellExecuteSmall();
        await executor.ExecuteScriptAsync(parameters);
    }

    /// <summary>
    /// A bit more advanced example showing multiple PowerShell methods as proposed by Santi
    /// </summary>
    /// <returns></returns>
    private static async Task ExampleBig() {
        var parameters = new Dictionary<string, object>
        {
            { "Number", 1 }
        };
        var executor = new PowerShellExecuteBig();
        await executor.ExecuteScriptAsync(parameters);
    }

    /// <summary>
    /// Main method to run the examples
    /// </summary>
    /// <param name="args"></param>
    /// <returns></returns>
    static async Task Main(string[] args) {
        await ExampleSmall();
        Console.WriteLine("-------");
        await ExampleBig();
        Console.WriteLine("-------");
    }
}