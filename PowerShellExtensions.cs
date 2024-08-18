using System.Management.Automation;

namespace RunPowerShellCSharp;

/// <summary>
/// Extension methods for PowerShell
/// </summary>
internal static class PowerShellExtensions {
    /// <summary>
    /// Invoke PowerShell asynchronously with input and output
    /// </summary>
    /// <param name="ps"></param>
    /// <param name="input"></param>
    /// <param name="output"></param>
    /// <returns></returns>
    internal static Task InvokePowerShellAsyncWithInput(this PowerShell ps, PSDataCollection<PSObject> input, PSDataCollection<PSObject> output) => Task.Factory.FromAsync(ps.BeginInvoke(input, output), ps.EndInvoke);

    /// <summary>
    /// Invoke PowerShell asynchronously with output
    /// </summary>
    /// <param name="powerShell"></param>
    /// <param name="output"></param>
    /// <returns></returns>
    internal static Task InvokePowerShellAsync(this PowerShell powerShell, PSDataCollection<PSObject> output) => Task.Factory.FromAsync(powerShell.BeginInvoke<PSObject, PSObject>(null, output), powerShell.EndInvoke);
}