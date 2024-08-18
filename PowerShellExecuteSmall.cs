using System.Diagnostics;
using System.Management.Automation;
using System.Management.Automation.Language;
using System.Management.Automation.Runspaces;

namespace RunPowerShellCSharp;

internal class PowerShellExecuteSmall {
    public async Task ExecuteScriptAsync(Dictionary<string, object> parameters) {
        // We will use the preferred method if it is set, otherwise we will use the default method
        var method = PowerShellMethod.NamedPipe;

        switch (method) {
            case PowerShellMethod.NamedPipe:
                await ExecuteScriptNamedPipeAsync(parameters);
                break;
            default:
                throw new ArgumentException("Invalid method specified.");
        }
    }

    private async Task ExecuteScriptNamedPipeAsync(Dictionary<string, object> parameters) {
        var startInfo = new ProcessStartInfo {
            FileName = "powershell.exe",
            Arguments = "-WindowStyle Hidden -NoProfile -NonInteractive -ExecutionPolicy Bypass",
            UseShellExecute = false,
            CreateNoWindow = true
        };

        using var process = Process.Start(startInfo);

        try {
            var pipe = new NamedPipeConnectionInfo(process!.Id);
            using var runspace = RunspaceFactory.CreateRunspace(pipe);
            runspace.Open();
            using PowerShell ps = PowerShell.Create();
            ps.Runspace = runspace;
            await ExecuteCommonScriptAsync(ps, parameters);
        } catch (Exception) {
            throw;
        } finally {
            process?.Kill();
        }
    }

    private async Task ExecuteCommonScriptAsync(PowerShell ps, Dictionary<string, object> parameters) {
        PSDataCollection<PSObject> results = new PSDataCollection<PSObject>();
        await ps
            .AddCommand("Get-Process")
            .AddCommand("Select-Object")
            .AddArgument(new string[] { "Name", "Id" })
            .InvokePowerShellAsync(results);

        if (results.Count > 0) {
            var WhereObject = ScriptBlock.Create("$_.Name -eq 'pwsh'");
            await FilterResultsInSession(ps, results, WhereObject);
        }
    }

    private async Task<PSDataCollection<PSObject>> FilterResultsInSession(
        PowerShell ps,
        PSDataCollection<PSObject> results,
        ScriptBlock whereObject) {
        PSDataCollection<PSObject> evaluationResult = new PSDataCollection<PSObject>();
        ps.Commands.Clear();
        ps.Streams.ClearStreams();
        results.Complete();

        Console.WriteLine("Filter used: " + whereObject);
        await ps
            // .AddParameter("FilterScript", ((ScriptBlockAst)whereObject.Ast).GetScriptBlock())
            .AddScript("$input | Where-Object ([ScriptBlock]::Create($args[0]))")
            .AddArgument(whereObject)
            .InvokePowerShellAsyncWithInput(results, evaluationResult);

        evaluationResult.Complete();
        Console.WriteLine("ResultsCount: " + results.Count);
        Console.WriteLine("EvaluationResultCount: " + evaluationResult.Count);

        return evaluationResult;
    }
}