using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Linq;
using System.Management.Automation.Runspaces;
using System.Management.Automation;

namespace RunPowerShellCSharp;


internal class PowerShellExecuteBig {

    public async Task ExecuteScriptAsync(Dictionary<string, object> parameters) {
        Environment.SetEnvironmentVariable("ADPS_LoadDefaultDrive", "0");

        // We will use the preferred method if it is set, otherwise we will use the default method
        PowerShellMethod method = PowerShellMethod.NamedPipe;

        switch (method) {
            case PowerShellMethod.InProcess:
                await ExecuteScriptInProcessAsync(parameters);
                break;
            case PowerShellMethod.OutOfProcess:
                await ExecuteScriptOutOfProcessAsync(parameters);
                break;
            case PowerShellMethod.NamedPipe:
                await ExecuteScriptNamedPipeAsync(parameters);
                break;
            default:
                throw new ArgumentException("Invalid method specified.");
        }
    }

    private async Task ExecuteScriptInProcessAsync(Dictionary<string, object> parameters) {
        using (PowerShell ps = PowerShell.Create(RunspaceMode.NewRunspace)) {
            await ExecuteCommonScriptAsync(ps, parameters);
        }
    }

    private async Task ExecuteScriptOutOfProcessAsync(Dictionary<string, object> parameters) {
        using (Runspace runspace = RunspaceFactory.CreateOutOfProcessRunspace(TypeTable.LoadDefaultTypeFiles())) {
            runspace.Open();
            using (PowerShell ps = PowerShell.Create()) {
                ps.Runspace = runspace;
                await ExecuteCommonScriptAsync(ps, parameters);
            }
        }
    }

    private async Task ExecuteScriptNamedPipeAsync(Dictionary<string, object> parameters) {
        var startInfo = new ProcessStartInfo {
            FileName = "powershell.exe",
            Arguments = "-WindowStyle Hidden -NoProfile -NonInteractive -ExecutionPolicy Bypass",
            UseShellExecute = false,
            RedirectStandardInput = true,
            RedirectStandardOutput = true,
            RedirectStandardError = true,
            CreateNoWindow = true
        };

        using var process = Process.Start(startInfo);
        if (process == null) {
            throw new InvalidOperationException("Failed to start PowerShell process.");
        }

        try {
            var pipe = new NamedPipeConnectionInfo(process.Id);
            using (var runspace = RunspaceFactory.CreateRunspace(pipe)) {
                runspace.Open();
                using (PowerShell ps = PowerShell.Create()) {
                    ps.Runspace = runspace;
                    await ExecuteCommonScriptAsync(ps, parameters);
                }
            }
        } catch (Exception ex) {
            throw;
        } finally {
            process?.Kill();
        }
    }
    private async Task ExecuteCommonScriptAsync(PowerShell ps, Dictionary<string, object> parameters) {
        var psParameters = parameters;
        PSDataCollection<PSObject> results = new PSDataCollection<PSObject>();
        try {

            string[] specialModules = { "Microsoft.PowerShell.Security" };
            foreach (var module in specialModules) {
                //ps.AddStatement().AddCommand("Import-Module").AddParameter("Name", module);
            }

            //ScriptBlock scriptBlock = AddParametersToScriptBlock(ScriptBlock.Create("get-disk"), parameters);

            //await ps.AddScript(scriptBlock.ToString()).AddParameters(psParameters).InvokePowerShellAsync(results);

            //await ps.AddStatement().AddCommand("Import-Module").AddParameter("Name", "ActiveDirectory").AddCommand("Get-Disk").InvokePowerShellAsync(results);

            await ps.AddCommand("Get-Disk").InvokePowerShellAsync(results);

            results.Complete();

            foreach (var information in ps.Streams.Information) {
                Console.WriteLine(information.MessageData.ToString());
            }

            foreach (var warning in ps.Streams.Warning) {
                Console.WriteLine("Warning: " + warning.Message);
            }

            foreach (var error in ps.Streams.Error) {
                Console.WriteLine($"Error: {error.Exception.Message}");
            }

            Console.WriteLine("Results count: " + results.Count);
            if (results.Count > 0) {
                var WhereObject = ScriptBlock.Create("$_.Number -eq 1");

                var test = await FilterResultsInSession(ps, results, WhereObject);
            }
        } catch (RuntimeException rex) {
            Console.WriteLine($"Exception in ExecuteCommonScriptAsync for rule: {rex.Message}");
        } catch (Exception ex) {
            Console.WriteLine($"Exception in ExecuteCommonScriptAsync for rule: {ex.Message}");
        } finally {

        }
    }

    private async Task<PSDataCollection<PSObject>> FilterResultsInSession(PowerShell ps, PSDataCollection<PSObject> results, ScriptBlock whereObject) {
        PSDataCollection<PSObject> evaluationResult = new PSDataCollection<PSObject>();
        ps.Commands.Clear();
        // Clear previous streams
        ps.Streams.ClearStreams();

        Console.WriteLine("Filter used: " + whereObject);
        ps.AddScript("$input | Where-Object ([ScriptBlock]::Create($args[0]))").AddArgument(whereObject);

        await ps.AddScript("$input | Where-Object ([ScriptBlock]::Create($args[0]))")
            .AddArgument(whereObject)
            .InvokePowerShellAsyncWithInput(results, evaluationResult);

        evaluationResult.Complete();

        foreach (var information in ps.Streams.Information) {
            Console.WriteLine("Information: " + information.MessageData.ToString());
        }

        foreach (var warning in ps.Streams.Warning) {
            Console.WriteLine("Warning: " + warning.Message);
        }
        // Check for errors
        if (ps.HadErrors) {
            var message = ps.Streams.Error[0];
            throw new Exception($"Evaluating WhereObject failed with message: {message}");
        }

        Console.WriteLine("ResultsCount: " + results.Count);
        Console.WriteLine("EvaluationResultCount: " + evaluationResult.Count);

        return evaluationResult;
    }

    private ScriptBlock AddParametersToScriptBlock(ScriptBlock scriptBlock, Dictionary<string, object> parameters) {
        if (parameters == null || parameters.Count == 0) {
            return scriptBlock;
        }

        var paramDefinitions = string.Join(", ", parameters.Keys.Select(key => $"${key}"));
        var paramBlock = $"param({paramDefinitions})";
        var scriptWithParams = $"{paramBlock}\n{scriptBlock}";

        return ScriptBlock.Create(scriptWithParams);
    }
}

