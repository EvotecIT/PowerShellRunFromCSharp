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
        Console.WriteLine("[i] Executing common script Async");
        var psParameters = parameters;
        PSDataCollection<PSObject> results = new PSDataCollection<PSObject>();
        try {

            string[] specialModules = { "ActiveDirectory", "Microsoft.PowerShell.Security" };
            foreach (var module in specialModules) {
                //ps.AddCommand("Import-Module").AddParameter("Name", module).AddStatement();
            }
            ScriptBlock scriptBlock = AddParametersToScriptBlock(ScriptBlock.Create("get-disk"), parameters);
            await ps.AddScript(scriptBlock.ToString()).AddParameters(psParameters).InvokePowerShellAsync(results);

            //await ps
            //    .AddCommand("Import-Module")
            //    .AddArgument(new string[] { "ActiveDirectory", "Microsoft.PowerShell.Security" })
            //    .AddStatement()
            //    .AddCommand("Get-Disk")
            //    .InvokePowerShellAsync(results);

            //await ps.AddCommand("Get-Disk").InvokePowerShellAsync(results);

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
            results.Complete();
            foreach (var result in results) {
                var friendlyName = result.Properties["FriendlyName"]?.Value;
                var serialNumber = result.Properties["SerialNumber"]?.Value;
                var totalSize = result.Properties["Total Size"]?.Value;
                Console.WriteLine($"> Friendly Name: {friendlyName}, Serial Number: {serialNumber}, Total Size: {totalSize}");
            }

            if (results.Count > 0) {
                var WhereObject = ScriptBlock.Create("$_.Number -eq 1");

                var test = await FilterResultsInSession(ps, results, WhereObject);
                foreach (var item in test) {
                    var friendlyName = item.Properties["FriendlyName"]?.Value;
                    var serialNumber = item.Properties["SerialNumber"]?.Value;
                    var totalSize = item.Properties["Total Size"]?.Value;
                    Console.WriteLine($"> Friendly Name: {friendlyName}, Serial Number: {serialNumber}, Total Size: {totalSize}");
                }

                var WhereObject1 = ScriptBlock.Create("$_.Number -eq 2");
                var test1 = await FilterResultsInSession(ps, results, WhereObject1);
                foreach (var item in test1) {
                    var friendlyName = item.Properties["FriendlyName"]?.Value;
                    var serialNumber = item.Properties["SerialNumber"]?.Value;
                    var totalSize = item.Properties["Total Size"]?.Value;
                    Console.WriteLine($"> Friendly Name: {friendlyName}, Serial Number: {serialNumber}, Total Size: {totalSize}");
                }
            }
        } catch (RuntimeException rex) {
            Console.WriteLine($"Exception in ExecuteCommonScriptAsync for rule: {rex.Message}");
        } catch (Exception ex) {
            Console.WriteLine($"Exception in ExecuteCommonScriptAsync for rule: {ex.Message}");
        } finally {

        }
    }

    private async Task<PSDataCollection<PSObject>> FilterResultsInSession(PowerShell ps, PSDataCollection<PSObject> results, ScriptBlock whereObject) {
        Console.WriteLine("[i] Applying filter to results");
        PSDataCollection<PSObject> evaluationResult = new PSDataCollection<PSObject>();
        ps.Commands.Clear();
        // Clear previous streams
        ps.Streams.ClearStreams();

        Console.WriteLine("Filter used: " + whereObject);
        //ps.AddScript("$input | Where-Object ([ScriptBlock]::Create($args[0]))").AddArgument(whereObject);

        await ps.AddScript("$input | Where-Object ([ScriptBlock]::Create($args[0]))")
            .AddArgument(whereObject)
            .InvokePowerShellAsyncWithInput(results, evaluationResult);

        //await ps.InvokePowerShellAsyncWithInput(results, evaluationResult);

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

        Console.WriteLine("ResultsCount (still available): " + results.Count);
        Console.WriteLine("EvaluationResultCount: " + evaluationResult.Count);

        evaluationResult.Complete();

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

