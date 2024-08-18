using System.Diagnostics;
using System.Management.Automation;
using System.Management.Automation.Runspaces;

namespace RunPowerShellCSharp;

internal static class Extensions {
    internal static Task InvokePowerShellAsyncWithInput(
        this PowerShell ps,
        PSDataCollection<PSObject> input,
        PSDataCollection<PSObject> output) =>
        Task.Factory.FromAsync(ps.BeginInvoke(input, output), ps.EndInvoke);
}

internal static class Program {
    static async Task Main(string[] args) {
        var parameters = new Dictionary<string, object>
        {
            { "Number", 1 }
        };

        var executor = new PowerShellExecutor();
        await executor.ExecuteScriptAsync(parameters);
    }
}


internal class PowerShellExecutor {

    public async Task ExecuteScriptAsync(Dictionary<string, object> parameters) {
        // We will use the preferred method if it is set, otherwise we will use the default method
        var method = PowerShellMethod.NamedPipe;

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
    private static Task InvokePowerShellAsync(PowerShell powerShell, PSDataCollection<PSObject> output) => Task.Factory.FromAsync(powerShell.BeginInvoke<PSObject, PSObject>(null, output), powerShell.EndInvoke);


    private async Task ExecuteCommonScriptAsync(PowerShell ps, Dictionary<string, object> parameters) {
        var psParameters = parameters;
        Environment.SetEnvironmentVariable("ADPS_LoadDefaultDrive", "0");

        try {

            string[] specialModules = { "ActiveDirectory", "ADEssentials" };
            foreach (var module in specialModules) {
                ps.AddStatement().AddCommand("Import-Module").AddParameter("Name", module);
            }

            ScriptBlock scriptBlock = AddParametersToScriptBlock(ScriptBlock.Create("get-disk"), parameters);

            ps.AddScript(scriptBlock.ToString());
            ps.AddParameters(psParameters);

            PSDataCollection<PSObject> results = new PSDataCollection<PSObject>();
            await InvokePowerShellAsync(ps, results);

            foreach (var information in ps.Streams.Information) {
                Console.WriteLine(information.MessageData.ToString());
            }

            foreach (var warning in ps.Streams.Warning) {
                Console.WriteLine("Warning: " + warning.Message);
            }

            foreach (var error in ps.Streams.Error) {
                Console.WriteLine($"PowerShell Error: {error.Exception.Message}");
            }

            if (results.Count > 0) {
                Console.WriteLine("Results count: " + results.Count);

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

        // Convert the ScriptBlock to a string
        string whereObjectString = whereObject.ToString();

        // Add the script to filter results
        //ps.AddScript("$Results | Where-Object {" + whereObjectString + "}");

        //// Pass the results as input to the script
        //ps.AddParameter("Results", results);

        //ps.AddCommand("Where-Object").AddParameter("FilterScript", whereObject);
        //ps.AddCommand("Where-Object").AddArgument(whereObject);


        ps.AddScript("$args[0] | Where-Object {" + whereObjectString + "}", true).AddArgument(results);
        //ps.AddScript("param($testing) $testing | Where-Object {" + whereObjectString + "}", useLocalScope: true);
        //ps.AddParameter("testing", results);
        //ps.AddScript("Write-Host $Results.Count");
        //ps.AddParameter("Results", results);

        var evaluationResult1 = ps.Invoke();

        Console.WriteLine("Evaluation Count: " + evaluationResult1.Count);

        // Invoke the script asynchronously
        //ps.InvokePowerShellAsyncWithInput(results, evaluationResult);

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


public enum PowerShellMethod {
    NamedPipe,
    InProcess,
    OutOfProcess
}