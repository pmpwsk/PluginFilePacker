using EnvDTE;
using EnvDTE80;
using Microsoft.VisualStudio.Shell;
using Microsoft.VisualStudio.Shell.Interop;
using System;
using System.Collections.Generic;
using System.ComponentModel.Design;
using System.IO;
using System.Linq;
using System.Resources;
using System.Text;
using VSLangProj;
using Task = System.Threading.Tasks.Task;

namespace uwap.VSIX.PluginFilePacker
{
    /// <summary>
    /// Command handler
    /// </summary>
    internal sealed class GenerateFileHandler
    {
        #region Stuff
        /// <summary>
        /// Command ID.
        /// </summary>
        public const int CommandId = 0x0100;

        /// <summary>
        /// Command menu group (command set GUID).
        /// </summary>
        public static readonly Guid CommandSet = new Guid("335da2a2-e8fe-4d34-98aa-6876d6248cfc");

        /// <summary>
        /// VS Package that provides this command, not null.
        /// </summary>
        private readonly AsyncPackage package;

        /// <summary>
        /// Initializes a new instance of the <see cref="GenerateFileHandler"/> class.
        /// Adds our command handlers for menu (commands must exist in the command table file)
        /// </summary>
        /// <param name="package">Owner package, not null.</param>
        /// <param name="commandService">Command service to add command to, not null.</param>
        private GenerateFileHandler(AsyncPackage package, OleMenuCommandService commandService)
        {
            this.package = package ?? throw new ArgumentNullException(nameof(package));
            commandService = commandService ?? throw new ArgumentNullException(nameof(commandService));

            var menuCommandID = new CommandID(CommandSet, CommandId);
            var menuItem = new MenuCommand(this.Execute, menuCommandID);
            commandService.AddCommand(menuItem);
        }

        /// <summary>
        /// Gets the instance of the command.
        /// </summary>
        public static GenerateFileHandler Instance
        {
            get;
            private set;
        }

        /// <summary>
        /// Gets the service provider from the owner package.
        /// </summary>
        private Microsoft.VisualStudio.Shell.IAsyncServiceProvider ServiceProvider
        {
            get
            {
                return this.package;
            }
        }

        /// <summary>
        /// Initializes the singleton instance of the command.
        /// </summary>
        /// <param name="package">Owner package, not null.</param>
        public static async Task InitializeAsync(AsyncPackage package)
        {
            // Switch to the main thread - the call to AddCommand in GenerateFileHandler's constructor requires
            // the UI thread.
            await ThreadHelper.JoinableTaskFactory.SwitchToMainThreadAsync(package.DisposalToken);

            OleMenuCommandService commandService = await package.GetServiceAsync(typeof(IMenuCommandService)) as OleMenuCommandService;
            Instance = new GenerateFileHandler(package, commandService);
        }

        /// <summary>
        /// This function is the callback used to execute the command when the menu item is clicked.
        /// See the constructor to see how the menu item is associated with this function using
        /// OleMenuCommandService service and MenuCommand class.
        /// </summary>
        /// <param name="sender">Event sender.</param>
        /// <param name="e">Event args.</param>
        private void Execute(object sender, EventArgs e)
        {
            ThreadHelper.ThrowIfNotOnUIThread();
            _ = ThreadHelper.JoinableTaskFactory.RunAsync(async delegate { await ExecuteAsync(); });
        }
        #endregion

        private async Task ExecuteAsync()
        {
            try
            {
                await ThreadHelper.JoinableTaskFactory.SwitchToMainThreadAsync();

                IVsStatusbar statusbar = (IVsStatusbar)await ServiceProvider.GetServiceAsync(typeof(SVsStatusbar));
                DTE dte = (DTE)await ServiceProvider.GetServiceAsync(typeof(DTE));
                
                Array projects = (Array)dte.ActiveSolutionProjects;
                if (projects == null || projects.Length == 0)
                {
                    MessageBox("Error!", "Select a project first!");
                    await StatusbarTextAsync("Select a project first!", statusbar);
                }
                else if (projects.Length > 1)
                {
                    MessageBox("Error!", "Please only select one project at a time!");
                    await StatusbarTextAsync("Please only select one project at a time!", statusbar);
                }
                else
                {
                    Project project = (Project)projects.GetValue(0);
                    try
                    {
                        Options options = ((PluginFilePackerPackage)this.package).Options;

                        await StatusbarTextAsync($"Generating FileHandler.cs for {project.Name}...", statusbar);
                        string projPath = new FileInfo(project.FullName).DirectoryName;

                        await GenerateAsync(project.Name, projPath, options, project, dte);

                        await Task.Delay(500);
                        await StatusbarTextAsync($"Successfully generated FileHandler.cs for {project.Name}!", statusbar);
                        if (options.ShowPopupWhenDone)
                            MessageBox("Done!", $"Successfully generated FileHandler.cs for {project.Name}!");
                    }
                    catch (Exception ex)
                    {
                        MessageBox("Error!", $"An error occurred while generating FileHandler.cs for {project.Name}:\n{ex.Message}");
                        await StatusbarTextAsync($"An error occurred while generating FileHandler.cs for {project.Name}!", statusbar);
                    }
                }

                await Task.Delay(5000);
                statusbar.FreezeOutput(0);
                statusbar.Clear();
            }
            catch (Exception ex)
            {
                MessageBox("Error!", $"An error stopped the operation:\n{ex.Message}");
            }
        }

        private async Task GenerateAsync(string projName, string projPath, Options options, Project project, DTE dte)
        {
            await ThreadHelper.JoinableTaskFactory.SwitchToMainThreadAsync();
            string filesPath = projPath + "/Files";
            if (!Directory.Exists(filesPath))
                throw new Exception("Directory 'Files' not found!");

            Dictionary<string,string> resourceFiles = options.UseBase64 ? null : new Dictionary<string, string>();

            string namespace_ = DetectNamespaceFromFile(projPath + "/FileHandler.cs") ?? DetectNamespaceFromFile(projPath + "/FileHandlerCustom.cs") ?? options.DefaultNamespace;
            string class_ = DetectClassFromFile(projPath + "/FileHandler.cs") ?? DetectClassFromFile(projPath + "/FileHandlerCustom.cs") ?? projName;
            using (StreamWriter writer = new StreamWriter($"{projPath}/FileHandler.cs", false, Encoding.UTF8))
            {
                try
                {
                    Dictionary<string, string> timestamps = new Dictionary<string, string>();
                    IEnumerable<string> textTypes = options.TextExtensionsClean;

                    string wfNamespacePrefix = (namespace_ == "uwap.WebFramework" || namespace_.StartsWith("uwap.WebFramework.")) ? "" : "uwap.WebFramework.";

                    if (!options.UseBase64)
                    {
                        await writer.WriteLineAsync($"using {projName}.Properties;");
                        await writer.WriteLineAsync();
                    }
                    await writer.WriteLineAsync($"namespace {namespace_};");
                    await writer.WriteLineAsync();
                    await writer.WriteLineAsync($"public partial class {class_} : {(namespace_ != "uwap.WebFramework.Plugins" ? "uwap.WebFramework.Plugins." : "")}Plugin");
                    await writer.WriteLineAsync("{");
                    await writer.WriteLineAsync("    public override byte[]? GetFile(string relPath, string pathPrefix, string domain)");
                    await writer.WriteLineAsync("        => relPath switch");
                    await writer.WriteLineAsync("        {");
                    int resxIndex = 0;
                    foreach (string file in Directory.GetFiles(filesPath, "*", SearchOption.AllDirectories))
                    {
                        string relPath = file.Remove(0, filesPath.Length).Replace('\\', '/');
                        timestamps.Add(relPath, File.GetLastWriteTimeUtc(file).Ticks.ToString());
                        //check content as text file if it's a text file type
                        if (textTypes.Any(relPath.EndsWith))
                        {
                            //text file
                            string content = File.ReadAllText(file);
                            if (content.Contains("[PATH_PREFIX]") || content.Contains("[PATH_HOME]") || content.Contains("[DOMAIN]"))
                            {
                                content = Microsoft.CodeAnalysis.CSharp.SymbolDisplay.FormatLiteral(content, true)
                                    .Replace("{", "{{").Replace("}", "}}")
                                    .Replace("[PATH_PREFIX]", "{pathPrefix}")
                                    .Replace("[PATH_HOME]", "{(pathPrefix == \"\" ? \"/\" : pathPrefix)}")
                                    .Replace("[DOMAIN]", "{domain}")
                                    .Replace("[DOMAIN_MAIN]", $"{{{wfNamespacePrefix}Parsers.DomainMain(domain)}}");
                                await writer.WriteLineAsync($"            \"{relPath}\" => System.Text.Encoding.UTF8.GetBytes(${content}),");
                                continue;
                            }
                        }
                        if (options.UseBase64)
                            await writer.WriteLineAsync($"            \"{relPath}\" => Convert.FromBase64String(\"{Convert.ToBase64String(File.ReadAllBytes(file))}\"),");
                        else
                        {
                            resourceFiles[$"File{resxIndex}"] = file;
                            await writer.WriteLineAsync($"            \"{relPath}\" => PluginFiles.File{resxIndex},");
                            resxIndex++;
                        }
                    }
                    if (File.Exists(projPath + "/FileHandlerCustom.cs"))
                        await writer.WriteLineAsync("            _ => GetFileCustom(relPath, pathPrefix, domain)");
                    else await writer.WriteLineAsync("            _ => null");
                    await writer.WriteLineAsync("        };");
                    await writer.WriteLineAsync("    ");
                    await writer.WriteLineAsync("    public override string? GetFileVersion(string relPath)");
                    await writer.WriteLineAsync("        => relPath switch");
                    await writer.WriteLineAsync("        {");
                    foreach (var kv in timestamps)
                        await writer.WriteLineAsync($"            \"{kv.Key}\" => \"{kv.Value}\",");
                    if (File.Exists(projPath + "/FileHandlerCustom.cs"))
                        await writer.WriteLineAsync("            _ => GetFileVersionCustom(relPath)");
                    else await writer.WriteLineAsync("            _ => null");
                    await writer.WriteLineAsync("        };");
                    await writer.WriteLineAsync("}");

                    await writer.FlushAsync();
                }
                finally
                {
                    writer.Close();
                    writer.Dispose();
                }
            }

            string csprojStart = "<!--PluginFilePacker start-->";
            string csprojMain = "<ItemGroup><Compile Update=\"Properties\\PluginFiles.Designer.cs\"><DesignTime>True</DesignTime><AutoGen>True</AutoGen><DependentUpon>PluginFiles.resx</DependentUpon></Compile></ItemGroup><ItemGroup><EmbeddedResource Update=\"Properties\\PluginFiles.resx\"><Generator>ResXFileCodeGenerator</Generator><LastGenOutput>PluginFiles.Designer.cs</LastGenOutput></EmbeddedResource></ItemGroup>";
            string csprojEnd = "<!--PluginFilePacker end-->";
            string csproj = File.ReadAllText($"{projPath}/{projName}.csproj");
            if (resourceFiles != null && resourceFiles.Count > 0)
            {
                if (!csproj.Contains(csprojStart + csprojMain + csprojEnd))
                {
                    if (SplitAtFirst(csproj, csprojStart, out string part1, out string part2))
                    {
                        if (SplitAtLast(part2, csprojEnd, out _, out string part2_2))
                            csproj = part1 + csprojStart + csprojMain + csprojEnd + part2_2;
                        else csproj = part1 + csprojStart + csprojMain + csprojEnd + part2;
                    }
                    else if (SplitAtLast(csproj, "</Project>", out string beforeEnd, out string afterEnd))
                    {
                        if (SplitAtLast(beforeEnd, "\n", out string beforeEnd_1, out string beforeEnd_2))
                            csproj = beforeEnd_1 + "\n" + beforeEnd_2 + "  " + csprojStart + csprojMain + csprojEnd + "\n" + beforeEnd_2 + "</Project>" + afterEnd;
                        else csproj = beforeEnd + csprojStart + csprojMain + csprojEnd + "</Project>" + afterEnd;
                    }
                    else throw new Exception("The RESX file couldn't be attached to the project!");
                    File.WriteAllText($"{projPath}/{projName}.csproj", csproj);
                    ReloadProject(projName, dte);
                }

                Directory.CreateDirectory(projPath + "/Properties");
                using (var writer = new ResXResourceWriter(projPath + "/Properties/PluginFiles.resx"))
                {
                    foreach (var kv in resourceFiles)
                        writer.AddResource(kv.Key, File.ReadAllBytes(kv.Value));
                }

                VSProjectItem resx = null;
                for (int i = 0; i < 10; i++)
                {
                    resx = dte.Solution.FindProjectItem(projPath + "/Properties/PluginFiles.resx")?.Object as VSProjectItem;
                    if (resx != null)
                        break;
                    else await Task.Delay(1000);
                }
                if (resx == null)
                    throw new Exception("The RESX file couldn't be found as a project item! This can usually be fixed by trying again - otherwise, try restarting Visual Studio.");
                else resx.RunCustomTool();
            }
            else
            {
                if (File.Exists(projPath + "/Properties/PluginFiles.resx"))
                    File.Delete(projPath + "/Properties/PluginFiles.resx");
                if (File.Exists(projPath + "/Properties/PluginFiles.Designer.cs"))
                    File.Delete(projPath + "/Properties/PluginFiles.Designer.cs");
                if (SplitAtFirst(csproj, csprojStart, out string part1, out string part2) && SplitAtLast(part2, csprojEnd, out _, out string part2_2))
                {
                    csproj = part1 + part2_2;
                    ReloadProject(projName, dte);
                }
            }
        }

        private string DetectNamespaceFromFile(string path)
        {
            if (!File.Exists(path))
                return null;
            foreach (string line in File.ReadAllLines(path))
                if (line.StartsWith("namespace "))
                    return line.Remove(0, 10).Split(' ', '{', ';').First();
            return null;
        }

        private string DetectClassFromFile(string path)
        {
            if (!File.Exists(path))
                return null;
            foreach (string line in File.ReadAllLines(path))
            {
                int index = line.IndexOf("class ");
                if (index == -1 || (index != 0 && !"\t ".Contains(line[index - 1])))
                    continue;
                string result = line.Substring(index + 6);
                index = result.IndexOfAny(new[] { ' ', '\t', ':', '{' });
                if (index == -1)
                    return result;
                else return result.Substring(0, index);
            }
            return null;
        }

        private void ReloadProject(string projName, DTE dte)
        {
            ThreadHelper.ThrowIfNotOnUIThread();
            string solutionName = Path.GetFileNameWithoutExtension(dte.Solution.FullName);
            dte.Windows.Item(EnvDTE.Constants.vsWindowKindSolutionExplorer).Activate();
            ((DTE2)dte).ToolWindows.SolutionExplorer.GetItem($"{solutionName}\\{projName}").Select(vsUISelectionType.vsUISelectionTypeSelect);
            dte.ExecuteCommand("Project.UnloadProject");
            dte.ExecuteCommand("Project.ReloadProject");
        }

        private async Task StatusbarTextAsync(string text, IVsStatusbar statusbar)
        {
            await ThreadHelper.JoinableTaskFactory.SwitchToMainThreadAsync();

            statusbar.IsFrozen(out int frozen);
            if (frozen != 0)
                statusbar.FreezeOutput(0);
            statusbar.SetText(text);
            statusbar.FreezeOutput(1);
        }

        private void MessageBox(string title, string text)
            => VsShellUtilities.ShowMessageBox(
                this.package,
                text,
                title,
                OLEMSGICON.OLEMSGICON_INFO,
                OLEMSGBUTTON.OLEMSGBUTTON_OK,
                OLEMSGDEFBUTTON.OLEMSGDEFBUTTON_FIRST);

        /// <summary>
        /// Splits the string at the first occurrence of the given separator and returns both parts.
        /// </summary>
        private static bool SplitAtFirst(string value, string separator, out string part1, out string part2)
        {
            int index = value.IndexOf(separator);
            if (index == -1)
            {
                part1 = value;
                part2 = "";
                return false;
            }
            else
            {
                part1 = value.Remove(index);
                part2 = value.Remove(0, index + separator.Length);
                return true;
            }
        }

        /// <summary>
        /// Splits the string at the last occurrence of the given separator and returns both parts.
        /// </summary>
        private static bool SplitAtLast(string value, string separator, out string part1, out string part2)
        {
            int index = value.LastIndexOf(separator);
            if (index == -1)
            {
                part1 = value;
                part2 = "";
                return false;
            }
            else
            {
                part1 = value.Remove(index);
                part2 = value.Remove(0, index + separator.Length);
                return true;
            }
        }
    }
}
