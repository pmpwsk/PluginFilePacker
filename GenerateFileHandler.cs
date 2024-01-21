using EnvDTE;
using Microsoft.VisualStudio.Shell;
using Microsoft.VisualStudio.Shell.Interop;
using System;
using System.Collections.Generic;
using System.ComponentModel.Design;
using System.IO;
using System.Linq;
using System.Text;
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

                        await GenerateAsync(project.Name, projPath, options);

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

        private async Task GenerateAsync(string projName, string projPath, Options options)
        {
            string filesPath = projPath + "/Files";
            if (!Directory.Exists(filesPath))
                throw new Exception("Directory 'Files' not found!");

            Dictionary<string,string> resourceFiles = options.UseBase64 ? null : new Dictionary<string, string>();

            string namespace_ = DetectNamespace(projPath, options);
            string class_ = DetectClass(projPath, projName);
            using (StreamWriter writer = new StreamWriter($"{projPath}/FileHandler.cs", false, Encoding.UTF8))
            {
                try
                {
                    Dictionary<string, string> timestamps = new Dictionary<string, string>();
                    IEnumerable<string> textTypes = options.TextExtensionsClean;

                    string wfNamespacePrefix = (namespace_ == "uwap.WebFramework" || namespace_.StartsWith("uwap.WebFramework.")) ? "" : "uwap.WebFramework.";

                    await writer.WriteLineAsync($"namespace {namespace_};");
                    await writer.WriteLineAsync();
                    await writer.WriteLineAsync($"public partial class {class_} : {(namespace_ != "uwap.WebFramework.Plugins" ? "uwap.WebFramework.Plugins." : "")}Plugin");
                    await writer.WriteLineAsync("{");
                    await writer.WriteLineAsync("    public override byte[]? GetFile(string relPath, string pathPrefix, string domain)");
                    await writer.WriteLineAsync("        => relPath switch");
                    await writer.WriteLineAsync("        {");
                    foreach (string file in Directory.GetFiles(filesPath, "*", SearchOption.AllDirectories))
                    {
                        string relPath = file.Remove(0, filesPath.Length).Replace('\\', '/');
                        Console.WriteLine("  " + relPath);
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
                            string key = Convert.ToBase64String(Encoding.UTF8.GetBytes(relPath)).TrimEnd('=');
                            resourceFiles[key] = file;
                            await writer.WriteLineAsync($"            \"{relPath}\" => {projName}.Properties.PluginFiles.{key},");
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
        }

        private string DetectNamespace(string projPath, Options options)
        {
            string result = null;
            if (File.Exists(projPath + "/FileHandler.cs"))
                result = DetectNamespaceFromFile(projPath + "/FileHandler.cs");
            if (result == null && File.Exists(projPath + "/FileHandlerCustom.cs"))
                result = DetectNamespaceFromFile(projPath + "/FileHandlerCustom.cs");
            if (result == null)
                result = options.DefaultNamespace;
            return result;
        }

        private string DetectNamespaceFromFile(string path)
        {
            foreach (string line in File.ReadAllLines(path))
                if (line.StartsWith("namespace "))
                    return line.Remove(0, 10).Split(' ', '{', ';').First();
            return null;
        }

        private string DetectClass(string projPath, string projName)
        {
            string result = null;
            if (File.Exists(projPath + "/FileHandler.cs"))
                result = DetectClassFromFile(projPath + "/FileHandler.cs");
            if (result == null && File.Exists(projPath + "/FileHandlerCustom.cs"))
                result = DetectClassFromFile(projPath + "/FileHandlerCustom.cs");
            if (result == null)
                result = projName;
            return result;
        }

        private string DetectClassFromFile(string path)
        {
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
    }
}
