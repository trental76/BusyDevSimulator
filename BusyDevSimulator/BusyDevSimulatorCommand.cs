using EnvDTE;
using EnvDTE80;
using Microsoft.VisualStudio.Shell;
using Microsoft.VisualStudio.Shell.Interop;
using Microsoft.VisualStudio.Threading;
using System;
using System.ComponentModel.Design;
using System.Diagnostics;
using System.Globalization;
using System.IO;
using System.Threading;
using System.Threading.Tasks;
using System.Windows.Forms;
using Task = System.Threading.Tasks.Task;

namespace BusyDevSimulator
{
    /// <summary>
    /// Command handler
    /// </summary>
    internal sealed class BusyDevSimulatorCommand
    {
        /// <summary>
        /// Command ID.
        /// </summary>
        public const int CommandId = 0x0100;

        /// <summary>
        /// Command menu group (command set GUID).
        /// </summary>
        public static readonly Guid CommandSet = new Guid("8164fed2-07a0-4058-a246-3a17fc298762");

        /// <summary>
        /// VS Package that provides this command, not null.
        /// </summary>
        private readonly AsyncPackage package;

        /// <summary>
        /// Initializes a new instance of the <see cref="BusyDevSimulatorCommand"/> class.
        /// Adds our command handlers for menu (commands must exist in the command table file)
        /// </summary>
        /// <param name="package">Owner package, not null.</param>
        /// <param name="commandService">Command service to add command to, not null.</param>
        private BusyDevSimulatorCommand(AsyncPackage package, OleMenuCommandService commandService)
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
        public static BusyDevSimulatorCommand Instance
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
            // Switch to the main thread - the call to AddCommand in BusyDevSimulatorCommand's constructor requires
            // the UI thread.
            await ThreadHelper.JoinableTaskFactory.SwitchToMainThreadAsync(package.DisposalToken);

            OleMenuCommandService commandService = await package.GetServiceAsync(typeof(IMenuCommandService)) as OleMenuCommandService;
            Instance = new BusyDevSimulatorCommand(package, commandService);
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
            //ThreadHelper.ThrowIfNotOnUIThread();
            //string message = string.Format(CultureInfo.CurrentCulture, "Inside {0}.MenuItemCallback()", this.GetType().FullName);
            //string title = "BusyDevSimulatorCommand";

            //// Show a message box to prove we were here
            //VsShellUtilities.ShowMessageBox(
            //    this.package,
            //    message,
            //    title,
            //    OLEMSGICON.OLEMSGICON_INFO,
            //    OLEMSGBUTTON.OLEMSGBUTTON_OK,
            //    OLEMSGDEFBUTTON.OLEMSGDEFBUTTON_FIRST);

            ThreadHelper.ThrowIfNotOnUIThread();

            ThreadHelper.JoinableTaskFactory.Run(async () =>
            {
                await ThreadHelper.JoinableTaskFactory.SwitchToMainThreadAsync();
                var dte = Package.GetGlobalService(typeof(DTE)) as DTE2;
                if (dte == null) return;

                // 1. Емуляція активності
                var activeDocBefore = dte.ActiveDocument;
                await Task.Delay(1000);

                //dte.ExecuteCommand("View.SolutionExplorer", "document");
                //await Task.Delay(1000);

                //SendKeys.Send("{ENTER}");
                //await Task.Delay(4000);

                //activeDocBefore.Activate();
                //dte.ExecuteCommand("Edit.GoToDefinition");
                //await Task.Delay(700);

                //dte.ExecuteCommand("Edit.Find", "public");
                //await Task.Delay(800);
                //SendKeys.Send("{ENTER}");
                //await Task.Delay(2000);

                //activeDocBefore.Activate();

                dte.ExecuteCommand("Window.NextTab");
                await Task.Delay(1000);

                dte.ExecuteCommand("Window.NextTab");
                await Task.Delay(4000);

                dte.ExecuteCommand("Edit.GoToDefinition");
                await Task.Delay(700);

                dte.ExecuteCommand("Edit.Copy");
                await Task.Delay(2000);

                dte.ExecuteCommand("Edit.Paste");
                await Task.Delay(3000);

                // 2. Зміна файлу
                var activeDoc = dte.ActiveDocument;
                var selection = (TextSelection)activeDoc.Selection;
                selection.EndOfDocument();
                selection.NewLine();
                selection.Text = "// Temporary change - will be reverted";

                await Task.Delay(1000);

                // 3. Git Revert
                string solutionPath = Path.GetDirectoryName(dte.Solution.FullName);
                if (!string.IsNullOrEmpty(solutionPath))
                {
                    var psi = new ProcessStartInfo
                    {
                        FileName = "git",
                        Arguments = "restore .",
                        WorkingDirectory = solutionPath,
                        RedirectStandardOutput = true,
                        RedirectStandardError = true,
                        CreateNoWindow = true,
                        UseShellExecute = false
                    };

                    try
                    {
                        using (var process = System.Diagnostics.Process.Start(psi))
                        {
                            string output = await process.StandardOutput.ReadToEndAsync();
                            string errors = await process.StandardError.ReadToEndAsync();
                            await process.WaitForExitAsync();

                            System.Diagnostics.Debug.WriteLine("Git restore output: " + output);
                            if (!string.IsNullOrWhiteSpace(errors))
                                System.Diagnostics.Debug.WriteLine("Git restore errors: " + errors);
                        }
                    }
                    catch (Exception ex)
                    {
                        System.Diagnostics.Debug.WriteLine("Git restore failed: " + ex.Message);
                    }
                }
            });
        }
    }
}
