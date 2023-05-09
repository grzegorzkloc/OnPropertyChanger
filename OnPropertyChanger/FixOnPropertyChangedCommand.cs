using EnvDTE;
using EnvDTE80;
using Microsoft.VisualStudio.Shell;
using Microsoft.VisualStudio.Shell.Interop;
using System;
using System.ComponentModel.Design;
using System.Text.RegularExpressions;
using Package = Microsoft.VisualStudio.Shell.Package;
using Task = System.Threading.Tasks.Task;

namespace OnPropertyChanger
{
	internal sealed class FixOnPropertyChangedCommand
	{
		public const int CommandId = 0x0100;
		public static readonly Guid CommandSet = new Guid("c8593044-2e2a-460f-bb25-c87055eab71f");

		private readonly AsyncPackage _package;

		private const string DialogTitle = "Fix OnPropertyChanged()";
		private const string Pattern = @"OnProperty(.+)\(\""(.+)\""\)";
		private const string Replacement = "OnProperty$1(nameof($2))";

		private FixOnPropertyChangedCommand(AsyncPackage package, OleMenuCommandService commandService)
		{
			_package = package ?? throw new ArgumentNullException(nameof(package));
			commandService = commandService ?? throw new ArgumentNullException(nameof(commandService));

			var menuCommandId = new CommandID(CommandSet, CommandId);
			var menuItem = new MenuCommand(Execute, menuCommandId);
			commandService.AddCommand(menuItem);
		}

		public static FixOnPropertyChangedCommand Instance { get; private set; }

		private Microsoft.VisualStudio.Shell.IAsyncServiceProvider ServiceProvider => _package;

		public static async Task InitializeAsync(AsyncPackage package)
		{
			// Switch to the main thread - the call to AddCommand in FixOnPropertyChangedCommand's constructor requires
			// the UI thread.
			await ThreadHelper.JoinableTaskFactory.SwitchToMainThreadAsync(package.DisposalToken);

			var commandService = await package.GetServiceAsync(typeof(IMenuCommandService)) as OleMenuCommandService;
			Instance = new FixOnPropertyChangedCommand(package, commandService);
		}

#pragma warning disable VSTHRD100 //I know what I am doing ;)
		private async void Execute(object sender, EventArgs e)
		{
			try
			{
				await ExecuteAsync();
			}
			catch (Exception exception)
			{
				ShowDialog($"Error while trying to fix OnPropertyChanged(): {exception.Message}");
			}
		}
#pragma warning restore VSTHRD100

		private async Task ExecuteAsync()
		{
			await ThreadHelper.JoinableTaskFactory.SwitchToMainThreadAsync(_package.DisposalToken);
			var dte = Package.GetGlobalService(typeof(SDTE)) as DTE2;
			var activeDocument = dte?.ActiveDocument;
			if (activeDocument == null)
			{
				ShowDialog("Error: No active document open.");
				return;
			}

			if (!(activeDocument.Object("TextDocument") is TextDocument textDocument))
			{
				ShowDialog("Error: active document cannot be converted to text document.");
				return;
			}

			var documentStartPoint = textDocument.StartPoint.CreateEditPoint();
			var wholeText = documentStartPoint.GetText(textDocument.EndPoint);
			if (Regex.IsMatch(wholeText, Pattern))
			{
				var documentAfterFixes = Regex.Replace(wholeText, Pattern, Replacement);
				documentStartPoint.ReplaceText(textDocument.EndPoint, documentAfterFixes, (int)vsEPReplaceTextOptions.vsEPReplaceTextAutoformat);
				ShowDialog("Fixed as requested. Please make sure that property name inside nameof() is correct.");
			}
			else
			{
				ShowDialog("There was noting to fix.");
			}
		}

		private void ShowDialog(string message)
		{
			VsShellUtilities.ShowMessageBox(
				_package,
				message,
				DialogTitle,
				OLEMSGICON.OLEMSGICON_INFO,
				OLEMSGBUTTON.OLEMSGBUTTON_OK,
				OLEMSGDEFBUTTON.OLEMSGDEFBUTTON_FIRST);
		}
	}
}
