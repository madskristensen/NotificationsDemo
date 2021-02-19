using System;
using System.Runtime.InteropServices;
using System.Threading;
using System.Windows;
using EnvDTE;
using Microsoft;
using Microsoft.VisualStudio;
using Microsoft.VisualStudio.Shell;
using Microsoft.VisualStudio.Shell.Interop;
using Microsoft.VisualStudio.TaskStatusCenter;
using Task = System.Threading.Tasks.Task;

namespace NotificationsDemo
{
    [PackageRegistration(UseManagedResourcesOnly = true, AllowsBackgroundLoading = true)]
    [Guid("12e8cc71-83c8-4e8d-8dc1-c8fafae01a92")]
    [ProvideAutoLoad(VSConstants.UICONTEXT.NoSolution_string, PackageAutoLoadFlags.BackgroundLoad)]
    public sealed class NotificationsDemoPackage : AsyncPackage
    {
        protected override async Task InitializeAsync(CancellationToken cancellationToken, IProgress<ServiceProgressData> progress)
        {
            ReportProgress(1, 3);
            await Task.Delay(2000);

            ReportProgress(2, 3);
            await Task.Delay(2000);

            ReportProgress(3, 3);
        }

        private void ReportProgress(int currentSteps, int numberOfSteps)
        {
            UseStatusBarAsync(currentSteps, numberOfSteps)
                .ConfigureAwait(false);
        }

        #region Status bar
        private async Task UseStatusBarAsync(int currentSteps, int numberOfSteps)
        {
            await JoinableTaskFactory.SwitchToMainThreadAsync(DisposalToken);
            var dte = await GetServiceAsync(typeof(DTE)) as DTE;
            Assumes.Present(dte);
            dte.StatusBar.Text = $"Step {currentSteps} of {numberOfSteps} completed";
        }
        #endregion

        #region Status bar with progress
        private async Task UseStatusBarProgressAsync(int currentSteps, int numberOfSteps)
        {
            await JoinableTaskFactory.SwitchToMainThreadAsync(DisposalToken);
            var dte = await GetServiceAsync(typeof(DTE)) as DTE;
            Assumes.Present(dte);

            dte.StatusBar.Progress(true, $"Step {currentSteps} of {numberOfSteps} completed", currentSteps, numberOfSteps);

            if (currentSteps == numberOfSteps)
            {
                await Task.Delay(1000);
                dte.StatusBar.Progress(false);
            }
        }
        #endregion

        #region Threaded Wait Dialog
        private IVsThreadedWaitDialog4 _dialog;

        private async Task UseThreadedWaitDialogAsync(int currentSteps, int numberOfSteps)
        {
            if (_dialog == null)
            {
                await JoinableTaskFactory.SwitchToMainThreadAsync(DisposalToken);
                var factory = await GetServiceAsync(typeof(SVsThreadedWaitDialogFactory)) as IVsThreadedWaitDialogFactory;
                Assumes.Present(factory);

                _dialog = factory.CreateInstance();
                _dialog.StartWaitDialog("Demo", "Working on it...", "", null, "", 1, false, true);
            }

            _dialog.UpdateProgress("In progress", $"Step {currentSteps} of {numberOfSteps} completed", $"Step {currentSteps} of {numberOfSteps} completed", currentSteps, numberOfSteps, true, out _);

            if (currentSteps == numberOfSteps)
            {
                await Task.Delay(1000);
                (_dialog as IDisposable).Dispose();
            }
        }
        #endregion

        #region Output Window
        private IVsOutputWindowPane _pane;
        private async Task UseOutputWindowAsync(int currentSteps, int numberOfSteps)
        {
            if (_pane == null)
            {
                await JoinableTaskFactory.SwitchToMainThreadAsync(DisposalToken);
                var ow = await GetServiceAsync(typeof(SVsOutputWindow)) as IVsOutputWindow;
                Assumes.Present(ow);

                var guid = Guid.NewGuid();
                ow.CreatePane(ref guid, "My output", 1, 1);
                ow.GetPane(ref guid, out _pane);

                _pane.Activate();
                _pane.OutputStringThreadSafe($"Step {currentSteps} of {numberOfSteps} completed\r\n");
            }
            else
            {
                _pane.OutputStringThreadSafe($"Step {currentSteps} of {numberOfSteps} completed\r\n");
            }
        }
        #endregion

        #region Task Status Center
        private ITaskHandler _handler;
        private TaskProgressData _data = default;
        private async Task UseTaskStatusCenterAsync(int currentSteps, int numberOfSteps)
        {
            await ThreadHelper.JoinableTaskFactory.SwitchToMainThreadAsync();

            if (_handler == null)
            {
                var taskStatusCenter = await GetServiceAsync(typeof(SVsTaskStatusCenterService)) as IVsTaskStatusCenterService;
                Assumes.Present(taskStatusCenter);
                var options = default(TaskHandlerOptions);
                options.Title = "A great title";
                options.ActionsAfterCompletion = CompletionActions.None;

                _data.CanBeCanceled = true;

                _handler = taskStatusCenter.PreRegister(options, _data);
                _handler.RegisterTask(Task.Run(async () => { await Task.Delay(5000); }));
            }

            _data.PercentComplete = currentSteps / numberOfSteps * 100;
            _data.ProgressText = $"Step {currentSteps} of {numberOfSteps} completed";

            _handler.Progress.Report(_data);
        }
        #endregion

        #region Message box
        private Task UseMessageBoxAsync(int currentSteps, int numberOfSteps)
        {
            MessageBox.Show("The title", $"Step {currentSteps} of {numberOfSteps} completed");
            return Task.CompletedTask;
        }
        #endregion

        #region VS Message box
        private Task UseVsMessageBoxAsync(int currentSteps, int numberOfSteps)
        {
            VsShellUtilities.ShowMessageBox(
                this,
                $"Step {currentSteps} of {numberOfSteps} completed",
                "The title",
                OLEMSGICON.OLEMSGICON_INFO,
                OLEMSGBUTTON.OLEMSGBUTTON_OK,
                OLEMSGDEFBUTTON.OLEMSGDEFBUTTON_FIRST);

            return Task.CompletedTask;
        }
        #endregion
    }
}
