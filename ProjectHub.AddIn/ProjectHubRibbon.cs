using Microsoft.Office.Tools.Ribbon;
using ProjectHub.AddIn.Forms;
using System;
using System.Drawing;
using System.Drawing.Imaging;
using System.Windows.Forms;

using MSProject = Microsoft.Office.Interop.MSProject;

namespace ProjectHub.AddIn
{
    public partial class ProjectHubRibbon
    {
        private PublishStatusForm _publishStatusForm;

        private Timer _stateTimer;

        private bool _publishRunning;
        private bool _refreshRunning;

        // =========================================================
        // ICON CACHE
        // =========================================================

        private Image _connectGray;
        private Image _connectGreen;

        private Image _publishGray;
        private Image _publishOrange;
        private Image _publishGreen;
        private Image _publishRed;

        private readonly Color StateGray =
            Color.FromArgb(
                145,
                145,
                145);

        private readonly Color StateGreen =
            Color.FromArgb(
                46,
                160,
                67);

        private readonly Color StateOrange =
            Color.FromArgb(
                239,
                91,
                18);

        private readonly Color StateRed =
            Color.FromArgb(
                210,
                45,
                45);

        // =========================================================
        // LOAD
        // =========================================================

        private void ProjectHubRibbon_Load(
            object sender,
            RibbonUIEventArgs e)
        {
            ApplyRibbonStyle();

            PrepareStateIcons();

            StartStateTimer();

            RefreshPublishState();
        }

        // =========================================================
        // RIBBON STYLE
        // =========================================================

        private void ApplyRibbonStyle()
        {
            try
            {
                if (Tabs != null &&
                    Tabs.Count > 0)
                {
                    Tabs[0].Label =
                        "Project Operations";

                    if (Tabs[0].Groups != null &&
                        Tabs[0].Groups.Count > 0)
                    {
                        Tabs[0].Groups[0].Label =
                            "Project Operations";
                    }
                }

                // CONNECT
                btnConnect.ControlSize =
                    Microsoft.Office.Core.RibbonControlSize
                        .RibbonControlSizeLarge;

                btnConnect.Label =
                    "Connect";

                btnConnect.ShowImage =
                    true;

                btnConnect.OfficeImageId =
                    "";

                // PUBLISH
                btnPublish.ControlSize =
                    Microsoft.Office.Core.RibbonControlSize
                        .RibbonControlSizeLarge;

                btnPublish.Label =
                    "Publish";

                btnPublish.ShowImage =
                    true;

                btnPublish.OfficeImageId =
                    "";
            }
            catch
            {
            }
        }

        // =========================================================
        // STATE ICONS
        // =========================================================

        private void PrepareStateIcons()
        {
            try
            {
                Image connectSource =
                    LoadImageResource(
                        "connect");

                Image publishSource =
                    LoadImageResource(
                        "publish");

                if (connectSource != null)
                {
                    _connectGray =
                        TintIcon(
                            connectSource,
                            StateGray);

                    _connectGreen =
                        TintIcon(
                            connectSource,
                            StateGreen);
                }

                if (publishSource != null)
                {
                    _publishGray =
                        TintIcon(
                            publishSource,
                            StateGray);

                    _publishOrange =
                        TintIcon(
                            publishSource,
                            StateOrange);

                    _publishGreen =
                        TintIcon(
                            publishSource,
                            StateGreen);

                    _publishRed =
                        TintIcon(
                            publishSource,
                            StateRed);
                }
            }
            catch
            {
            }
        }

        private static Image TintIcon(
            Image source,
            Color color)
        {
            if (source == null)
                return null;

            Bitmap input =
                new Bitmap(
                    source);

            Bitmap output =
                new Bitmap(
                    input.Width,
                    input.Height,
                    PixelFormat.Format32bppArgb);

            for (int y = 0;
                 y < input.Height;
                 y++)
            {
                for (int x = 0;
                     x < input.Width;
                     x++)
                {
                    Color pixel =
                        input.GetPixel(
                            x,
                            y);

                    if (pixel.A == 0)
                        continue;

                    output.SetPixel(
                        x,
                        y,
                        Color.FromArgb(
                            pixel.A,
                            color.R,
                            color.G,
                            color.B));
                }
            }

            input.Dispose();

            return output;
        }

        // =========================================================
        // TIMER
        // =========================================================

        private void StartStateTimer()
        {
            if (_stateTimer != null)
                return;

            _stateTimer =
                new Timer();

            _stateTimer.Interval =
                1000;

            _stateTimer.Tick +=
                StateTimer_Tick;

            _stateTimer.Start();
        }

        private void StateTimer_Tick(
            object sender,
            EventArgs e)
        {
            RefreshPublishState();
        }

        // =========================================================
        // CONNECT
        // =========================================================

        private void btnConnect_Click(
            object sender,
            RibbonControlEventArgs e)
        {
            using (ConnectForm form =
                   new ConnectForm())
            {
                form.ShowDialog();
            }

            RefreshPublishState();
        }

        // =========================================================
        // PUBLISH
        // =========================================================

        private async void btnPublish_Click(
            object sender,
            RibbonControlEventArgs e)
        {
            if (_publishRunning)
                return;

            if (!CanPublish())
            {
                RefreshPublishState();
                return;
            }

            Guid projectId =
                Globals.ThisAddIn
                    .SelectedCrmProjectId
                    .Value;

            PublishStatusForm statusForm =
                GetPublishStatusForm();

            try
            {
                _publishRunning =
                    true;

                RefreshPublishState();

                statusForm.BeginPublish();

                CrmPublishResult result =
                    await Globals.ThisAddIn.Crm
                        .PublishProjectAsync(
                            projectId,
                            statusForm.Report);

                if (result == null)
                {
                    statusForm.FinishError(
                        "Publish sonucu alınamadı.");

                    return;
                }

                if (result.NoChanges)
                {
                    statusForm.FinishNoChanges();

                    return;
                }

                if (result.Verified)
                {
                    statusForm.FinishSuccess(
                        BuildSuccessMessage(
                            result));

                    return;
                }

                if (result.Failed)
                {
                    statusForm.FinishError(
                        string.IsNullOrWhiteSpace(
                            result.ErrorMessage)
                            ? "Publish işlemi tamamlanamadı."
                            : result.ErrorMessage);

                    return;
                }

                if (result.Blocked)
                {
                    statusForm.FinishBlocked(
                        string.IsNullOrWhiteSpace(
                            result.VerificationMessage)
                            ? "Publish işlemi durduruldu."
                            : result.VerificationMessage);

                    return;
                }

                if (result.VerificationWarning)
                {
                    statusForm.FinishWarning(
                        string.IsNullOrWhiteSpace(
                            result.VerificationMessage)
                            ? "Publish gönderildi ancak henüz doğrulanamadı."
                            : result.VerificationMessage);

                    return;
                }

                statusForm.FinishWarning(
                    "Publish sonucu kesin olarak doğrulanamadı.");
            }
            catch (System.Exception ex)
            {
                statusForm.FinishError(
                    ex.Message);
            }
            finally
            {
                _publishRunning =
                    false;

                RefreshPublishState();
            }
        }

        // =========================================================
        // STATUS FORM
        // =========================================================

        private PublishStatusForm GetPublishStatusForm()
        {
            if (_publishStatusForm == null ||
                _publishStatusForm.IsDisposed)
            {
                _publishStatusForm =
                    new PublishStatusForm();
            }

            return
                _publishStatusForm;
        }

        // =========================================================
        // STATE
        // =========================================================

        public void RefreshPublishState()
        {
            if (_refreshRunning)
                return;

            _refreshRunning =
                true;

            try
            {
                bool connected =
                    Globals.ThisAddIn != null &&
                    Globals.ThisAddIn.Crm != null &&
                    Globals.ThisAddIn.Crm.IsConnected;

                bool loaded =
                    IsSelectedProjectLoaded();

                // =================================================
                // CONNECT
                // =================================================

                btnConnect.Label =
                    "Connect";

                if (connected)
                {
                    if (_connectGreen != null)
                    {
                        btnConnect.Image =
                            _connectGreen;
                    }

                    btnConnect.ScreenTip =
                        "Connected";
                }
                else
                {
                    if (_connectGray != null)
                    {
                        btnConnect.Image =
                            _connectGray;
                    }

                    btnConnect.ScreenTip =
                        "Not connected";
                }

                // =================================================
                // PUBLISH
                // =================================================

                btnPublish.Label =
                    "Publish";

                btnPublish.Enabled =
                    CanPublish() &&
                    !_publishRunning;

                if (!connected ||
                    !loaded ||
                    !Globals.ThisAddIn
                        .SelectedCrmProjectId
                        .HasValue)
                {
                    SetPublishGray();

                    return;
                }

                Guid projectId =
                    Globals.ThisAddIn
                        .SelectedCrmProjectId
                        .Value;

                // ---------------------------------------------
                // Pending / verification required
                // ---------------------------------------------

                if (Globals.ThisAddIn.Crm
                    .HasPendingPublish(
                        projectId))
                {
                    if (_publishRed != null)
                    {
                        btnPublish.Image =
                            _publishRed;
                    }

                    btnPublish.ScreenTip =
                        "Publish verification required";

                    return;
                }

                // ---------------------------------------------
                // Fully synchronized
                // ---------------------------------------------

                if (Globals.ThisAddIn.Crm
                    .IsCurrentProjectPublished(
                        projectId))
                {
                    if (_publishGreen != null)
                    {
                        btnPublish.Image =
                            _publishGreen;
                    }

                    btnPublish.ScreenTip =
                        "Up to date";

                    return;
                }

                // ---------------------------------------------
                // Local changes exist
                // ---------------------------------------------

                if (_publishOrange != null)
                {
                    btnPublish.Image =
                        _publishOrange;
                }

                btnPublish.ScreenTip =
                    "Changes ready to publish";
            }
            catch
            {
                btnConnect.Label =
                    "Connect";

                btnPublish.Label =
                    "Publish";

                if (_connectGray != null)
                {
                    btnConnect.Image =
                        _connectGray;
                }

                SetPublishGray();
            }
            finally
            {
                _refreshRunning =
                    false;
            }
        }

        private void SetPublishGray()
        {
            if (_publishGray != null)
            {
                btnPublish.Image =
                    _publishGray;
            }

            btnPublish.ScreenTip =
                "Publish unavailable";
        }

        // =========================================================
        // PROJECT STATE
        // =========================================================

        private bool IsSelectedProjectLoaded()
        {
            try
            {
                if (Globals.ThisAddIn == null ||
                    !Globals.ThisAddIn
                        .SelectedCrmProjectId
                        .HasValue)
                {
                    return false;
                }

                MSProject.Project project =
                    Globals.ThisAddIn
                        .Application
                        .ActiveProject;

                if (project == null)
                    return false;

                MSProject.Task summary =
                    project.ProjectSummaryTask;

                if (summary == null)
                    return false;

                Guid projectId;

                if (!Guid.TryParse(
                    summary.Text30,
                    out projectId))
                {
                    return false;
                }

                return
                    projectId ==
                    Globals.ThisAddIn
                        .SelectedCrmProjectId
                        .Value;
            }
            catch
            {
                return false;
            }
        }

        // =========================================================
        // CAN PUBLISH
        // =========================================================

        private bool CanPublish()
        {
            try
            {
                if (_publishRunning)
                    return false;

                if (Globals.ThisAddIn == null ||
                    Globals.ThisAddIn.Crm == null ||
                    !Globals.ThisAddIn.Crm.IsConnected ||
                    !Globals.ThisAddIn
                        .SelectedCrmProjectId
                        .HasValue)
                {
                    return false;
                }

                if (!IsSelectedProjectLoaded())
                    return false;

                MSProject.Project project =
                    Globals.ThisAddIn
                        .Application
                        .ActiveProject;

                if (project == null ||
                    project.Tasks == null)
                {
                    return false;
                }

                foreach (MSProject.Task task
                         in project.Tasks)
                {
                    if (task != null &&
                        !string.IsNullOrWhiteSpace(
                            task.Name))
                    {
                        return true;
                    }
                }

                return false;
            }
            catch
            {
                return false;
            }
        }

        // =========================================================
        // RESULT
        // =========================================================

        private static string BuildSuccessMessage(
            CrmPublishResult result)
        {
            return
                "CRM güncellemesi doğrulandı." +
                "\r\n\r\n" +

                "Project Progress: " +
                (
                    result.UpdatedProjectFields > 0
                        ? "updated"
                        : "unchanged"
                ) +
                "\r\n" +

                "Tasks: " +
                result.CreatedTasks +
                " new, " +
                result.UpdatedTasks +
                " updated, " +
                result.DeletedTasks +
                " deleted" +
                "\r\n" +

                "Predecessors: " +
                result.CreatedDependencies +
                " new, " +
                result.DeletedDependencies +
                " deleted" +
                "\r\n" +

                "Resources: " +
                result.CreatedAssignments +
                " assigned, " +
                result.DeletedAssignments +
                " removed";
        }

        // =========================================================
        // RESOURCES
        // =========================================================

        private static Image LoadImageResource(
            string name)
        {
            try
            {
                return
                    global::ProjectHub.AddIn
                        .Properties
                        .Resources
                        .ResourceManager
                        .GetObject(
                            name)
                    as Image;
            }
            catch
            {
                return null;
            }
        }
    }
}