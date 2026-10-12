using ProjectHub.AddIn;
using System;
using System.Collections.Generic;
using System.Drawing;
using System.Linq;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace ProjectHub.AddIn.Forms
{
    public partial class ConnectForm : Form
    {
        private DataGridView gridProjects;

        private Label lblProjectCount;
        private Label lblConnectionState;
        private Label lblBuildInfo;

        private PictureBox picProductLogo;
        private PictureBox picCompanyLogo;

        private bool _busy;

        private readonly Color Accent =
            Color.FromArgb(239, 91, 18);

        private readonly Color AccentHover =
            Color.FromArgb(211, 72, 10);

        private readonly Color TextPrimary =
            Color.FromArgb(32, 31, 30);

        private readonly Color TextSecondary =
            Color.FromArgb(96, 94, 92);

        private readonly Color Border =
            Color.FromArgb(190, 190, 190);

        public ConnectForm()
        {
            InitializeComponent();

            BuildLayout();
            WireEvents();
        }

        private void BuildLayout()
        {
            SuspendLayout();

            Controls.Clear();

            Text =
                "Project Operation Connector";

            StartPosition =
                FormStartPosition.CenterScreen;

            FormBorderStyle =
                FormBorderStyle.FixedSingle;

            MaximizeBox =
                false;

            MinimizeBox =
                true;

            ShowIcon =
                false;

            AutoScaleMode =
                AutoScaleMode.Dpi;

            ClientSize =
                new Size(
                    760,
                    485);

            BackColor =
                Color.White;

            Font =
                new Font(
                    "Segoe UI",
                    9F,
                    FontStyle.Regular);

            // =====================================================
            // HEADER
            // =====================================================

            picProductLogo =
                new PictureBox
                {
                    Left = 20,
                    Top = 16,
                    Width = 58,
                    Height = 58,
                    SizeMode =
                        PictureBoxSizeMode.Zoom,
                    Image =
                        LoadImageResource(
                            "PO_logo")
                };

            Controls.Add(
                picProductLogo);

            Label lblTitle =
                new Label
                {
                    Text =
                        "Project Operation Connector",

                    Left = 92,
                    Top = 27,
                    Width = 500,
                    Height = 38,

                    ForeColor =
                        TextPrimary,

                    Font =
                        new Font(
                            "Segoe UI",
                            19F,
                            FontStyle.Regular)
                };

            Controls.Add(
                lblTitle);

            // =====================================================
            // URL
            // =====================================================

            Label lblUrl =
                new Label
                {
                    Text =
                        "Dynamics URL",

                    Left = 20,
                    Top = 94,
                    Width = 105,
                    Height = 27,

                    TextAlign =
                        ContentAlignment.MiddleLeft,

                    ForeColor =
                        TextPrimary,

                    Font =
                        new Font(
                            "Segoe UI",
                            10F)
                };

            Controls.Add(
                lblUrl);

            txtUrl.Left =
                126;

            txtUrl.Top =
                94;

            txtUrl.Width =
                480;

            txtUrl.Height =
                28;

            txtUrl.Font =
                new Font(
                    "Segoe UI",
                    10F);

            Controls.Add(
                txtUrl);

            btnConnect.Text =
                "Connect";

            btnConnect.Left =
                616;

            btnConnect.Top =
                92;

            btnConnect.Width =
                124;

            btnConnect.Height =
                31;

            btnConnect.FlatStyle =
                FlatStyle.Flat;

            btnConnect.FlatAppearance.BorderSize =
                0;

            btnConnect.FlatAppearance.MouseOverBackColor =
                AccentHover;

            btnConnect.BackColor =
                Accent;

            btnConnect.ForeColor =
                Color.White;

            btnConnect.Font =
                new Font(
                    "Segoe UI",
                    9.5F,
                    FontStyle.Bold);

            Controls.Add(
                btnConnect);

            lblConnectionState =
                new Label
                {
                    Text = "",

                    Left = 126,
                    Top = 126,
                    Width = 480,
                    Height = 18,

                    ForeColor =
                        TextSecondary,

                    Font =
                        new Font(
                            "Segoe UI",
                            8F)
                };

            Controls.Add(
                lblConnectionState);

            // =====================================================
            // GRID
            // =====================================================

            gridProjects =
                new DataGridView
                {
                    Left = 20,
                    Top = 151,

                    Width = 720,
                    Height = 255,

                    BackgroundColor =
                        Color.White,

                    BorderStyle =
                        BorderStyle.FixedSingle,

                    AllowUserToAddRows =
                        false,

                    AllowUserToDeleteRows =
                        false,

                    AllowUserToResizeRows =
                        false,

                    MultiSelect =
                        false,

                    ReadOnly =
                        true,

                    SelectionMode =
                        DataGridViewSelectionMode.FullRowSelect,

                    RowHeadersVisible =
                        false,

                    AutoGenerateColumns =
                        false,

                    EnableHeadersVisualStyles =
                        false,

                    ColumnHeadersHeight =
                        31
                };

            gridProjects.RowTemplate.Height =
                28;

            gridProjects.ColumnHeadersDefaultCellStyle.BackColor =
                Color.FromArgb(
                    245,
                    245,
                    245);

            gridProjects.ColumnHeadersDefaultCellStyle.ForeColor =
                TextPrimary;

            gridProjects.ColumnHeadersDefaultCellStyle.Font =
                new Font(
                    "Segoe UI",
                    9F,
                    FontStyle.Regular);

            gridProjects.DefaultCellStyle.Font =
                new Font(
                    "Segoe UI",
                    9F);

            gridProjects.DefaultCellStyle.ForeColor =
                TextPrimary;

            gridProjects.DefaultCellStyle.BackColor =
                Color.White;

            gridProjects.DefaultCellStyle.SelectionBackColor =
                Color.FromArgb(
                    255,
                    220,
                    195);

            gridProjects.DefaultCellStyle.SelectionForeColor =
                Color.Black;

            gridProjects.GridColor =
                Color.FromArgb(
                    215,
                    215,
                    215);

            gridProjects.Columns.Add(
                CreateTextColumn(
                    "ProjectName",
                    "Project Name",
                    240,
                    DataGridViewAutoSizeColumnMode.Fill));

            gridProjects.Columns.Add(
                CreateTextColumn(
                    "ProjectManager",
                    "Project Manager",
                    160,
                    DataGridViewAutoSizeColumnMode.None));

            gridProjects.Columns.Add(
                CreateTextColumn(
                    "Start",
                    "Start",
                    105,
                    DataGridViewAutoSizeColumnMode.None));

            gridProjects.Columns.Add(
                CreateTextColumn(
                    "Finish",
                    "Finish",
                    105,
                    DataGridViewAutoSizeColumnMode.None));

            gridProjects.Columns.Add(
                CreateTextColumn(
                    "PercentComplete",
                    "% Complete",
                    105,
                    DataGridViewAutoSizeColumnMode.None));

            gridProjects.CellFormatting +=
                GridProjects_CellFormatting;

            gridProjects.CellDoubleClick +=
                GridProjects_CellDoubleClick;

            Controls.Add(
                gridProjects);

            // =====================================================
            // FOOTER
            // =====================================================

            picCompanyLogo =
                new PictureBox
                {
                    Left = 20,
                    Top = 419,

                    Width = 225,
                    Height = 48,

                    SizeMode =
                        PictureBoxSizeMode.Zoom,

                    Image =
                        LoadImageResource(
                            "TuruncuTeknolojiLogo")
                };

            Controls.Add(
                picCompanyLogo);

            lblProjectCount =
                new Label
                {
                    Text = "",

                    Left = 460,
                    Top = 416,

                    Width = 150,
                    Height = 20,

                    TextAlign =
                        ContentAlignment.MiddleRight,

                    ForeColor =
                        TextSecondary,

                    Font =
                        new Font(
                            "Segoe UI",
                            8F)
                };

            Controls.Add(
                lblProjectCount);

            lblBuildInfo =
                new Label
                {
                    Text =
                        "Build " +
                        GetBuildVersion(),

                    Left = 255,
                    Top = 444,

                    Width = 190,
                    Height = 18,

                    TextAlign =
                        ContentAlignment.MiddleLeft,

                    ForeColor =
                        TextSecondary,

                    Font =
                        new Font(
                            "Segoe UI",
                            8F)
                };

            Controls.Add(
                lblBuildInfo);

            btnOpen.Text =
                "Open";

            btnOpen.Left =
                616;

            btnOpen.Top =
                423;

            btnOpen.Width =
                124;

            btnOpen.Height =
                34;

            btnOpen.FlatStyle =
                FlatStyle.Flat;

            btnOpen.FlatAppearance.BorderColor =
                Accent;

            btnOpen.FlatAppearance.BorderSize =
                1;

            btnOpen.FlatAppearance.MouseOverBackColor =
                Color.FromArgb(
                    255,
                    240,
                    230);

            btnOpen.BackColor =
                Color.White;

            btnOpen.ForeColor =
                TextPrimary;

            btnOpen.Font =
                new Font(
                    "Segoe UI",
                    9.5F);

            Controls.Add(
                btnOpen);

            AcceptButton =
                btnConnect;

            ResumeLayout(
                true);
        }

        private static DataGridViewTextBoxColumn CreateTextColumn(
            string name,
            string header,
            int width,
            DataGridViewAutoSizeColumnMode sizeMode)
        {
            return
                new DataGridViewTextBoxColumn
                {
                    Name =
                        name,

                    HeaderText =
                        header,

                    Width =
                        width,

                    AutoSizeMode =
                        sizeMode,

                    SortMode =
                        DataGridViewColumnSortMode.Automatic,

                    ReadOnly =
                        true
                };
        }

        // =========================================================
        // EVENTS
        // =========================================================

        private void WireEvents()
        {
            btnConnect.Click -=
                btnConnect_Click;

            btnConnect.Click +=
                btnConnect_Click;

            btnOpen.Click -=
                btnOpen_Click;

            btnOpen.Click +=
                btnOpen_Click;

            Shown -=
                ConnectForm_Shown;

            Shown +=
                ConnectForm_Shown;
        }

        // =========================================================
        // SHOWN
        // =========================================================

        private async void ConnectForm_Shown(
            object sender,
            EventArgs e)
        {
            if (Globals.ThisAddIn == null ||
                Globals.ThisAddIn.Crm == null)
            {
                return;
            }

            if (!Globals.ThisAddIn.Crm.IsConnected)
            {
                UpdateConnectionState(
                    false);

                return;
            }

            txtUrl.Text =
                Globals.ThisAddIn.Crm.ConnectedUrl
                ?? "";

            UpdateConnectionState(
                true);

            List<CrmProject> cached =
                Globals.ThisAddIn.Crm.CachedProjects;

            if (cached != null &&
                cached.Count > 0)
            {
                ShowProjects(
                    cached);

                return;
            }

            await LoadProjectsAsync();
        }

        // =========================================================
        // CONNECT
        // =========================================================

        private async void btnConnect_Click(
            object sender,
            EventArgs e)
        {
            if (_busy)
                return;

            string url =
                txtUrl.Text != null
                    ? txtUrl.Text.Trim()
                    : "";

            if (string.IsNullOrWhiteSpace(
                url))
            {
                MessageBox.Show(
                    "Dynamics 365 URL girin.",
                    "Project Operation Connector",
                    MessageBoxButtons.OK,
                    MessageBoxIcon.Warning);

                return;
            }

            SetBusy(
                true,
                "Connecting...");

            try
            {
                bool connected =
                    await Globals.ThisAddIn.Crm
                        .ConnectAsync(
                            url);

                if (!connected)
                {
                    if (Globals.ThisAddIn.Crm
                        .LastLoginCancelled)
                    {
                        return;
                    }

                    MessageBox.Show(
                        string.IsNullOrWhiteSpace(
                            Globals.ThisAddIn.Crm.LastError)
                            ? "Dynamics 365 bağlantısı kurulamadı."
                            : Globals.ThisAddIn.Crm.LastError,
                        "Project Operation Connector",
                        MessageBoxButtons.OK,
                        MessageBoxIcon.Error);

                    UpdateConnectionState(
                        false);

                    return;
                }

                txtUrl.Text =
                    Globals.ThisAddIn.Crm.ConnectedUrl
                    ?? url;

                UpdateConnectionState(
                    true);

                await LoadProjectsAsync();
            }
            catch (Exception ex)
            {
                MessageBox.Show(
                    ex.Message,
                    "Project Operation Connector",
                    MessageBoxButtons.OK,
                    MessageBoxIcon.Error);
            }
            finally
            {
                SetBusy(
                    false,
                    null);
            }
        }

        // =========================================================
        // LOAD PROJECTS
        // =========================================================

        private async Task LoadProjectsAsync()
        {
            if (Globals.ThisAddIn == null ||
                Globals.ThisAddIn.Crm == null ||
                !Globals.ThisAddIn.Crm.IsConnected)
            {
                return;
            }

            SetBusy(
                true,
                "Loading projects...");

            try
            {
                List<CrmProject> projects =
                    await Globals.ThisAddIn.Crm
                        .GetProjectsAsync();

                ShowProjects(
                    projects);
            }
            catch (Exception ex)
            {
                MessageBox.Show(
                    ex.Message,
                    "Project Operation Connector",
                    MessageBoxButtons.OK,
                    MessageBoxIcon.Error);
            }
            finally
            {
                SetBusy(
                    false,
                    null);
            }
        }

        // =========================================================
        // GRID
        // =========================================================

        private void ShowProjects(
            List<CrmProject> projects)
        {
            projects =
                projects
                ?? new List<CrmProject>();

            Guid? previous =
                Globals.ThisAddIn
                    .SelectedCrmProjectId;

            gridProjects.SuspendLayout();

            try
            {
                gridProjects.Rows.Clear();

                foreach (CrmProject project
                         in projects)
                {
                    int index =
                        gridProjects.Rows.Add(
                            project.Name ?? "",
                            project.ProjectManager ?? "",
                            project.Start,
                            project.Finish,
                            project.PercentComplete);

                    gridProjects.Rows[
                        index].Tag =
                        project;
                }

                gridProjects.ClearSelection();

                DataGridViewRow selectedRow =
                    null;

                if (previous.HasValue)
                {
                    selectedRow =
                        gridProjects.Rows
                            .Cast<DataGridViewRow>()
                            .FirstOrDefault(
                                row =>
                                {
                                    CrmProject project =
                                        row.Tag
                                        as CrmProject;

                                    return
                                        project != null &&
                                        project.Id ==
                                        previous.Value;
                                });
                }

                if (selectedRow == null &&
                    gridProjects.Rows.Count > 0)
                {
                    selectedRow =
                        gridProjects.Rows[0];
                }

                if (selectedRow != null)
                {
                    selectedRow.Selected =
                        true;

                    gridProjects.CurrentCell =
                        selectedRow.Cells[0];
                }
            }
            finally
            {
                gridProjects.ResumeLayout();
            }

            lblProjectCount.Text =
                projects.Count +
                " project" +
                (
                    projects.Count == 1
                        ? ""
                        : "s"
                );

            btnOpen.Enabled =
                projects.Count > 0;
        }

        private void GridProjects_CellFormatting(
            object sender,
            DataGridViewCellFormattingEventArgs e)
        {
            if (e.RowIndex < 0)
                return;

            string column =
                gridProjects.Columns[
                    e.ColumnIndex].Name;

            if (column == "Start" ||
                column == "Finish")
            {
                if (e.Value is DateTime)
                {
                    e.Value =
                        ((DateTime)e.Value)
                        .ToString(
                            "dd.MM.yyyy");

                    e.FormattingApplied =
                        true;
                }

                return;
            }

            if (column ==
                "PercentComplete")
            {
                if (e.Value != null)
                {
                    e.Value =
                        Convert.ToInt32(
                            e.Value) +
                        "%";

                    e.FormattingApplied =
                        true;
                }
            }
        }

        private void GridProjects_CellDoubleClick(
            object sender,
            DataGridViewCellEventArgs e)
        {
            if (_busy ||
                e.RowIndex < 0)
            {
                return;
            }

            btnOpen.PerformClick();
        }

        private CrmProject GetSelectedProject()
        {
            if (gridProjects.CurrentRow ==
                null)
            {
                return null;
            }

            return
                gridProjects.CurrentRow.Tag
                as CrmProject;
        }

        // =========================================================
        // OPEN
        // =========================================================

        private async void btnOpen_Click(
            object sender,
            EventArgs e)
        {
            if (_busy)
                return;

            CrmProject selected =
                GetSelectedProject();

            if (selected == null)
            {
                MessageBox.Show(
                    "Bir proje seçin.",
                    "Project Operation Connector",
                    MessageBoxButtons.OK,
                    MessageBoxIcon.Warning);

                return;
            }

            SetBusy(
                true,
                "Opening project...");

            try
            {
                CrmScheduleData schedule =
                    await Globals.ThisAddIn.Crm
                        .GetScheduleDataAsync(
                            selected.Id);

                if (schedule == null ||
                    schedule.Tasks == null ||
                    schedule.Tasks.Entities.Count == 0)
                {
                    MessageBox.Show(
                        "Seçilen projede task bulunamadı.",
                        "Project Operation Connector",
                        MessageBoxButtons.OK,
                        MessageBoxIcon.Warning);

                    return;
                }

                Globals.ThisAddIn
                    .SelectedCrmProjectId =
                    selected.Id;

                await Task.Yield();

                MsProjectService.ImportProject(
                    schedule);

                Globals.ThisAddIn.Crm
                    .SetLoadedBaseline(
                        schedule);

                try
                {
                    Globals.Ribbons
                        .ProjectHubRibbon
                        .RefreshPublishState();
                }
                catch
                {
                }

                DialogResult =
                    DialogResult.OK;

                Close();
            }
            catch (Exception ex)
            {
                MessageBox.Show(
                    ex.Message,
                    "Project Operation Connector",
                    MessageBoxButtons.OK,
                    MessageBoxIcon.Error);
            }
            finally
            {
                SetBusy(
                    false,
                    null);
            }
        }

        // =========================================================
        // STATE
        // =========================================================

        private void UpdateConnectionState(
            bool connected)
        {
            if (connected)
            {
                lblConnectionState.Text =
                    "● Connected";

                lblConnectionState.ForeColor =
                    Color.FromArgb(
                        16,
                        124,
                        16);
            }
            else
            {
                lblConnectionState.Text =
                    "";

                lblConnectionState.ForeColor =
                    TextSecondary;
            }
        }

        private void SetBusy(
            bool busy,
            string text)
        {
            _busy =
                busy;

            txtUrl.Enabled =
                !busy;

            gridProjects.Enabled =
                !busy;

            btnConnect.Enabled =
                !busy;

            btnOpen.Enabled =
                !busy &&
                gridProjects.Rows.Count > 0;

            UseWaitCursor =
                busy;

            if (busy)
            {
                lblConnectionState.Text =
                    text ?? "";

                lblConnectionState.ForeColor =
                    Accent;
            }
            else
            {
                UpdateConnectionState(
                    Globals.ThisAddIn != null &&
                    Globals.ThisAddIn.Crm != null &&
                    Globals.ThisAddIn.Crm.IsConnected);
            }
        }

        // =========================================================
        // RESOURCES
        // =========================================================

        private static string GetBuildVersion()
        {
            try
            {
                Version version =
                    typeof(ConnectForm)
                        .Assembly
                        .GetName()
                        .Version;

                return version != null
                    ? version.ToString()
                    : "unknown";
            }
            catch
            {
                return "unknown";
            }
        }

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