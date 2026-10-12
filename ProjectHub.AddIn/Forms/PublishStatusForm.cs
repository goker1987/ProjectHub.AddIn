using ProjectHub.AddIn;
using System;
using System.Drawing;
using System.Text;
using System.Windows.Forms;

namespace ProjectHub.AddIn.Forms
{
    public class PublishStatusForm : Form
    {
        private PictureBox picProductLogo;
        private PictureBox picCompanyLogo;

        private Label lblTitle;
        private Label lblStatus;

        private Panel statusPanel;
        private ProgressBar progressBar;

        private Label lblDetails;
        private TextBox txtLog;

        private Button btnReverify;
        private Button btnClose;

        private bool _running;

        private readonly Color Accent =
            Color.FromArgb(239, 91, 18);

        private readonly Color TextPrimary =
            Color.FromArgb(32, 31, 30);

        private readonly Color Border =
            Color.FromArgb(190, 190, 190);

        public PublishStatusForm()
        {
            BuildLayout();

            FormClosing +=
                PublishStatusForm_FormClosing;
        }

        private void BuildLayout()
        {
            Text =
                "Project Operation Connector";

            ClientSize =
                new Size(
                    620,
                    455);

            StartPosition =
                FormStartPosition.CenterScreen;

            FormBorderStyle =
                FormBorderStyle.FixedSingle;

            MaximizeBox =
                false;

            ShowInTaskbar =
                true;

            ShowIcon =
                false;

            BackColor =
                Color.White;

            Font =
                new Font(
                    "Segoe UI",
                    9F);

            picProductLogo =
                new PictureBox
                {
                    Left = 20,
                    Top = 14,

                    Width = 48,
                    Height = 48,

                    SizeMode =
                        PictureBoxSizeMode.Zoom,

                    Image =
                        LoadImageResource(
                            "PO_logo")
                };

            Controls.Add(
                picProductLogo);

            Label lblProduct =
                new Label
                {
                    Text =
                        "Project Operation Connector",

                    Left = 82,
                    Top = 23,

                    Width = 450,
                    Height = 34,

                    ForeColor =
                        TextPrimary,

                    Font =
                        new Font(
                            "Segoe UI",
                            17F)
                };

            Controls.Add(
                lblProduct);

            Panel divider =
                new Panel
                {
                    Left = 20,
                    Top = 76,

                    Width = 580,
                    Height = 1,

                    BackColor =
                        Color.FromArgb(
                            220,
                            220,
                            220)
                };

            Controls.Add(
                divider);

            lblTitle =
                new Label
                {
                    Text = "Publish",

                    Left = 22,
                    Top = 91,

                    Width = 576,
                    Height = 30,

                    ForeColor =
                        TextPrimary,

                    Font =
                        new Font(
                            "Segoe UI",
                            14F,
                            FontStyle.Bold)
                };

            Controls.Add(
                lblTitle);

            statusPanel =
                new Panel
                {
                    Left = 22,
                    Top = 127,

                    Width = 576,
                    Height = 70,

                    BackColor =
                        Color.FromArgb(
                            243,
                            242,
                            241),

                    BorderStyle =
                        BorderStyle.FixedSingle
                };

            lblStatus =
                new Label
                {
                    Dock =
                        DockStyle.Fill,

                    Padding =
                        new Padding(12),

                    Text = "Ready",

                    ForeColor =
                        TextPrimary,

                    Font =
                        new Font(
                            "Segoe UI",
                            9.5F),

                    TextAlign =
                        ContentAlignment.MiddleLeft
                };

            statusPanel.Controls.Add(
                lblStatus);

            Controls.Add(
                statusPanel);

            progressBar =
                new ProgressBar
                {
                    Left = 22,
                    Top = 207,

                    Width = 576,
                    Height = 7
                };

            Controls.Add(
                progressBar);

            lblDetails =
                new Label
                {
                    Text = "Details",

                    Left = 22,
                    Top = 228,

                    Width = 100,
                    Height = 20,

                    ForeColor =
                        TextPrimary,

                    Font =
                        new Font(
                            "Segoe UI",
                            9F,
                            FontStyle.Bold)
                };

            Controls.Add(
                lblDetails);

            txtLog =
                new TextBox
                {
                    Left = 22,
                    Top = 250,

                    Width = 576,
                    Height = 135,

                    Multiline =
                        true,

                    ReadOnly =
                        true,

                    ScrollBars =
                        ScrollBars.Vertical,

                    WordWrap =
                        true,

                    BackColor =
                        Color.White,

                    BorderStyle =
                        BorderStyle.FixedSingle,

                    Font =
                        new Font(
                            "Consolas",
                            8.5F)
                };

            Controls.Add(
                txtLog);

            picCompanyLogo =
                new PictureBox
                {
                    Left = 22,
                    Top = 397,

                    Width = 190,
                    Height = 40,

                    SizeMode =
                        PictureBoxSizeMode.Zoom,

                    Image =
                        LoadImageResource(
                            "TuruncuTeknolojiLogo")
                };

            Controls.Add(
                picCompanyLogo);

            btnReverify =
                new Button
                {
                    Text = "Re-verify",

                    Width = 105,
                    Height = 31,

                    Left = 369,
                    Top = 401,

                    Visible =
                        false,

                    Enabled =
                        false,

                    FlatStyle =
                        FlatStyle.Flat,

                    BackColor =
                        Accent,

                    ForeColor =
                        Color.White,

                    Font =
                        new Font(
                            "Segoe UI",
                            9F,
                            FontStyle.Bold)
                };

            btnReverify.FlatAppearance.BorderSize =
                0;

            btnReverify.Click +=
                btnReverify_Click;

            Controls.Add(
                btnReverify);

            btnClose =
                new Button
                {
                    Text = "Close",

                    Width = 105,
                    Height = 31,

                    Left = 493,
                    Top = 401,

                    FlatStyle =
                        FlatStyle.Flat,

                    BackColor =
                        Color.White,

                    ForeColor =
                        TextPrimary
                };

            btnClose.FlatAppearance.BorderColor =
                Border;

            btnClose.Click +=
                delegate
                {
                    if (!_running)
                    {
                        Hide();
                    }
                };

            Controls.Add(
                btnClose);
        }

        private async void btnReverify_Click(
            object sender,
            EventArgs e)
        {
            if (_running)
                return;

            try
            {
                if (Globals.ThisAddIn == null ||
                    Globals.ThisAddIn.Crm == null ||
                    !Globals.ThisAddIn
                        .SelectedCrmProjectId
                        .HasValue)
                {
                    FinishError(
                        "Aktif CRM projesi bulunamadı.");

                    return;
                }

                Guid projectId =
                    Globals.ThisAddIn
                        .SelectedCrmProjectId
                        .Value;

                if (!Globals.ThisAddIn.Crm
                    .HasPendingPublish(
                        projectId))
                {
                    FinishWarning(
                        "Doğrulanmayı bekleyen Publish bulunamadı.");

                    return;
                }

                BeginReverify();

                CrmPublishResult result =
                    await Globals.ThisAddIn.Crm
                        .ReverifyPendingPublishAsync(
                            projectId,
                            Report);

                if (result == null)
                {
                    FinishError(
                        "Re-verify sonucu alınamadı.");

                    return;
                }

                if (result.Verified)
                {
                    FinishSuccess(
                        "CRM sonucu doğrulandı. Baseline güncellendi.");

                    return;
                }

                if (result.Failed)
                {
                    FinishError(
                        result.ErrorMessage);

                    return;
                }

                if (result.Blocked)
                {
                    FinishBlocked(
                        result.VerificationMessage);

                    return;
                }

                FinishWarning(
                    result.VerificationMessage);
            }
            catch (Exception ex)
            {
                FinishError(
                    ex.Message);
            }
        }

        private void BeginReverify()
        {
            RunOnUiThread(
                delegate
                {
                    _running = true;

                    btnClose.Enabled =
                        false;

                    btnReverify.Enabled =
                        false;

                    lblTitle.Text =
                        "Re-verifying";

                    lblStatus.Text =
                        "CRM sonucu yeniden kontrol ediliyor...";

                    SetStatusVisual(
                        PublishStatusLevel.Info);

                    progressBar.Style =
                        ProgressBarStyle.Marquee;

                    progressBar.MarqueeAnimationSpeed =
                        25;

                    AppendLine(
                        DateTime.Now,
                        "INFO",
                        "Re-verify started.");
                });
        }

        private void PublishStatusForm_FormClosing(
            object sender,
            FormClosingEventArgs e)
        {
            e.Cancel = true;

            if (!_running)
            {
                Hide();
            }
        }

        public void BeginPublish()
        {
            RunOnUiThread(
                delegate
                {
                    _running = true;

                    btnClose.Enabled =
                        false;

                    btnReverify.Enabled =
                        false;

                    btnReverify.Visible =
                        false;

                    lblTitle.Text =
                        "Publishing";

                    lblStatus.Text =
                        "CRM değişiklikleri hazırlanıyor...";

                    SetStatusVisual(
                        PublishStatusLevel.Info);

                    progressBar.Style =
                        ProgressBarStyle.Marquee;

                    progressBar.MarqueeAnimationSpeed =
                        25;

                    txtLog.Clear();

                    AppendLine(
                        DateTime.Now,
                        "INFO",
                        "Publish started.");

                    if (!Visible)
                    {
                        Show();
                    }

                    BringToFront();
                    Activate();
                });
        }

        public void Report(
            PublishProgressInfo info)
        {
            if (info == null)
                return;

            RunOnUiThread(
                delegate
                {
                    lblTitle.Text =
                        string.IsNullOrWhiteSpace(
                            info.Title)
                            ? "Publish"
                            : info.Title;

                    lblStatus.Text =
                        GetSummary(
                            info.Message);

                    SetStatusVisual(
                        info.Level);

                    AppendLine(
                        info.Time ==
                        default(DateTime)
                            ? DateTime.Now
                            : info.Time,

                        info.Level
                            .ToString()
                            .ToUpperInvariant(),

                        (
                            string.IsNullOrWhiteSpace(
                                info.Title)
                                ? ""
                                : info.Title + ": "
                        ) +
                        (
                            info.Message ?? ""
                        ));
                });
        }

        public void FinishSuccess(
            string message)
        {
            Finish(
                PublishStatusLevel.Success,
                "Published ✓",
                message);
        }

        public void FinishNoChanges()
        {
            Finish(
                PublishStatusLevel.Success,
                "Up to date ✓",
                "CRM ve Microsoft Project arasında gönderilecek değişiklik bulunamadı.");
        }

        public void FinishWarning(
            string message)
        {
            Finish(
                PublishStatusLevel.Warning,
                "Publish Warning",
                message);
        }

        public void FinishBlocked(
            string message)
        {
            Finish(
                PublishStatusLevel.Warning,
                "Publish Blocked",
                message);
        }

        public void FinishError(
            string message)
        {
            Finish(
                PublishStatusLevel.Error,
                "Publish Failed",
                message);
        }

        private void Finish(
            PublishStatusLevel level,
            string title,
            string message)
        {
            RunOnUiThread(
                delegate
                {
                    _running = false;

                    btnClose.Enabled =
                        true;

                    progressBar.MarqueeAnimationSpeed =
                        0;

                    progressBar.Style =
                        ProgressBarStyle.Blocks;

                    progressBar.Value =
                        level ==
                        PublishStatusLevel.Success
                            ? 100
                            : 0;

                    lblTitle.Text =
                        title;

                    lblStatus.Text =
                        GetSummary(
                            message);

                    SetStatusVisual(
                        level);

                    AppendLine(
                        DateTime.Now,

                        level
                            .ToString()
                            .ToUpperInvariant(),

                        title +
                        ": " +
                        (
                            message ?? ""
                        ));

                    UpdateReverifyButton();

                    if (!Visible)
                    {
                        Show();
                    }

                    BringToFront();
                    Activate();
                });
        }

        private void UpdateReverifyButton()
        {
            bool show = false;

            try
            {
                if (Globals.ThisAddIn != null &&
                    Globals.ThisAddIn.Crm != null &&
                    Globals.ThisAddIn
                        .SelectedCrmProjectId
                        .HasValue)
                {
                    show =
                        Globals.ThisAddIn.Crm
                            .HasPendingPublish(
                                Globals.ThisAddIn
                                    .SelectedCrmProjectId
                                    .Value);
                }
            }
            catch
            {
                show = false;
            }

            btnReverify.Visible =
                show;

            btnReverify.Enabled =
                show &&
                !_running;
        }

        private static string GetSummary(
            string message)
        {
            if (string.IsNullOrWhiteSpace(
                message))
            {
                return "";
            }

            string normalized =
                message
                    .Replace(
                        "\r\n",
                        "\n")
                    .Replace(
                        "\r",
                        "\n")
                    .Trim();

            int paragraphEnd =
                normalized.IndexOf(
                    "\n\n",
                    StringComparison.Ordinal);

            string summary =
                paragraphEnd >= 0
                    ? normalized.Substring(
                        0,
                        paragraphEnd)
                    : normalized;

            summary =
                summary
                    .Replace(
                        "\n",
                        " ")
                    .Trim();

            if (summary.Length > 240)
            {
                summary =
                    summary.Substring(
                        0,
                        237)
                    +
                    "...";
            }

            return summary;
        }

        private void SetStatusVisual(
            PublishStatusLevel level)
        {
            switch (level)
            {
                case PublishStatusLevel.Success:

                    statusPanel.BackColor =
                        Color.FromArgb(
                            223,
                            246,
                            221);

                    lblStatus.ForeColor =
                        Color.FromArgb(
                            16,
                            124,
                            16);

                    break;

                case PublishStatusLevel.Warning:

                    statusPanel.BackColor =
                        Color.FromArgb(
                            255,
                            244,
                            206);

                    lblStatus.ForeColor =
                        Color.FromArgb(
                            121,
                            78,
                            0);

                    break;

                case PublishStatusLevel.Error:

                    statusPanel.BackColor =
                        Color.FromArgb(
                            253,
                            231,
                            233);

                    lblStatus.ForeColor =
                        Color.FromArgb(
                            164,
                            38,
                            44);

                    break;

                default:

                    statusPanel.BackColor =
                        Color.FromArgb(
                            255,
                            242,
                            232);

                    lblStatus.ForeColor =
                        Accent;

                    break;
            }
        }

        private void AppendLine(
            DateTime time,
            string level,
            string message)
        {
            StringBuilder line =
                new StringBuilder();

            line.Append(
                time.ToString(
                    "HH:mm:ss"));

            line.Append(
                "  ");

            line.Append(
                (
                    level ?? ""
                ).PadRight(
                    7));

            line.Append(
                "  ");

            line.Append(
                message ?? "");

            line.AppendLine();

            txtLog.AppendText(
                line.ToString());

            txtLog.SelectionStart =
                txtLog.TextLength;

            txtLog.ScrollToCaret();
        }

        private void RunOnUiThread(
            Action action)
        {
            if (IsDisposed ||
                Disposing)
            {
                return;
            }

            if (InvokeRequired)
            {
                try
                {
                    BeginInvoke(
                        action);
                }
                catch
                {
                }

                return;
            }

            action();
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