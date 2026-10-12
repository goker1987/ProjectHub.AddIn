using Microsoft.Identity.Client;
using Microsoft.PowerPlatform.Dataverse.Client;
using Microsoft.Xrm.Sdk;
using Microsoft.Xrm.Sdk.Messages;
using Microsoft.Xrm.Sdk.Metadata;
using Microsoft.Xrm.Sdk.Query;
using ProjectHub.AddIn;
using System;
using System.Collections.Generic;
using System.Globalization;
using System.IO;
using System.Linq;
using System.Net;
using System.Security.Cryptography;
using System.Text;
using System.Threading;

using MSProject = Microsoft.Office.Interop.MSProject;
using Tasks = System.Threading.Tasks;

namespace ProjectHub.AddIn.Services
{
    public class CrmService : IDisposable
    {
        private ServiceClient _client;

        private IPublicClientApplication _msalClient;

        private IAccount _account;

        private string _lastError;

        private System.Exception _lastException;

        private readonly List<CrmProject> _cachedProjects =
            new List<CrmProject>();

        private readonly Dictionary<Guid, string> _loadedCrmBaselines =
            new Dictionary<Guid, string>();

        private readonly Dictionary<Guid, ProjectSnapshot> _loadedProjectBaselines =
            new Dictionary<Guid, ProjectSnapshot>();

        private readonly Dictionary<Guid, PendingPublishState> _pendingPublishes =
            new Dictionary<Guid, PendingPublishState>();

        private readonly Dictionary<string, AttributeMetadata> _percentMetadata =
            new Dictionary<string, AttributeMetadata>(
                StringComparer.OrdinalIgnoreCase);

        private readonly object _baselineLock =
            new object();

        private readonly object _metadataLock =
            new object();

        private readonly SemaphoreSlim _publishGate =
            new SemaphoreSlim(1, 1);

        private const string PublicClientId =
            "51f81489-12ee-4a9e-aaae-a2591f45987d";

        private const string RedirectUri =
            "app://58145B91-0C36-4500-8554-080854F2AC97";

        private const string PercentCompletedField =
            "tt_percentcompleted";

        private const int MaxOperationCount =
            200;

        private const int QuickVerificationSeconds =
            8;

        // =========================================================
        // STATE
        // =========================================================

        public bool IsConnected
        {
            get
            {
                return
                    _client != null &&
                    _client.IsReady;
            }
        }

        public bool LastLoginCancelled
        {
            get;
            private set;
        }

        public string LastError
        {
            get
            {
                if (!string.IsNullOrWhiteSpace(
                    _lastError))
                {
                    return _lastError;
                }

                return
                    _client != null
                        ? _client.LastError
                        : null;
            }
        }

        public System.Exception LastException
        {
            get
            {
                if (_lastException != null)
                    return _lastException;

                return
                    _client != null
                        ? _client.LastException
                        : null;
            }
        }

        public string ConnectedUrl
        {
            get;
            private set;
        }

        public List<CrmProject> CachedProjects
        {
            get
            {
                return new List<CrmProject>(
                    _cachedProjects);
            }
        }

        // =========================================================
        // STATUS
        // =========================================================

        private static void Report(
            Action<PublishProgressInfo> progress,
            PublishStatusLevel level,
            string title,
            string message)
        {
            if (progress == null)
                return;

            try
            {
                progress(
                    new PublishProgressInfo
                    {
                        Time =
                            DateTime.Now,

                        Level =
                            level,

                        Title =
                            title,

                        Message =
                            message
                    });
            }
            catch
            {
            }
        }

        // =========================================================
        // PENDING
        // =========================================================

        public bool HasPendingPublish(
            Guid projectId)
        {
            lock (_baselineLock)
            {
                return
                    _pendingPublishes.ContainsKey(
                        projectId);
            }
        }

        private PendingPublishState GetPendingPublish(
            Guid projectId)
        {
            lock (_baselineLock)
            {
                PendingPublishState value;

                if (_pendingPublishes.TryGetValue(
                    projectId,
                    out value))
                {
                    return value;
                }

                return null;
            }
        }

        private void StorePendingPublish(
            Guid projectId,
            Guid operationSetId,
            ProjectSnapshot snapshot,
            PublishPlan plan)
        {
            lock (_baselineLock)
            {
                _pendingPublishes[
                    projectId] =
                    new PendingPublishState
                    {
                        ProjectId =
                            projectId,

                        OperationSetId =
                            operationSetId,

                        Snapshot =
                            snapshot,

                        Plan =
                            plan,

                        StartedOn =
                            DateTime.Now
                    };
            }
        }

        private void ClearPendingPublish(
            Guid projectId)
        {
            lock (_baselineLock)
            {
                _pendingPublishes.Remove(
                    projectId);
            }
        }

        // =========================================================
        // BASELINE
        // =========================================================

        public void SetLoadedBaseline(
            CrmScheduleData schedule)
        {
            if (schedule == null)
                return;

            string crmFingerprint =
                ComputeScheduleFingerprint(
                    schedule);

            ProjectSnapshot projectSnapshot =
                null;

            try
            {
                projectSnapshot =
                    BuildProjectSnapshot(
                        schedule.ProjectId);
            }
            catch
            {
            }

            lock (_baselineLock)
            {
                _loadedCrmBaselines[
                    schedule.ProjectId] =
                    crmFingerprint;

                if (projectSnapshot != null)
                {
                    _loadedProjectBaselines[
                        schedule.ProjectId] =
                        projectSnapshot;
                }

                _pendingPublishes.Remove(
                    schedule.ProjectId);
            }
        }

        private string GetLoadedCrmBaseline(
            Guid projectId)
        {
            lock (_baselineLock)
            {
                string value;

                if (_loadedCrmBaselines.TryGetValue(
                    projectId,
                    out value))
                {
                    return value;
                }

                return null;
            }
        }

        private ProjectSnapshot GetLoadedProjectBaseline(
            Guid projectId)
        {
            lock (_baselineLock)
            {
                ProjectSnapshot value;

                if (_loadedProjectBaselines.TryGetValue(
                    projectId,
                    out value))
                {
                    return value;
                }

                return null;
            }
        }

        private void SetVerifiedBaselines(
            CrmScheduleData verifiedCrm,
            ProjectSnapshot publishedProject)
        {
            if (verifiedCrm == null ||
                publishedProject == null)
            {
                return;
            }

            string crmFingerprint =
                ComputeScheduleFingerprint(
                    verifiedCrm);

            lock (_baselineLock)
            {
                _loadedCrmBaselines[
                    verifiedCrm.ProjectId] =
                    crmFingerprint;

                _loadedProjectBaselines[
                    verifiedCrm.ProjectId] =
                    publishedProject;

                _pendingPublishes.Remove(
                    verifiedCrm.ProjectId);
            }
        }

        // =========================================================
        // PUBLISHED STATE
        // =========================================================

        public bool IsCurrentProjectPublished(
            Guid projectId)
        {
            try
            {
                if (!IsConnected)
                    return false;

                if (HasPendingPublish(
                    projectId))
                {
                    return false;
                }

                ProjectSnapshot baseline =
                    GetLoadedProjectBaseline(
                        projectId);

                if (baseline == null)
                    return false;

                string baselineFingerprint =
                    ComputeProjectFingerprint(
                        baseline);

                string currentFingerprint =
                    GetCurrentProjectFingerprint(
                        projectId);

                return
                    string.Equals(
                        baselineFingerprint,
                        currentFingerprint,
                        StringComparison.Ordinal);
            }
            catch
            {
                return false;
            }
        }

        // =========================================================
        // CONNECTION
        // =========================================================

        public async Tasks.Task<bool> ConnectAsync(
            string crmUrl)
        {
            LastLoginCancelled =
                false;

            _lastError =
                null;

            _lastException =
                null;

            if (string.IsNullOrWhiteSpace(
                crmUrl))
            {
                throw new ArgumentException(
                    "CRM URL boş olamaz.");
            }

            string normalizedUrl =
                NormalizeUrl(
                    crmUrl);

            Uri serviceUri;

            if (!Uri.TryCreate(
                normalizedUrl,
                UriKind.Absolute,
                out serviceUri))
            {
                throw new ArgumentException(
                    "Geçerli bir CRM URL girilmedi.");
            }

            ServicePointManager.SecurityProtocol =
                SecurityProtocolType.Tls12;

            ServiceClient newClient =
                null;

            try
            {
                IPublicClientApplication newMsalClient =
                    PublicClientApplicationBuilder
                        .Create(
                            PublicClientId)
                        .WithAuthority(
                            AadAuthorityAudience.AzureAdMultipleOrgs)
                        .WithRedirectUri(
                            RedirectUri)
                        .Build();

                string resource =
                    serviceUri
                        .GetLeftPart(
                            UriPartial.Authority)
                        .TrimEnd('/');

                string[] scopes =
                {
                    resource +
                    "/user_impersonation"
                };

                AuthenticationResult authResult;

                try
                {
                    authResult =
                        await newMsalClient
                            .AcquireTokenInteractive(
                                scopes)
                            .WithUseEmbeddedWebView(
                                true)
                            .WithPrompt(
                                Prompt.SelectAccount)
                            .ExecuteAsync();
                }
                catch (MsalClientException ex)
                {
                    if (IsLoginCancelled(
                        ex))
                    {
                        LastLoginCancelled =
                            true;

                        return false;
                    }

                    throw;
                }
                catch (OperationCanceledException)
                {
                    LastLoginCancelled =
                        true;

                    return false;
                }

                IAccount newAccount =
                    authResult.Account;

                if (newAccount == null)
                {
                    throw new InvalidOperationException(
                        "Microsoft hesabı alınamadı.");
                }

                Func<string, Tasks.Task<string>> tokenProvider =
                    async instanceUri =>
                        await GetAccessTokenSilentAsync(
                            newMsalClient,
                            newAccount,
                            instanceUri);

                newClient =
                    await Tasks.Task.Run(
                        () =>
                            new ServiceClient(
                                serviceUri,
                                tokenProvider,
                                true,
                                null));

                if (newClient == null ||
                    !newClient.IsReady)
                {
                    throw new InvalidOperationException(
                        newClient != null &&
                        !string.IsNullOrWhiteSpace(
                            newClient.LastError)
                            ? newClient.LastError
                            : "Dataverse bağlantısı kurulamadı.",
                        newClient != null
                            ? newClient.LastException
                            : null);
                }

                ServiceClient oldClient =
                    _client;

                _client =
                    newClient;

                newClient =
                    null;

                _msalClient =
                    newMsalClient;

                _account =
                    newAccount;

                ConnectedUrl =
                    resource;

                _cachedProjects.Clear();

                lock (_baselineLock)
                {
                    _loadedCrmBaselines.Clear();
                    _loadedProjectBaselines.Clear();
                    _pendingPublishes.Clear();
                }

                lock (_metadataLock)
                {
                    _percentMetadata.Clear();
                }

                _lastError =
                    null;

                _lastException =
                    null;

                LastLoginCancelled =
                    false;

                if (oldClient != null)
                {
                    try
                    {
                        oldClient.Dispose();
                    }
                    catch
                    {
                    }
                }

                return true;
            }
            catch (MsalClientException ex)
            {
                if (IsLoginCancelled(
                    ex))
                {
                    LastLoginCancelled =
                        true;

                    return false;
                }

                _lastException =
                    ex;

                _lastError =
                    ex.Message;

                return false;
            }
            catch (OperationCanceledException)
            {
                LastLoginCancelled =
                    true;

                return false;
            }
            catch (System.Exception ex)
            {
                _lastException =
                    ex;

                _lastError =
                    ex.Message;

                return false;
            }
            finally
            {
                if (newClient != null)
                {
                    try
                    {
                        newClient.Dispose();
                    }
                    catch
                    {
                    }
                }
            }
        }

        private static bool IsLoginCancelled(
            MsalClientException ex)
        {
            return
                ex != null &&
                (
                    string.Equals(
                        ex.ErrorCode,
                        "authentication_canceled",
                        StringComparison.OrdinalIgnoreCase)
                    ||
                    string.Equals(
                        ex.ErrorCode,
                        "user_canceled",
                        StringComparison.OrdinalIgnoreCase)
                );
        }

        private async Tasks.Task<string> GetAccessTokenSilentAsync(
            IPublicClientApplication client,
            IAccount account,
            string instanceUri)
        {
            string resource =
                NormalizeUrl(
                    instanceUri)
                .TrimEnd('/');

            string[] scopes =
            {
                resource +
                "/user_impersonation"
            };

            try
            {
                AuthenticationResult result =
                    await client
                        .AcquireTokenSilent(
                            scopes,
                            account)
                        .ExecuteAsync();

                return result.AccessToken;
            }
            catch (MsalUiRequiredException ex)
            {
                throw new InvalidOperationException(
                    "Microsoft oturumunun yenilenmesi gerekiyor. " +
                    "Lütfen yeniden Connect yapın.",
                    ex);
            }
        }

        // =========================================================
        // PROJECT LIST
        // =========================================================

        public async System.Threading.Tasks.Task<List<CrmProject>> GetProjectsAsync()
        {
            EnsureConnected();

            QueryExpression query =
                new QueryExpression(
                    "msdyn_project");

            // ColumnSet(true) kullanıyoruz.
            // Böylece farklı Project Operations sürümlerindeki
            // start / finish alanı farklılıklarında query patlamıyor.
            query.ColumnSet =
                new ColumnSet(true);

            query.AddOrder(
                "msdyn_subject",
                OrderType.Ascending);

            EntityCollection result =
                await _client
                    .RetrieveMultipleAsync(
                        query);

            _cachedProjects.Clear();

            foreach (Entity entity
                     in result.Entities)
            {
                EntityReference managerReference =
                    entity.GetAttributeValue<EntityReference>(
                        "msdyn_projectmanager");

                string managerName =
                    managerReference != null
                        ? managerReference.Name
                        : "";

                if (string.IsNullOrWhiteSpace(
                        managerName) &&
                    entity.FormattedValues.Contains(
                        "msdyn_projectmanager"))
                {
                    managerName =
                        entity.FormattedValues[
                            "msdyn_projectmanager"];
                }

                DateTime? start =
                    GetDateTime(
                        entity,
                        "msdyn_start")
                    ??
                    GetDateTime(
                        entity,
                        "msdyn_scheduledstart");

                DateTime? finish =
                    GetDateTime(
                        entity,
                        "msdyn_finish")
                    ??
                    GetDateTime(
                        entity,
                        "msdyn_scheduledend");

                int? percent =
                    GetPercentValue(
                        entity,
                        "tt_percentcompleted");

                _cachedProjects.Add(
                    new CrmProject
                    {
                        Id =
                            entity.Id,

                        Name =
                            entity.GetAttributeValue<string>(
                                "msdyn_subject")
                            ??
                            "(İsimsiz Proje)",

                        ProjectManager =
                            managerName ?? "",

                        Start =
                            start,

                        Finish =
                            finish,

                        PercentComplete =
                            percent
                    });
            }

            return
                new List<CrmProject>(
                    _cachedProjects);
        }

        // =========================================================
        // SCHEDULE READ
        // =========================================================

        public async Tasks.Task<CrmScheduleData> GetScheduleDataAsync(
            Guid projectId)
        {
            EnsureConnected();

            return await Tasks.Task.Run(
                () =>
                    GetScheduleDataSync(
                        projectId));
        }

        private CrmScheduleData GetScheduleDataSync(
            Guid projectId)
        {
            Entity projectEntity =
                _client.Retrieve(
                    "msdyn_project",
                    projectId,
                    new ColumnSet(true));

            EntityCollection tasks =
                _client.RetrieveMultiple(
                    BuildProjectQuery(
                        "msdyn_projecttask",
                        "msdyn_project",
                        projectId));

            EntityCollection dependencies =
                _client.RetrieveMultiple(
                    BuildProjectQuery(
                        "msdyn_projecttaskdependency",
                        "msdyn_project",
                        projectId));

            EntityCollection assignments =
                _client.RetrieveMultiple(
                    BuildProjectQuery(
                        "msdyn_resourceassignment",
                        "msdyn_projectid",
                        projectId));

            EntityCollection team =
                _client.RetrieveMultiple(
                    BuildProjectQuery(
                        "msdyn_projectteam",
                        "msdyn_project",
                        projectId));

            EntityCollection buckets =
                _client.RetrieveMultiple(
                    BuildProjectQuery(
                        "msdyn_projectbucket",
                        "msdyn_project",
                        projectId));

            HashSet<Guid> resourceIds =
                new HashSet<Guid>();

            foreach (Entity member
                     in team.Entities)
            {
                EntityReference resource =
                    member.GetAttributeValue<EntityReference>(
                        "msdyn_bookableresourceid");

                if (resource != null)
                {
                    resourceIds.Add(
                        resource.Id);
                }
            }

            EntityCollection resources =
                new EntityCollection();

            if (resourceIds.Count > 0)
            {
                QueryExpression query =
                    new QueryExpression(
                        "bookableresource");

                query.ColumnSet =
                    new ColumnSet(
                        "bookableresourceid",
                        "name");

                query.Criteria.AddCondition(
                    "bookableresourceid",
                    ConditionOperator.In,
                    resourceIds
                        .Select(
                            x => (object)x)
                        .ToArray());

                resources =
                    _client.RetrieveMultiple(
                        query);
            }

            return
                new CrmScheduleData
                {
                    ProjectId =
                        projectId,

                    Project =
                        projectEntity,

                    Tasks =
                        tasks,

                    Dependencies =
                        dependencies,

                    ResourceAssignments =
                        assignments,

                    TeamMembers =
                        team,

                    BookableResources =
                        resources,

                    Buckets =
                        buckets
                };
        }

        private QueryExpression BuildProjectQuery(
            string entityName,
            string projectField,
            Guid projectId)
        {
            QueryExpression query =
                new QueryExpression(
                    entityName);

            query.ColumnSet =
                new ColumnSet(true);

            query.Criteria.AddCondition(
                projectField,
                ConditionOperator.Equal,
                projectId);

            if (entityName ==
                "msdyn_projecttask")
            {
                query.AddOrder(
                    "msdyn_displaysequence",
                    OrderType.Ascending);
            }

            return query;
        }

        // =========================================================
        // FINGERPRINT
        // =========================================================

        public string GetCurrentProjectFingerprint(
            Guid projectId)
        {
            ProjectSnapshot snapshot =
                BuildProjectSnapshot(
                    projectId);

            return
                ComputeProjectFingerprint(
                    snapshot);
        }

        // =========================================================
        // REVERIFY
        // =========================================================

        public async Tasks.Task<CrmPublishResult> ReverifyPendingPublishAsync(
            Guid projectId,
            Action<PublishProgressInfo> progress = null)
        {
            bool lockTaken =
                await _publishGate
                    .WaitAsync(0);

            if (!lockTaken)
            {
                return BlockedResult(
                    "Başka bir Publish işlemi devam ediyor.",
                    progress);
            }

            try
            {
                EnsureConnected();

                PendingPublishState pending =
                    GetPendingPublish(
                        projectId);

                if (pending == null)
                {
                    return BlockedResult(
                        "Doğrulanmayı bekleyen bir Publish bulunamadı.",
                        progress);
                }

                Report(
                    progress,
                    PublishStatusLevel.Info,
                    "Re-verifying",
                    "Önceki Publish CRM üzerinde yeniden kontrol ediliyor.");

                OperationSetState state =
                    await Tasks.Task.Run(
                        () =>
                            TryGetOperationSetState(
                                pending.OperationSetId));

                if (state != null &&
                    state.ErrorLog != null)
                {
                    string error =
                        await Tasks.Task.Run(
                            () =>
                                ReadPssError(
                                    state.ErrorLog));

                    ClearPendingPublish(
                        projectId);

                    return FailedResult(
                        pending.OperationSetId,
                        ComputeProjectFingerprint(
                            pending.Snapshot),
                        "Önceki Publish CRM scheduling service tarafından başarısız oldu." +
                        "\r\n\r\n" +
                        error,
                        progress);
                }

                CrmScheduleData currentCrm;

                try
                {
                    currentCrm =
                        await GetScheduleDataAsync(
                            projectId);
                }
                catch (System.Exception ex)
                {
                    return FailedResult(
                        pending.OperationSetId,
                        ComputeProjectFingerprint(
                            pending.Snapshot),
                        "CRM read-back yapılamadı." +
                        "\r\n\r\n" +
                        GetServiceExceptionMessage(
                            ex),
                        progress);
                }

                string mismatch;

                if (TryVerifyPublishedChanges(
                    pending.Snapshot,
                    currentCrm,
                    pending.Plan,
                    out mismatch))
                {
                    SetVerifiedBaselines(
                        currentCrm,
                        pending.Snapshot);

                    CrmPublishResult success =
                        BuildPendingResult(
                            pending,
                            true,
                            null);

                    success.VerifiedOn =
                        DateTime.Now;

                    Report(
                        progress,
                        PublishStatusLevel.Success,
                        "Verified",
                        "Önceki Publish CRM üzerinde doğrulandı. Baseline güncellendi.");

                    WriteAuditSuccess(
                        projectId,
                        success);

                    return success;
                }

                string stateText =
                    state != null &&
                    !string.IsNullOrWhiteSpace(
                        state.StatusText)
                        ? state.StatusText
                        : "Unknown";

                CrmPublishResult warning =
                    BuildPendingResult(
                        pending,
                        false,

                        "Publish henüz doğrulanamadı." +
                        "\r\n\r\n" +
                        "Son kontrol:\r\n" +
                        (
                            string.IsNullOrWhiteSpace(
                                mismatch)
                                ? "(bilinmiyor)"
                                : mismatch
                        ) +
                        "\r\n\r\n" +
                        "Operation Set Status: " +
                        stateText +
                        "\r\nOperation Set: " +
                        pending.OperationSetId +
                        "\r\n\r\nBaseline değiştirilmedi.");

                Report(
                    progress,
                    PublishStatusLevel.Warning,
                    "Publish Pending",
                    warning.VerificationMessage);

                WriteAuditWarning(
                    projectId,
                    warning);

                return warning;
            }
            catch (System.Exception ex)
            {
                return FailedResult(
                    Guid.Empty,
                    null,
                    GetServiceExceptionMessage(
                        ex),
                    progress);
            }
            finally
            {
                _publishGate.Release();
            }
        }

        private CrmPublishResult BuildPendingResult(
            PendingPublishState pending,
            bool verified,
            string message)
        {
            return
                new CrmPublishResult
                {
                    OperationSetId =
                        pending.OperationSetId,

                    UpdatedProjectFields =
                        pending.Plan.ProjectPercentChanged
                            ? 1
                            : 0,

                    CreatedTasks =
                        pending.Plan.TaskCreates.Count,

                    UpdatedTasks =
                        pending.Plan.TaskUpdates.Count,

                    DeletedTasks =
                        pending.Plan.TaskDeletes.Count,

                    CreatedDependencies =
                        pending.Plan.DependencyCreates.Count,

                    DeletedDependencies =
                        pending.Plan.DependencyDeletes.Count,

                    UnchangedDependencies =
                        pending.Plan.UnchangedDependencies,

                    CreatedAssignments =
                        pending.Plan.AssignmentCreates.Count,

                    DeletedAssignments =
                        pending.Plan.AssignmentDeletes.Count,

                    UnchangedAssignments =
                        pending.Plan.UnchangedAssignments,

                    Verified =
                        verified,

                    VerificationWarning =
                        !verified,

                    VerificationMessage =
                        message,

                    SnapshotTakenOn =
                        pending.Snapshot.SnapshotTakenOn,

                    PublishedProjectFingerprint =
                        ComputeProjectFingerprint(
                            pending.Snapshot),

                    AuditLogPath =
                        GetAuditLogPath()
                };
        }

        // =========================================================
        // PUBLISH
        // =========================================================

        public async Tasks.Task<CrmPublishResult> PublishProjectAsync(
            Guid projectId,
            Action<PublishProgressInfo> progress = null)
        {
            bool lockTaken =
                await _publishGate
                    .WaitAsync(0);

            if (!lockTaken)
            {
                return BlockedResult(
                    "Bir Publish işlemi zaten devam ediyor.",
                    progress);
            }

            string snapshotFingerprint =
                null;

            try
            {
                EnsureConnected();

                Report(
                    progress,
                    PublishStatusLevel.Info,
                    "Preflight",
                    "Microsoft Project değişiklikleri hazırlanıyor.");

                string crmBaseline =
                    GetLoadedCrmBaseline(
                        projectId);

                ProjectSnapshot projectBaseline =
                    GetLoadedProjectBaseline(
                        projectId);

                if (string.IsNullOrWhiteSpace(
                        crmBaseline) ||
                    projectBaseline == null)
                {
                    return BlockedResult(
                        "Synchronization baseline bulunamadı." +
                        "\r\n\r\nProjeyi Open Project ile yeniden yükleyin.",
                        progress);
                }

                ProjectSnapshot currentProject =
                    BuildProjectSnapshot(
                        projectId);

                snapshotFingerprint =
                    ComputeProjectFingerprint(
                        currentProject);

                CrmScheduleData currentCrm;

                try
                {
                    currentCrm =
                        await GetScheduleDataAsync(
                            projectId);
                }
                catch (System.Exception ex)
                {
                    return FailedResult(
                        Guid.Empty,
                        snapshotFingerprint,
                        "CRM verisi okunamadı." +
                        "\r\n\r\n" +
                        GetServiceExceptionMessage(
                            ex),
                        progress);
                }

                string currentCrmFingerprint =
                    ComputeScheduleFingerprint(
                        currentCrm);

                PendingPublishState pending =
                    GetPendingPublish(
                        projectId);

                if (pending != null)
                {
                    string pendingMismatch;

                    if (TryVerifyPublishedChanges(
                        pending.Snapshot,
                        currentCrm,
                        pending.Plan,
                        out pendingMismatch))
                    {
                        SetVerifiedBaselines(
                            currentCrm,
                            pending.Snapshot);

                        crmBaseline =
                            currentCrmFingerprint;

                        projectBaseline =
                            pending.Snapshot;
                    }
                    else
                    {
                        OperationSetState pendingState =
                            TryGetOperationSetState(
                                pending.OperationSetId);

                        if (pendingState != null &&
                            pendingState.ErrorLog != null)
                        {
                            string pssError =
                                ReadPssError(
                                    pendingState.ErrorLog);

                            ClearPendingPublish(
                                projectId);

                            return FailedResult(
                                pending.OperationSetId,
                                snapshotFingerprint,
                                "Önceki Publish başarısız oldu." +
                                "\r\n\r\n" +
                                pssError,
                                progress);
                        }

                        return BlockedResult(
                            "Önceki Publish henüz doğrulanmadı." +
                            "\r\n\r\n" +
                            (
                                string.IsNullOrWhiteSpace(
                                    pendingMismatch)
                                    ? ""
                                    : "Son kontrol:\r\n" +
                                      pendingMismatch +
                                      "\r\n\r\n"
                            ) +
                            "Operation Set: " +
                            pending.OperationSetId +
                            "\r\n\r\nPublish penceresindeki Re-verify butonunu kullanın.",
                            progress);
                    }
                }

                if (!string.Equals(
                    crmBaseline,
                    currentCrmFingerprint,
                    StringComparison.Ordinal))
                {
                    return BlockedResult(
                        "CRM tarafında bu ProjectHub oturumu dışında yeni değişiklik bulundu." +
                        "\r\n\r\nVeri kaybını önlemek için Publish durduruldu." +
                        "\r\n\r\nProjeyi Open Project ile yeniden yükleyin.",
                        progress);
                }

                string baselineProjectFingerprint =
                    ComputeProjectFingerprint(
                        projectBaseline);

                if (string.Equals(
                    baselineProjectFingerprint,
                    snapshotFingerprint,
                    StringComparison.Ordinal))
                {
                    CrmPublishResult noChanges =
                        new CrmPublishResult
                        {
                            NoChanges =
                                true,

                            Verified =
                                true,

                            SnapshotTakenOn =
                                currentProject.SnapshotTakenOn,

                            VerifiedOn =
                                DateTime.Now,

                            PublishedProjectFingerprint =
                                snapshotFingerprint,

                            AuditLogPath =
                                GetAuditLogPath()
                        };

                    Report(
                        progress,
                        PublishStatusLevel.Success,
                        "Up to date",
                        "Gönderilecek değişiklik bulunamadı.");

                    WriteAuditSuccess(
                        projectId,
                        noChanges);

                    return noChanges;
                }

                PublishPlan plan =
                    BuildPublishPlan(
                        projectBaseline,
                        currentProject,
                        currentCrm);

                string validationError =
                    ValidatePublishPlan(
                        plan,
                        currentProject,
                        currentCrm);

                if (!string.IsNullOrWhiteSpace(
                    validationError))
                {
                    return BlockedResult(
                        validationError,
                        progress);
                }

                if (plan.OperationCount == 0)
                {
                    SetVerifiedBaselines(
                        currentCrm,
                        currentProject);

                    return
                        new CrmPublishResult
                        {
                            NoChanges =
                                true,

                            Verified =
                                true,

                            SnapshotTakenOn =
                                currentProject.SnapshotTakenOn,

                            VerifiedOn =
                                DateTime.Now,

                            PublishedProjectFingerprint =
                                snapshotFingerprint,

                            AuditLogPath =
                                GetAuditLogPath()
                        };
                }

                Report(
                    progress,
                    PublishStatusLevel.Info,
                    "Publishing",
                    BuildPlanDescription(
                        plan));

                CrmPublishResult result =
                    await Tasks.Task.Run(
                        () =>
                            ExecutePublishPlan(
                                projectId,
                                currentProject,
                                currentCrm,
                                plan,
                                snapshotFingerprint,
                                currentCrmFingerprint,
                                progress));

                if (result.Verified)
                {
                    WriteAuditSuccess(
                        projectId,
                        result);
                }
                else
                {
                    WriteAuditWarning(
                        projectId,
                        result);
                }

                return result;
            }
            catch (System.Exception ex)
            {
                WriteAuditFailure(
                    projectId,
                    snapshotFingerprint,
                    ex);

                return FailedResult(
                    Guid.Empty,
                    snapshotFingerprint,
                    GetServiceExceptionMessage(
                        ex),
                    progress);
            }
            finally
            {
                _publishGate.Release();
            }
        }

        // =========================================================
        // RESULTS
        // =========================================================

        private CrmPublishResult BlockedResult(
            string message,
            Action<PublishProgressInfo> progress)
        {
            Report(
                progress,
                PublishStatusLevel.Warning,
                "Publish Blocked",
                message);

            return
                new CrmPublishResult
                {
                    Blocked =
                        true,

                    VerificationMessage =
                        message,

                    AuditLogPath =
                        GetAuditLogPath()
                };
        }

        private CrmPublishResult FailedResult(
            Guid operationSetId,
            string fingerprint,
            string message,
            Action<PublishProgressInfo> progress)
        {
            Report(
                progress,
                PublishStatusLevel.Error,
                "Publish Failed",
                message);

            return
                new CrmPublishResult
                {
                    OperationSetId =
                        operationSetId,

                    Failed =
                        true,

                    ErrorMessage =
                        message,

                    PublishedProjectFingerprint =
                        fingerprint,

                    AuditLogPath =
                        GetAuditLogPath()
                };
        }

        private static string BuildPlanDescription(
            PublishPlan plan)
        {
            return
                "Project Progress: " +
                (
                    plan.ProjectPercentChanged
                        ? "update"
                        : "unchanged"
                ) +
                "\r\n" +

                "Tasks: " +
                plan.TaskCreates.Count +
                " new, " +
                plan.TaskUpdates.Count +
                " update, " +
                plan.TaskDeletes.Count +
                " delete" +
                "\r\n" +

                "Predecessors: " +
                plan.DependencyCreates.Count +
                " new, " +
                plan.DependencyDeletes.Count +
                " delete" +
                "\r\n" +

                "Resources: " +
                plan.AssignmentCreates.Count +
                " assign, " +
                plan.AssignmentDeletes.Count +
                " unassign";
        }

        // =========================================================
        // SNAPSHOT
        // =========================================================

        private ProjectSnapshot BuildProjectSnapshot(
            Guid projectId)
        {
            MSProject.Project project =
                Globals.ThisAddIn
                    .Application
                    .ActiveProject;

            if (project == null)
            {
                throw new InvalidOperationException(
                    "Aktif Microsoft Project bulunamadı.");
            }

            double hoursPerDay =
                project.HoursPerDay;

            if (hoursPerDay <= 0)
            {
                hoursPerDay =
                    8.0;
            }

            ProjectSnapshot snapshot =
                new ProjectSnapshot
                {
                    ProjectId =
                        projectId,

                    SnapshotTakenOn =
                        DateTime.Now,

                    HoursPerDay =
                        hoursPerDay,

                    ProjectPercentComplete =
                        GetProjectSummaryPercent(
                            project)
                };

            foreach (MSProject.Task task
                     in project.Tasks)
            {
                if (task == null ||
                    string.IsNullOrWhiteSpace(
                        task.Name))
                {
                    continue;
                }

                Guid taskId;

                if (!Guid.TryParse(
                    task.Text30,
                    out taskId))
                {
                    taskId =
                        Guid.NewGuid();

                    task.Text30 =
                        taskId.ToString();
                }

                snapshot.TaskIds[
                    task.ID] =
                    taskId;
            }

            foreach (MSProject.Task task
                     in project.Tasks)
            {
                if (task == null ||
                    string.IsNullOrWhiteSpace(
                        task.Name))
                {
                    continue;
                }

                Guid taskId;

                if (!snapshot.TaskIds.TryGetValue(
                    task.ID,
                    out taskId))
                {
                    continue;
                }

                TaskSnapshot item =
                    new TaskSnapshot
                    {
                        Id =
                            taskId,

                        Name =
                            task.Name ?? "",

                        IsSummary =
                            task.Summary,

                        OutlineLevel =
                            task.OutlineLevel,

                        PercentComplete =
                            GetTaskPercent(
                                task)
                    };

                try
                {
                    item.Start =
                        Convert.ToDateTime(
                            task.Start);
                }
                catch
                {
                }

                try
                {
                    item.Finish =
                        Convert.ToDateTime(
                            task.Finish);
                }
                catch
                {
                }

                try
                {
                    item.DurationMinutes =
                        Convert.ToDouble(
                            task.Duration,
                            CultureInfo.InvariantCulture);

                    item.DurationDays =
                        item.DurationMinutes /
                        (
                            hoursPerDay *
                            60.0
                        );
                }
                catch
                {
                    item.DurationMinutes =
                        0;

                    item.DurationDays =
                        0;
                }

                Guid bucketId;

                if (Guid.TryParse(
                    task.Text29,
                    out bucketId))
                {
                    item.BucketId =
                        bucketId;
                }

                if (task.OutlineLevel > 1)
                {
                    try
                    {
                        MSProject.Task parent =
                            task.OutlineParent;

                        Guid parentId;

                        if (parent != null &&
                            snapshot.TaskIds.TryGetValue(
                                parent.ID,
                                out parentId))
                        {
                            item.ParentId =
                                parentId;
                        }
                    }
                    catch
                    {
                    }
                }

                snapshot.Tasks.Add(
                    item);
            }

            BuildDependencySnapshot(
                project,
                snapshot);

            BuildAssignmentSnapshot(
                project,
                snapshot);

            return snapshot;
        }

        private static int GetProjectSummaryPercent(
            MSProject.Project project)
        {
            try
            {
                if (project.ProjectSummaryTask == null)
                    return 0;

                return ClampPercent(
                    Convert.ToInt32(
                        project.ProjectSummaryTask.PercentComplete,
                        CultureInfo.InvariantCulture));
            }
            catch
            {
                return 0;
            }
        }

        private static int GetTaskPercent(
            MSProject.Task task)
        {
            try
            {
                return ClampPercent(
                    Convert.ToInt32(
                        task.PercentComplete,
                        CultureInfo.InvariantCulture));
            }
            catch
            {
                return 0;
            }
        }

        private static int ClampPercent(
            int value)
        {
            if (value < 0)
                return 0;

            if (value > 100)
                return 100;

            return value;
        }

        private void BuildDependencySnapshot(
            MSProject.Project project,
            ProjectSnapshot snapshot)
        {
            HashSet<string> dependencyKeys =
                new HashSet<string>(
                    StringComparer.OrdinalIgnoreCase);

            foreach (MSProject.Task task
                     in project.Tasks)
            {
                if (task == null)
                    continue;

                Guid successorId;

                if (!snapshot.TaskIds.TryGetValue(
                    task.ID,
                    out successorId))
                {
                    continue;
                }

                try
                {
                    foreach (MSProject.TaskDependency dependency
                             in task.TaskDependencies)
                    {
                        if (dependency == null ||
                            dependency.From == null ||
                            dependency.To == null ||
                            dependency.To.ID != task.ID)
                        {
                            continue;
                        }

                        Guid predecessorId;

                        if (!snapshot.TaskIds.TryGetValue(
                            dependency.From.ID,
                            out predecessorId))
                        {
                            continue;
                        }

                        DependencySnapshot item =
                            new DependencySnapshot
                            {
                                PredecessorId =
                                    predecessorId,

                                SuccessorId =
                                    successorId,

                                LinkType =
                                    ConvertLinkType(
                                        dependency.Type)
                            };

                        if (dependencyKeys.Add(
                            item.Key))
                        {
                            snapshot.Dependencies.Add(
                                item);
                        }
                    }
                }
                catch
                {
                }
            }
        }

        private void BuildAssignmentSnapshot(
            MSProject.Project project,
            ProjectSnapshot snapshot)
        {
            HashSet<string> assignmentKeys =
                new HashSet<string>(
                    StringComparer.OrdinalIgnoreCase);

            foreach (MSProject.Task task
                     in project.Tasks)
            {
                if (task == null ||
                    task.Summary)
                {
                    continue;
                }

                Guid taskId;

                if (!snapshot.TaskIds.TryGetValue(
                    task.ID,
                    out taskId))
                {
                    continue;
                }

                foreach (MSProject.Assignment assignment
                         in task.Assignments)
                {
                    if (assignment == null ||
                        assignment.ResourceID <= 0)
                    {
                        continue;
                    }

                    MSProject.Resource resource =
                        project.Resources[
                            assignment.ResourceID];

                    if (resource == null)
                        continue;

                    Guid teamMemberId;

                    bool mapped =
                        Guid.TryParse(
                            assignment.Text29,
                            out teamMemberId)
                        ||
                        Guid.TryParse(
                            resource.Text29,
                            out teamMemberId);

                    Guid bookableId;

                    Guid? bookableResourceId =
                        Guid.TryParse(
                            resource.Text30,
                            out bookableId)
                            ? (Guid?)bookableId
                            : null;

                    AssignmentSnapshot item =
                        new AssignmentSnapshot
                        {
                            TaskId =
                                taskId,

                            TeamMemberId =
                                mapped
                                    ? (Guid?)teamMemberId
                                    : null,

                            BookableResourceId =
                                bookableResourceId,

                            TaskName =
                                task.Name,

                            ResourceName =
                                resource.Name
                        };

                    if (!item.TeamMemberId.HasValue)
                    {
                        snapshot.Assignments.Add(
                            item);

                        continue;
                    }

                    if (assignmentKeys.Add(
                        item.Key))
                    {
                        snapshot.Assignments.Add(
                            item);
                    }
                }
            }
        }

        // =========================================================
        // FINGERPRINT
        // =========================================================

        private string ComputeProjectFingerprint(
            ProjectSnapshot snapshot)
        {
            List<string> values =
                new List<string>();

            values.Add(
                "P|" +
                snapshot.ProjectPercentComplete);

            foreach (TaskSnapshot task
                     in snapshot.Tasks
                         .OrderBy(
                             x => x.Id))
            {
                StringBuilder value =
                    new StringBuilder();

                value.Append(
                    "T|");

                value.Append(
                    task.Id);

                value.Append(
                    "|");

                value.Append(
                    NormalizeFingerprintText(
                        task.Name));

                value.Append(
                    "|PC=");

                value.Append(
                    task.PercentComplete);

                if (!task.IsSummary)
                {
                    value.Append(
                        "|");

                    value.Append(
                        DateText(
                            task.Start));

                    value.Append(
                        "|");

                    value.Append(
                        DateText(
                            task.Finish));

                    value.Append(
                        "|");

                    value.Append(
                        task.DurationMinutes.ToString(
                            "0.####",
                            CultureInfo.InvariantCulture));
                }

                values.Add(
                    value.ToString());
            }

            foreach (DependencySnapshot dependency
                     in snapshot.Dependencies
                         .OrderBy(
                             x => x.Key))
            {
                values.Add(
                    "D|" +
                    dependency.Key);
            }

            foreach (AssignmentSnapshot assignment
                     in snapshot.Assignments
                         .OrderBy(
                             x => x.Key))
            {
                values.Add(
                    "A|" +
                    assignment.TaskId +
                    "|" +
                    NullableGuidText(
                        assignment.TeamMemberId));
            }

            return ComputeSha256(
                string.Join(
                    "\n",
                    values));
        }

        // =========================================================
        // PLAN
        // =========================================================

        private PublishPlan BuildPublishPlan(
            ProjectSnapshot baseline,
            ProjectSnapshot current,
            CrmScheduleData crm)
        {
            PublishPlan plan =
                new PublishPlan
                {
                    ProjectPercentChanged =
                        baseline.ProjectPercentComplete !=
                        current.ProjectPercentComplete
                };

            Dictionary<Guid, TaskSnapshot> baselineTasks =
                baseline.Tasks
                    .ToDictionary(
                        x => x.Id);

            Dictionary<Guid, TaskSnapshot> currentTasks =
                current.Tasks
                    .ToDictionary(
                        x => x.Id);

            Dictionary<Guid, Entity> crmTasks =
                crm.Tasks.Entities
                    .ToDictionary(
                        x => x.Id);

            foreach (TaskSnapshot currentTask
                     in current.Tasks)
            {
                TaskSnapshot baselineTask;

                if (!baselineTasks.TryGetValue(
                    currentTask.Id,
                    out baselineTask))
                {
                    plan.TaskCreates.Add(
                        currentTask);

                    continue;
                }

                TaskUpdatePlan update =
                    new TaskUpdatePlan
                    {
                        Task =
                            currentTask,

                        NameChanged =
                            !string.Equals(
                                NormalizeText(
                                    baselineTask.Name),
                                NormalizeText(
                                    currentTask.Name),
                                StringComparison.Ordinal),

                        PercentCompleteChanged =
                            baselineTask.PercentComplete !=
                            currentTask.PercentComplete
                    };

                if (!currentTask.IsSummary)
                {
                    update.StartChanged =
                        !SameProjectMinute(
                            baselineTask.Start,
                            currentTask.Start);

                    update.DurationChanged =
                        Math.Abs(
                            baselineTask.DurationMinutes -
                            currentTask.DurationMinutes)
                        > 0.5;

                    update.FinishChanged =
                        !update.StartChanged &&
                        !update.DurationChanged &&
                        !SameProjectMinute(
                            baselineTask.Finish,
                            currentTask.Finish);
                }

                if (update.HasChanges)
                {
                    plan.TaskUpdates.Add(
                        update);
                }
            }

            foreach (TaskSnapshot baselineTask
                     in baseline.Tasks)
            {
                if (currentTasks.ContainsKey(
                    baselineTask.Id))
                {
                    continue;
                }

                if (!crmTasks.ContainsKey(
                    baselineTask.Id))
                {
                    continue;
                }

                plan.TaskDeletes.Add(
                    new TaskDeletePlan
                    {
                        TaskId =
                            baselineTask.Id,

                        TaskName =
                            baselineTask.Name,

                        OutlineLevel =
                            baselineTask.OutlineLevel
                    });
            }

            BuildDependencyPlan(
                baseline,
                current,
                crm,
                plan);

            BuildAssignmentPlan(
                baseline,
                current,
                crm,
                plan);

            return plan;
        }

        private void BuildDependencyPlan(
            ProjectSnapshot baseline,
            ProjectSnapshot current,
            CrmScheduleData crm,
            PublishPlan plan)
        {
            Dictionary<string, DependencySnapshot> baselineDependencies =
                baseline.Dependencies
                    .ToDictionary(
                        x => x.Key,
                        x => x,
                        StringComparer.OrdinalIgnoreCase);

            Dictionary<string, DependencySnapshot> currentDependencies =
                current.Dependencies
                    .ToDictionary(
                        x => x.Key,
                        x => x,
                        StringComparer.OrdinalIgnoreCase);

            Dictionary<string, Entity> crmDependencies =
                new Dictionary<string, Entity>(
                    StringComparer.OrdinalIgnoreCase);

            foreach (Entity entity
                     in crm.Dependencies.Entities)
            {
                string key =
                    GetCrmDependencyKey(
                        entity);

                if (!string.IsNullOrWhiteSpace(
                    key))
                {
                    crmDependencies[
                        key] =
                        entity;
                }
            }

            foreach (DependencySnapshot dependency
                     in current.Dependencies)
            {
                if (!baselineDependencies.ContainsKey(
                    dependency.Key))
                {
                    plan.DependencyCreates.Add(
                        dependency);
                }
                else
                {
                    plan.UnchangedDependencies++;
                }
            }

            foreach (DependencySnapshot dependency
                     in baseline.Dependencies)
            {
                if (currentDependencies.ContainsKey(
                    dependency.Key))
                {
                    continue;
                }

                Entity crmDependency;

                if (crmDependencies.TryGetValue(
                    dependency.Key,
                    out crmDependency))
                {
                    plan.DependencyDeletes.Add(
                        new DependencyDeletePlan
                        {
                            Id =
                                crmDependency.Id,

                            Key =
                                dependency.Key
                        });
                }
            }
        }

        private void BuildAssignmentPlan(
            ProjectSnapshot baseline,
            ProjectSnapshot current,
            CrmScheduleData crm,
            PublishPlan plan)
        {
            Dictionary<string, AssignmentSnapshot> baselineAssignments =
                baseline.Assignments
                    .Where(
                        x =>
                            x.TeamMemberId.HasValue)
                    .ToDictionary(
                        x => x.Key,
                        x => x,
                        StringComparer.OrdinalIgnoreCase);

            Dictionary<string, AssignmentSnapshot> currentAssignments =
                current.Assignments
                    .Where(
                        x =>
                            x.TeamMemberId.HasValue)
                    .ToDictionary(
                        x => x.Key,
                        x => x,
                        StringComparer.OrdinalIgnoreCase);

            Dictionary<string, Entity> crmAssignments =
                BuildCrmAssignmentMap(
                    crm);

            foreach (AssignmentSnapshot assignment
                     in current.Assignments)
            {
                if (!assignment.TeamMemberId.HasValue)
                    continue;

                if (!baselineAssignments.ContainsKey(
                    assignment.Key))
                {
                    plan.AssignmentCreates.Add(
                        assignment);
                }
                else
                {
                    plan.UnchangedAssignments++;
                }
            }

            foreach (AssignmentSnapshot assignment
                     in baseline.Assignments)
            {
                if (!assignment.TeamMemberId.HasValue)
                    continue;

                if (currentAssignments.ContainsKey(
                    assignment.Key))
                {
                    continue;
                }

                Entity crmAssignment;

                if (crmAssignments.TryGetValue(
                    assignment.Key,
                    out crmAssignment))
                {
                    plan.AssignmentDeletes.Add(
                        new AssignmentDeletePlan
                        {
                            Id =
                                crmAssignment.Id,

                            Key =
                                assignment.Key
                        });
                }
            }
        }

        // =========================================================
        // VALIDATE
        // =========================================================

        private string ValidatePublishPlan(
            PublishPlan plan,
            ProjectSnapshot current,
            CrmScheduleData crm)
        {
            try
            {
                GetPercentAttributeMetadata(
                    "msdyn_project");

                GetPercentAttributeMetadata(
                    "msdyn_projecttask");
            }
            catch (System.Exception ex)
            {
                return
                    "tt_percentcompleted alanı doğrulanamadı." +
                    "\r\n\r\n" +
                    ex.Message;
            }

            HashSet<Guid> validTeamMembers =
                new HashSet<Guid>(
                    crm.TeamMembers.Entities
                        .Select(
                            x => x.Id));

            Dictionary<Guid, Guid?> teamBookableMap =
                new Dictionary<Guid, Guid?>();

            foreach (Entity teamMember
                     in crm.TeamMembers.Entities)
            {
                EntityReference resource =
                    teamMember.GetAttributeValue<EntityReference>(
                        "msdyn_bookableresourceid");

                teamBookableMap[
                    teamMember.Id] =
                    resource != null
                        ? (Guid?)resource.Id
                        : null;
            }

            HashSet<string> seenAssignments =
                new HashSet<string>(
                    StringComparer.OrdinalIgnoreCase);

            foreach (AssignmentSnapshot assignment
                     in current.Assignments)
            {
                if (!assignment.TeamMemberId.HasValue)
                {
                    return
                        assignment.TaskName +
                        ": Resource eşleştirilemedi." +
                        "\r\n\r\nResource: " +
                        assignment.ResourceName;
                }

                Guid teamMemberId =
                    assignment.TeamMemberId.Value;

                if (!validTeamMembers.Contains(
                    teamMemberId))
                {
                    return
                        assignment.TaskName +
                        ": Resource mapping geçersiz." +
                        "\r\n\r\nResource: " +
                        assignment.ResourceName;
                }

                Guid? expectedBookable =
                    teamBookableMap[
                        teamMemberId];

                if (assignment.BookableResourceId.HasValue &&
                    expectedBookable.HasValue &&
                    assignment.BookableResourceId.Value !=
                    expectedBookable.Value)
                {
                    return
                        assignment.TaskName +
                        ": Resource mapping tutarsız." +
                        "\r\n\r\nResource: " +
                        assignment.ResourceName;
                }

                if (!seenAssignments.Add(
                    assignment.Key))
                {
                    return
                        assignment.TaskName +
                        ": Aynı resource birden fazla kez atanmış." +
                        "\r\n\r\nResource: " +
                        assignment.ResourceName;
                }
            }

            if (plan.OperationCount >
                MaxOperationCount)
            {
                return
                    "Publish işlem sınırı aşıldı." +
                    "\r\n\r\nİşlem sayısı: " +
                    plan.OperationCount +
                    "\r\nSınır: " +
                    MaxOperationCount;
            }

            return null;
        }

        // =========================================================
        // EXECUTE
        // =========================================================

        private CrmPublishResult ExecutePublishPlan(
            Guid projectId,
            ProjectSnapshot snapshot,
            CrmScheduleData oldCrm,
            PublishPlan plan,
            string snapshotFingerprint,
            string expectedCrmFingerprint,
            Action<PublishProgressInfo> progress)
        {
            CrmScheduleData latestCrm;

            try
            {
                latestCrm =
                    GetScheduleDataSync(
                        projectId);
            }
            catch (System.Exception ex)
            {
                return FailedResult(
                    Guid.Empty,
                    snapshotFingerprint,
                    "Publish öncesi CRM kontrolü yapılamadı." +
                    "\r\n\r\n" +
                    GetServiceExceptionMessage(
                        ex),
                    progress);
            }

            string latestCrmFingerprint =
                ComputeScheduleFingerprint(
                    latestCrm);

            if (!string.Equals(
                expectedCrmFingerprint,
                latestCrmFingerprint,
                StringComparison.Ordinal))
            {
                return BlockedResult(
                    "CRM tarafı Publish hazırlanırken değişti.",
                    progress);
            }

            EntityReference projectReference =
                new EntityReference(
                    "msdyn_project",
                    projectId);

            Guid defaultBucketId;

            try
            {
                defaultBucketId =
                    GetDefaultBucketId(
                        oldCrm);
            }
            catch (System.Exception ex)
            {
                return FailedResult(
                    Guid.Empty,
                    snapshotFingerprint,
                    ex.Message,
                    progress);
            }

            Guid operationSetId;

            string requestError;

            if (!TryCreateOperationSet(
                projectId,
                "ProjectHub Publish " +
                DateTime.Now.ToString(
                    "yyyy-MM-dd HH:mm:ss"),
                out operationSetId,
                out requestError))
            {
                return FailedResult(
                    Guid.Empty,
                    snapshotFingerprint,
                    requestError,
                    progress);
            }

            if (plan.ProjectPercentChanged)
            {
                Entity projectUpdate =
                    new Entity(
                        "msdyn_project",
                        projectId);

                SetNormalizedPercent(
                    projectUpdate,
                    "msdyn_project",
                    snapshot.ProjectPercentComplete);

                if (!TryPssUpdate(
                    projectUpdate,
                    operationSetId,
                    out requestError))
                {
                    TryAbandonOperationSet(
                        operationSetId);

                    return FailedResult(
                        operationSetId,
                        snapshotFingerprint,
                        "Project % Complete güncellenemedi." +
                        "\r\n\r\n" +
                        requestError,
                        progress);
                }
            }

            foreach (DependencyDeletePlan dependency
                     in plan.DependencyDeletes)
            {
                if (!TryPssDelete(
                    dependency.Id,
                    "msdyn_projecttaskdependency",
                    operationSetId,
                    out requestError))
                {
                    TryAbandonOperationSet(
                        operationSetId);

                    return FailedResult(
                        operationSetId,
                        snapshotFingerprint,
                        requestError,
                        progress);
                }
            }

            foreach (AssignmentDeletePlan assignment
                     in plan.AssignmentDeletes)
            {
                if (!TryPssDelete(
                    assignment.Id,
                    "msdyn_resourceassignment",
                    operationSetId,
                    out requestError))
                {
                    TryAbandonOperationSet(
                        operationSetId);

                    return FailedResult(
                        operationSetId,
                        snapshotFingerprint,
                        requestError,
                        progress);
                }
            }

            foreach (TaskDeletePlan task
                     in plan.TaskDeletes
                         .OrderByDescending(
                             x => x.OutlineLevel))
            {
                if (!TryPssDelete(
                    task.TaskId,
                    "msdyn_projecttask",
                    operationSetId,
                    out requestError))
                {
                    TryAbandonOperationSet(
                        operationSetId);

                    return FailedResult(
                        operationSetId,
                        snapshotFingerprint,
                        task.TaskName +
                        ": Task silinemedi." +
                        "\r\n\r\n" +
                        requestError,
                        progress);
                }
            }

            foreach (TaskSnapshot task
                     in plan.TaskCreates
                         .OrderBy(
                             x => x.OutlineLevel))
            {
                Entity entity =
                    CreateTaskEntity(
                        task,
                        projectReference,
                        defaultBucketId);

                if (!TryPssCreate(
                    entity,
                    operationSetId,
                    out requestError))
                {
                    TryAbandonOperationSet(
                        operationSetId);

                    return FailedResult(
                        operationSetId,
                        snapshotFingerprint,
                        task.Name +
                        ": Task oluşturulamadı." +
                        "\r\n\r\n" +
                        requestError,
                        progress);
                }
            }

            foreach (TaskUpdatePlan update
                     in plan.TaskUpdates)
            {
                Entity entity =
                    CreateTaskUpdateEntity(
                        update);

                if (!TryPssUpdate(
                    entity,
                    operationSetId,
                    out requestError))
                {
                    TryAbandonOperationSet(
                        operationSetId);

                    return FailedResult(
                        operationSetId,
                        snapshotFingerprint,
                        update.Task.Name +
                        ": Task güncellenemedi." +
                        "\r\n\r\n" +
                        requestError,
                        progress);
                }
            }

            foreach (DependencySnapshot dependency
                     in plan.DependencyCreates)
            {
                Entity entity =
                    CreateDependencyEntity(
                        dependency,
                        projectReference);

                if (!TryPssCreate(
                    entity,
                    operationSetId,
                    out requestError))
                {
                    TryAbandonOperationSet(
                        operationSetId);

                    return FailedResult(
                        operationSetId,
                        snapshotFingerprint,
                        "Predecessor oluşturulamadı." +
                        "\r\n\r\n" +
                        requestError,
                        progress);
                }
            }

            foreach (AssignmentSnapshot assignment
                     in plan.AssignmentCreates)
            {
                Entity entity =
                    CreateResourceAssignmentEntity(
                        assignment,
                        projectReference);

                if (!TryPssCreate(
                    entity,
                    operationSetId,
                    out requestError))
                {
                    TryAbandonOperationSet(
                        operationSetId);

                    return FailedResult(
                        operationSetId,
                        snapshotFingerprint,
                        assignment.TaskName +
                        ": Resource atanamadı." +
                        "\r\n\r\nResource: " +
                        assignment.ResourceName +
                        "\r\n" +
                        requestError,
                        progress);
                }
            }

            if (!TryExecuteOperationSet(
                operationSetId,
                out requestError))
            {
                TryAbandonOperationSet(
                    operationSetId);

                return FailedResult(
                    operationSetId,
                    snapshotFingerprint,
                    requestError,
                    progress);
            }

            VerificationOutcome verification =
                WaitForVerification(
                    projectId,
                    snapshot,
                    plan,
                    operationSetId,
                    progress);

            CrmPublishResult result =
                new CrmPublishResult
                {
                    OperationSetId =
                        operationSetId,

                    UpdatedProjectFields =
                        plan.ProjectPercentChanged
                            ? 1
                            : 0,

                    CreatedTasks =
                        plan.TaskCreates.Count,

                    UpdatedTasks =
                        plan.TaskUpdates.Count,

                    DeletedTasks =
                        plan.TaskDeletes.Count,

                    CreatedDependencies =
                        plan.DependencyCreates.Count,

                    DeletedDependencies =
                        plan.DependencyDeletes.Count,

                    UnchangedDependencies =
                        plan.UnchangedDependencies,

                    CreatedAssignments =
                        plan.AssignmentCreates.Count,

                    DeletedAssignments =
                        plan.AssignmentDeletes.Count,

                    UnchangedAssignments =
                        plan.UnchangedAssignments,

                    Verified =
                        verification.Verified,

                    VerificationWarning =
                        !verification.Verified &&
                        !verification.TerminalFailure,

                    Failed =
                        verification.TerminalFailure,

                    VerificationMessage =
                        verification.Message,

                    ErrorMessage =
                        verification.TerminalFailure
                            ? verification.Message
                            : null,

                    SnapshotTakenOn =
                        snapshot.SnapshotTakenOn,

                    PublishedProjectFingerprint =
                        snapshotFingerprint,

                    AuditLogPath =
                        GetAuditLogPath()
                };

            if (verification.Verified)
            {
                result.VerifiedOn =
                    DateTime.Now;

                SetVerifiedBaselines(
                    verification.Schedule,
                    snapshot);
            }
            else if (verification.TerminalFailure)
            {
                ClearPendingPublish(
                    projectId);
            }
            else
            {
                StorePendingPublish(
                    projectId,
                    operationSetId,
                    snapshot,
                    plan);
            }

            return result;
        }

        // =========================================================
        // TASK PAYLOADS
        // =========================================================

        private Entity CreateTaskEntity(
            TaskSnapshot task,
            EntityReference projectReference,
            Guid defaultBucketId)
        {
            Entity entity =
                new Entity(
                    "msdyn_projecttask",
                    task.Id);

            SetCrmLookup(
                entity,
                "msdyn_project",
                projectReference);

            SetCrmString(
                entity,
                "msdyn_subject",
                string.IsNullOrWhiteSpace(
                    task.Name)
                    ? "(İsimsiz Görev)"
                    : task.Name);

            SetCrmLookup(
                entity,
                "msdyn_projectbucket",
                new EntityReference(
                    "msdyn_projectbucket",
                    task.BucketId ??
                    defaultBucketId));

            SetCrmOptionSet(
                entity,
                "msdyn_LinkStatus",
                192350000);

            SetNormalizedPercent(
                entity,
                "msdyn_projecttask",
                task.PercentComplete);

            if (task.ParentId.HasValue)
            {
                SetCrmLookup(
                    entity,
                    "msdyn_parenttask",
                    new EntityReference(
                        "msdyn_projecttask",
                        task.ParentId.Value));
            }
            else
            {
                SetCrmInt(
                    entity,
                    "msdyn_outlinelevel",
                    1);
            }

            if (!task.IsSummary)
            {
                if (task.Start.HasValue)
                {
                    SetCrmDateTime(
                        entity,
                        "msdyn_start",
                        task.Start.Value);
                }

                SetCrmDouble(
                    entity,
                    "msdyn_duration",
                    task.DurationDays);
            }

            return entity;
        }

        private Entity CreateTaskUpdateEntity(
            TaskUpdatePlan update)
        {
            TaskSnapshot task =
                update.Task;

            Entity entity =
                new Entity(
                    "msdyn_projecttask",
                    task.Id);

            if (update.NameChanged)
            {
                SetCrmString(
                    entity,
                    "msdyn_subject",
                    task.Name ?? "");
            }

            if (update.PercentCompleteChanged)
            {
                SetNormalizedPercent(
                    entity,
                    "msdyn_projecttask",
                    task.PercentComplete);
            }

            if (!task.IsSummary)
            {
                if (update.StartChanged &&
                    task.Start.HasValue)
                {
                    SetCrmDateTime(
                        entity,
                        "msdyn_start",
                        task.Start.Value);
                }

                if (update.DurationChanged)
                {
                    SetCrmDouble(
                        entity,
                        "msdyn_duration",
                        task.DurationDays);
                }

                if (update.FinishChanged &&
                    task.Finish.HasValue)
                {
                    SetCrmDateTime(
                        entity,
                        "msdyn_finish",
                        task.Finish.Value);
                }
            }

            return entity;
        }

        // =========================================================
        // DEPENDENCY / RESOURCE
        // =========================================================

        private Entity CreateDependencyEntity(
            DependencySnapshot dependency,
            EntityReference projectReference)
        {
            Entity entity =
                new Entity(
                    "msdyn_projecttaskdependency",
                    Guid.NewGuid());

            SetCrmLookup(
                entity,
                "msdyn_project",
                projectReference);

            SetCrmLookup(
                entity,
                "msdyn_predecessortask",
                new EntityReference(
                    "msdyn_projecttask",
                    dependency.PredecessorId));

            SetCrmLookup(
                entity,
                "msdyn_successortask",
                new EntityReference(
                    "msdyn_projecttask",
                    dependency.SuccessorId));

            SetCrmOptionSet(
                entity,
                "msdyn_linktype",
                dependency.LinkType);

            return entity;
        }

        private Entity CreateResourceAssignmentEntity(
            AssignmentSnapshot assignment,
            EntityReference projectReference)
        {
            if (!assignment.TeamMemberId.HasValue)
            {
                throw new InvalidOperationException(
                    "CRM Project Team mapping bulunamadı.");
            }

            Entity entity =
                new Entity(
                    "msdyn_resourceassignment",
                    Guid.NewGuid());

            SetCrmLookup(
                entity,
                "msdyn_projectid",
                projectReference);

            SetCrmLookup(
                entity,
                "msdyn_taskid",
                new EntityReference(
                    "msdyn_projecttask",
                    assignment.TaskId));

            SetCrmLookup(
                entity,
                "msdyn_projectteamid",
                new EntityReference(
                    "msdyn_projectteam",
                    assignment.TeamMemberId.Value));

            SetCrmString(
                entity,
                "msdyn_name",
                (assignment.ResourceName ?? "Resource") +
                " - " +
                (assignment.TaskName ?? "Task"));

            return entity;
        }

        // =========================================================
        // CUSTOM PERCENT FIELD METADATA
        // =========================================================

        private AttributeMetadata GetPercentAttributeMetadata(
            string entityLogicalName)
        {
            string key =
                entityLogicalName +
                "|" +
                PercentCompletedField;

            lock (_metadataLock)
            {
                AttributeMetadata cached;

                if (_percentMetadata.TryGetValue(
                    key,
                    out cached))
                {
                    return cached;
                }
            }

            RetrieveAttributeRequest request =
                new RetrieveAttributeRequest
                {
                    EntityLogicalName =
                        entityLogicalName,

                    LogicalName =
                        PercentCompletedField,

                    RetrieveAsIfPublished =
                        true
                };

            RetrieveAttributeResponse response;

            try
            {
                response =
                    (RetrieveAttributeResponse)
                    _client.Execute(
                        request);
            }
            catch (System.Exception ex)
            {
                throw new InvalidOperationException(
                    entityLogicalName +
                    "." +
                    PercentCompletedField +
                    " alanı Dataverse metadata içerisinde bulunamadı.",
                    ex);
            }

            AttributeMetadata metadata =
                response.AttributeMetadata;

            if (metadata == null)
            {
                throw new InvalidOperationException(
                    entityLogicalName +
                    "." +
                    PercentCompletedField +
                    " metadata bilgisi alınamadı.");
            }

            lock (_metadataLock)
            {
                _percentMetadata[
                    key] =
                    metadata;
            }

            return metadata;
        }

        private void SetNormalizedPercent(
            Entity entity,
            string entityLogicalName,
            int percent)
        {
            percent =
                ClampPercent(
                    percent);

            AttributeMetadata metadata =
                GetPercentAttributeMetadata(
                    entityLogicalName);

            AttributeTypeCode? type =
                metadata.AttributeType;

            if (!type.HasValue)
            {
                throw new InvalidOperationException(
                    entityLogicalName +
                    "." +
                    PercentCompletedField +
                    " veri tipi alınamadı.");
            }

            switch (type.Value)
            {
                case AttributeTypeCode.Integer:

                    entity[
                        PercentCompletedField] =
                        percent;

                    return;

                case AttributeTypeCode.Decimal:

                    entity[
                        PercentCompletedField] =
                        Convert.ToDecimal(
                            percent,
                            CultureInfo.InvariantCulture);

                    return;

                case AttributeTypeCode.Double:

                    entity[
                        PercentCompletedField] =
                        Convert.ToDouble(
                            percent,
                            CultureInfo.InvariantCulture);

                    return;

                case AttributeTypeCode.Money:

                    entity[
                        PercentCompletedField] =
                        new Money(
                            Convert.ToDecimal(
                                percent,
                                CultureInfo.InvariantCulture));

                    return;

                case AttributeTypeCode.BigInt:

                    entity[
                        PercentCompletedField] =
                        Convert.ToInt64(
                            percent,
                            CultureInfo.InvariantCulture);

                    return;

                default:

                    throw new InvalidOperationException(
                        entityLogicalName +
                        "." +
                        PercentCompletedField +
                        " alanı desteklenmeyen veri tipinde: " +
                        type.Value);
            }
        }

        // =========================================================
        // BASIC CRM TYPES
        // =========================================================

        private static void SetCrmString(
            Entity entity,
            string field,
            string value)
        {
            entity[field] =
                value ?? "";
        }

        private static void SetCrmDouble(
            Entity entity,
            string field,
            double value)
        {
            if (double.IsNaN(
                    value) ||
                double.IsInfinity(
                    value))
            {
                throw new InvalidOperationException(
                    field +
                    " için geçersiz Double değeri.");
            }

            entity[field] =
                value;
        }

        private static void SetCrmInt(
            Entity entity,
            string field,
            int value)
        {
            entity[field] =
                value;
        }

        private static void SetCrmDateTime(
            Entity entity,
            string field,
            DateTime value)
        {
            entity[field] =
                value;
        }

        private static void SetCrmLookup(
            Entity entity,
            string field,
            EntityReference value)
        {
            if (value == null ||
                value.Id == Guid.Empty)
            {
                throw new InvalidOperationException(
                    field +
                    " lookup değeri geçersiz.");
            }

            entity[field] =
                value;
        }

        private static void SetCrmOptionSet(
            Entity entity,
            string field,
            int value)
        {
            entity[field] =
                new OptionSetValue(
                    value);
        }

        // =========================================================
        // VERIFY
        // =========================================================

        private VerificationOutcome WaitForVerification(
            Guid projectId,
            ProjectSnapshot publishedSnapshot,
            PublishPlan plan,
            Guid operationSetId,
            Action<PublishProgressInfo> progress)
        {
            DateTime deadline =
                DateTime.UtcNow.AddSeconds(
                    QuickVerificationSeconds);

            string lastMismatch =
                null;

            CrmScheduleData lastSchedule =
                null;

            while (DateTime.UtcNow <
                   deadline)
            {
                OperationSetState state =
                    TryGetOperationSetState(
                        operationSetId);

                if (state != null &&
                    state.ErrorLog != null)
                {
                    return
                        new VerificationOutcome
                        {
                            Verified =
                                false,

                            TerminalFailure =
                                true,

                            Schedule =
                                lastSchedule,

                            Message =
                                "CRM scheduling işlemi başarısız oldu." +
                                "\r\n\r\n" +
                                ReadPssError(
                                    state.ErrorLog)
                        };
                }

                try
                {
                    lastSchedule =
                        GetScheduleDataSync(
                            projectId);

                    if (TryVerifyPublishedChanges(
                        publishedSnapshot,
                        lastSchedule,
                        plan,
                        out lastMismatch))
                    {
                        return
                            new VerificationOutcome
                            {
                                Verified =
                                    true,

                                Schedule =
                                    lastSchedule
                            };
                    }
                }
                catch (System.Exception ex)
                {
                    lastMismatch =
                        "CRM read-back yapılamadı." +
                        "\r\n\r\n" +
                        GetServiceExceptionMessage(
                            ex);
                }

                Thread.Sleep(
                    1000);
            }

            return
                new VerificationOutcome
                {
                    Verified =
                        false,

                    TerminalFailure =
                        false,

                    Schedule =
                        lastSchedule,

                    Message =
                        "CRM işlemi gönderildi ancak sonuç henüz doğrulanamadı." +
                        "\r\n\r\n" +
                        (
                            string.IsNullOrWhiteSpace(
                                lastMismatch)
                                ? ""
                                : "Son kontrol:\r\n" +
                                  lastMismatch +
                                  "\r\n\r\n"
                        ) +
                        "Operation Set: " +
                        operationSetId +
                        "\r\n\r\nRe-verify ile tekrar kontrol edebilirsiniz."
                };
        }

        private bool TryVerifyPublishedChanges(
            ProjectSnapshot project,
            CrmScheduleData crm,
            PublishPlan plan,
            out string mismatch)
        {
            mismatch =
                null;

            if (plan.ProjectPercentChanged)
            {
                int? crmProjectPercent =
                    GetPercentValue(
                        crm.Project,
                        PercentCompletedField);

                if (!crmProjectPercent.HasValue ||
                    crmProjectPercent.Value !=
                    project.ProjectPercentComplete)
                {
                    mismatch =
                        "Project: % Complete güncellenemedi." +
                        "\r\nBeklenen: " +
                        project.ProjectPercentComplete +
                        "%" +
                        "\r\nCRM: " +
                        (
                            crmProjectPercent.HasValue
                                ? crmProjectPercent.Value +
                                  "%"
                                : "(boş)"
                        );

                    return false;
                }
            }

            Dictionary<Guid, Entity> crmTasks =
                crm.Tasks.Entities
                    .ToDictionary(
                        x => x.Id);

            foreach (TaskDeletePlan deletedTask
                     in plan.TaskDeletes)
            {
                if (crmTasks.ContainsKey(
                    deletedTask.TaskId))
                {
                    mismatch =
                        deletedTask.TaskName +
                        ": Task halen CRM'de mevcut.";

                    return false;
                }
            }

            foreach (TaskSnapshot expected
                     in plan.TaskCreates)
            {
                Entity actual;

                if (!crmTasks.TryGetValue(
                    expected.Id,
                    out actual))
                {
                    mismatch =
                        expected.Name +
                        ": Task henüz CRM'de bulunamadı.";

                    return false;
                }

                string crmName =
                    actual.GetAttributeValue<string>(
                        "msdyn_subject")
                    ?? "";

                if (!string.Equals(
                    NormalizeText(
                        expected.Name),
                    NormalizeText(
                        crmName),
                    StringComparison.Ordinal))
                {
                    mismatch =
                        expected.Name +
                        ": Task Name eşleşmiyor." +
                        "\r\nBeklenen: " +
                        expected.Name +
                        "\r\nCRM: " +
                        crmName;

                    return false;
                }

                int? crmPercent =
                    GetPercentValue(
                        actual,
                        PercentCompletedField);

                if (!crmPercent.HasValue ||
                    crmPercent.Value !=
                    expected.PercentComplete)
                {
                    mismatch =
                        expected.Name +
                        ": % Complete eşleşmiyor." +
                        "\r\nBeklenen: " +
                        expected.PercentComplete +
                        "%" +
                        "\r\nCRM: " +
                        (
                            crmPercent.HasValue
                                ? crmPercent.Value +
                                  "%"
                                : "(boş)"
                        );

                    return false;
                }

                if (!expected.IsSummary)
                {
                    DateTime? crmStart =
                        GetTaskStart(
                            actual);

                    if (!SameCalendarDate(
                        expected.Start,
                        crmStart))
                    {
                        mismatch =
                            expected.Name +
                            ": Start eşleşmiyor." +
                            "\r\nBeklenen: " +
                            DisplayDate(
                                expected.Start) +
                            "\r\nCRM: " +
                            DisplayDate(
                                crmStart);

                        return false;
                    }

                    double? crmDuration =
                        GetNumber(
                            actual,
                            "msdyn_duration");

                    if (!crmDuration.HasValue ||
                        !SameDuration(
                            expected.DurationDays,
                            crmDuration.Value))
                    {
                        mismatch =
                            expected.Name +
                            ": Duration eşleşmiyor." +
                            "\r\nBeklenen: " +
                            expected.DurationDays.ToString(
                                "0.####",
                                CultureInfo.InvariantCulture) +
                            " days" +
                            "\r\nCRM: " +
                            (
                                crmDuration.HasValue
                                    ? crmDuration.Value.ToString(
                                        "0.####",
                                        CultureInfo.InvariantCulture) +
                                      " days"
                                    : "(boş)"
                            );

                        return false;
                    }
                }
            }

            foreach (TaskUpdatePlan update
                     in plan.TaskUpdates)
            {
                Entity actual;

                if (!crmTasks.TryGetValue(
                    update.Task.Id,
                    out actual))
                {
                    mismatch =
                        update.Task.Name +
                        ": Task CRM'de bulunamadı.";

                    return false;
                }

                if (update.NameChanged)
                {
                    string crmName =
                        actual.GetAttributeValue<string>(
                            "msdyn_subject")
                        ?? "";

                    if (!string.Equals(
                        NormalizeText(
                            update.Task.Name),
                        NormalizeText(
                            crmName),
                        StringComparison.Ordinal))
                    {
                        mismatch =
                            update.Task.Name +
                            ": Task Name güncellenemedi.";

                        return false;
                    }
                }

                if (update.PercentCompleteChanged)
                {
                    int? crmPercent =
                        GetPercentValue(
                            actual,
                            PercentCompletedField);

                    if (!crmPercent.HasValue ||
                        crmPercent.Value !=
                        update.Task.PercentComplete)
                    {
                        mismatch =
                            update.Task.Name +
                            ": % Complete güncellenemedi." +
                            "\r\nBeklenen: " +
                            update.Task.PercentComplete +
                            "%" +
                            "\r\nCRM: " +
                            (
                                crmPercent.HasValue
                                    ? crmPercent.Value +
                                      "%"
                                    : "(boş)"
                            );

                        return false;
                    }
                }

                if (!update.Task.IsSummary)
                {
                    if (update.StartChanged)
                    {
                        DateTime? crmStart =
                            GetTaskStart(
                                actual);

                        if (!SameCalendarDate(
                            update.Task.Start,
                            crmStart))
                        {
                            mismatch =
                                update.Task.Name +
                                ": Start güncellenemedi." +
                                "\r\nBeklenen: " +
                                DisplayDate(
                                    update.Task.Start) +
                                "\r\nCRM: " +
                                DisplayDate(
                                    crmStart);

                            return false;
                        }
                    }

                    if (update.DurationChanged)
                    {
                        double? crmDuration =
                            GetNumber(
                                actual,
                                "msdyn_duration");

                        if (!crmDuration.HasValue ||
                            !SameDuration(
                                update.Task.DurationDays,
                                crmDuration.Value))
                        {
                            mismatch =
                                update.Task.Name +
                                ": Duration güncellenemedi." +
                                "\r\nBeklenen: " +
                                update.Task.DurationDays.ToString(
                                    "0.####",
                                    CultureInfo.InvariantCulture) +
                                " days" +
                                "\r\nCRM: " +
                                (
                                    crmDuration.HasValue
                                        ? crmDuration.Value.ToString(
                                            "0.####",
                                            CultureInfo.InvariantCulture) +
                                          " days"
                                        : "(boş)"
                                );

                            return false;
                        }
                    }

                    if (update.FinishChanged)
                    {
                        DateTime? crmFinish =
                            GetTaskFinish(
                                actual);

                        if (!SameCalendarDate(
                            update.Task.Finish,
                            crmFinish))
                        {
                            mismatch =
                                update.Task.Name +
                                ": Finish güncellenemedi." +
                                "\r\nBeklenen: " +
                                DisplayDate(
                                    update.Task.Finish) +
                                "\r\nCRM: " +
                                DisplayDate(
                                    crmFinish);

                            return false;
                        }
                    }
                }
            }

            HashSet<string> crmDependencyKeys =
                new HashSet<string>(
                    crm.Dependencies.Entities
                        .Select(
                            GetCrmDependencyKey)
                        .Where(
                            x =>
                                !string.IsNullOrWhiteSpace(
                                    x)),
                    StringComparer.OrdinalIgnoreCase);

            foreach (DependencySnapshot dependency
                     in plan.DependencyCreates)
            {
                if (!crmDependencyKeys.Contains(
                    dependency.Key))
                {
                    mismatch =
                        "Predecessor henüz CRM'de bulunamadı.";

                    return false;
                }
            }

            foreach (DependencyDeletePlan dependency
                     in plan.DependencyDeletes)
            {
                if (crmDependencyKeys.Contains(
                    dependency.Key))
                {
                    mismatch =
                        "Silinen predecessor halen CRM'de mevcut.";

                    return false;
                }
            }

            HashSet<string> crmAssignmentKeys =
                new HashSet<string>(
                    BuildCrmAssignmentMap(
                        crm)
                        .Keys,
                    StringComparer.OrdinalIgnoreCase);

            foreach (AssignmentSnapshot assignment
                     in plan.AssignmentCreates)
            {
                if (!crmAssignmentKeys.Contains(
                    assignment.Key))
                {
                    mismatch =
                        assignment.TaskName +
                        ": Resource assignment henüz CRM'de bulunamadı." +
                        "\r\nResource: " +
                        assignment.ResourceName;

                    return false;
                }
            }

            foreach (AssignmentDeletePlan assignment
                     in plan.AssignmentDeletes)
            {
                if (crmAssignmentKeys.Contains(
                    assignment.Key))
                {
                    mismatch =
                        "Silinen resource assignment halen CRM'de mevcut.";

                    return false;
                }
            }

            return true;
        }

        // =========================================================
        // CRM FINGERPRINT
        // =========================================================

        private string ComputeScheduleFingerprint(
            CrmScheduleData schedule)
        {
            List<string> values =
                new List<string>();

            int projectPercent =
                GetPercentValue(
                    schedule.Project,
                    PercentCompletedField)
                ?? 0;

            values.Add(
                "P|" +
                projectPercent);

            foreach (Entity task
                     in schedule.Tasks.Entities
                         .OrderBy(
                             x => x.Id))
            {
                values.Add(
                    "T|" +
                    task.Id +
                    "|" +
                    NormalizeFingerprintText(
                        GetString(
                            task,
                            "msdyn_subject")) +
                    "|PC=" +
                    (
                        GetPercentValue(
                            task,
                            PercentCompletedField)
                        ?? 0
                    ) +
                    "|" +
                    DateText(
                        GetTaskStart(
                            task)) +
                    "|" +
                    DateText(
                        GetTaskFinish(
                            task)) +
                    "|" +
                    (
                        GetNumber(
                            task,
                            "msdyn_duration")
                        ?? 0
                    ).ToString(
                        "0.####",
                        CultureInfo.InvariantCulture));
            }

            foreach (string key
                     in schedule.Dependencies.Entities
                         .Select(
                             GetCrmDependencyKey)
                         .Where(
                             x =>
                                 !string.IsNullOrWhiteSpace(
                                     x))
                         .OrderBy(
                             x => x,
                             StringComparer.OrdinalIgnoreCase))
            {
                values.Add(
                    "D|" +
                    key);
            }

            foreach (string key
                     in BuildCrmAssignmentMap(
                         schedule)
                         .Keys
                         .OrderBy(
                             x => x,
                             StringComparer.OrdinalIgnoreCase))
            {
                values.Add(
                    "A|" +
                    key);
            }

            return ComputeSha256(
                string.Join(
                    "\n",
                    values));
        }

        // =========================================================
        // MAPS
        // =========================================================

        private Dictionary<string, Entity> BuildCrmAssignmentMap(
            CrmScheduleData crm)
        {
            Dictionary<string, Entity> result =
                new Dictionary<string, Entity>(
                    StringComparer.OrdinalIgnoreCase);

            foreach (Entity assignment
                     in crm.ResourceAssignments.Entities)
            {
                EntityReference task =
                    assignment.GetAttributeValue<EntityReference>(
                        "msdyn_taskid");

                EntityReference team =
                    assignment.GetAttributeValue<EntityReference>(
                        "msdyn_projectteamid");

                if (task == null ||
                    team == null)
                {
                    continue;
                }

                string key =
                    task.Id +
                    "|" +
                    team.Id;

                if (!result.ContainsKey(
                    key))
                {
                    result[
                        key] =
                        assignment;
                }
            }

            return result;
        }

        private static string GetCrmDependencyKey(
            Entity dependency)
        {
            EntityReference predecessor =
                dependency.GetAttributeValue<EntityReference>(
                    "msdyn_predecessortask");

            EntityReference successor =
                dependency.GetAttributeValue<EntityReference>(
                    "msdyn_successortask");

            if (predecessor == null ||
                successor == null)
            {
                return null;
            }

            OptionSetValue type =
                dependency.GetAttributeValue<OptionSetValue>(
                    "msdyn_linktype");

            int linkType =
                type != null
                    ? type.Value
                    : 192350000;

            return
                predecessor.Id +
                "|" +
                successor.Id +
                "|" +
                linkType;
        }

        // =========================================================
        // SCHEDULE API
        // =========================================================

        private bool TryCreateOperationSet(
            Guid projectId,
            string description,
            out Guid operationSetId,
            out string error)
        {
            operationSetId =
                Guid.Empty;

            error =
                null;

            OrganizationRequest request =
                new OrganizationRequest(
                    "msdyn_CreateOperationSetV1");

            request["ProjectId"] =
                projectId.ToString();

            request["Description"] =
                description;

            OrganizationResponse response;

            if (!TryExecuteRequest(
                "Operation Set oluşturma",
                request,
                out response,
                out error))
            {
                return false;
            }

            try
            {
                object value =
                    response[
                        "OperationSetId"];

                operationSetId =
                    value is Guid
                        ? (Guid)value
                        : Guid.Parse(
                            value.ToString());

                return
                    operationSetId !=
                    Guid.Empty;
            }
            catch (System.Exception ex)
            {
                error =
                    "Operation Set ID alınamadı." +
                    "\r\n\r\n" +
                    ex.Message;

                return false;
            }
        }

        private bool TryPssCreate(
            Entity entity,
            Guid operationSetId,
            out string error)
        {
            OrganizationRequest request =
                new OrganizationRequest(
                    "msdyn_PssCreateV1");

            request["Entity"] =
                entity;

            request["OperationSetId"] =
                operationSetId.ToString();

            OrganizationResponse response;

            return TryExecuteRequest(
                "Create " +
                entity.LogicalName,
                request,
                out response,
                out error);
        }

        private bool TryPssUpdate(
            Entity entity,
            Guid operationSetId,
            out string error)
        {
            OrganizationRequest request =
                new OrganizationRequest(
                    "msdyn_PssUpdateV1");

            request["Entity"] =
                entity;

            request["OperationSetId"] =
                operationSetId.ToString();

            OrganizationResponse response;

            return TryExecuteRequest(
                "Update " +
                entity.LogicalName,
                request,
                out response,
                out error);
        }

        private bool TryPssDelete(
            Guid id,
            string logicalName,
            Guid operationSetId,
            out string error)
        {
            OrganizationRequest request =
                new OrganizationRequest(
                    "msdyn_PssDeleteV1");

            request["RecordId"] =
                id.ToString();

            request["EntityLogicalName"] =
                logicalName;

            request["OperationSetId"] =
                operationSetId.ToString();

            OrganizationResponse response;

            return TryExecuteRequest(
                "Delete " +
                logicalName,
                request,
                out response,
                out error);
        }

        private bool TryExecuteOperationSet(
            Guid operationSetId,
            out string error)
        {
            OrganizationRequest request =
                new OrganizationRequest(
                    "msdyn_ExecuteOperationSetV1");

            request["OperationSetId"] =
                operationSetId.ToString();

            OrganizationResponse response;

            return TryExecuteRequest(
                "Execute Operation Set",
                request,
                out response,
                out error);
        }

        private void TryAbandonOperationSet(
            Guid operationSetId)
        {
            if (operationSetId ==
                Guid.Empty)
            {
                return;
            }

            try
            {
                OrganizationRequest request =
                    new OrganizationRequest(
                        "msdyn_AbandonOperationSetV1");

                request["OperationSetId"] =
                    operationSetId.ToString();

                OrganizationResponse response;

                string error;

                TryExecuteRequest(
                    "Abandon Operation Set",
                    request,
                    out response,
                    out error);
            }
            catch
            {
            }
        }

        private bool TryExecuteRequest(
            string operationName,
            OrganizationRequest request,
            out OrganizationResponse response,
            out string error)
        {
            response =
                null;

            error =
                null;

            try
            {
                response =
                    _client.Execute(
                        request);

                return true;
            }
            catch (System.Exception ex)
            {
                error =
                    operationName +
                    " başarısız oldu." +
                    "\r\n\r\n" +
                    GetServiceExceptionMessage(
                        ex);

                return false;
            }
        }

        // =========================================================
        // OPERATION SET STATE
        // =========================================================

        private OperationSetState TryGetOperationSetState(
            Guid operationSetId)
        {
            try
            {
                Entity entity =
                    _client.Retrieve(
                        "msdyn_operationset",
                        operationSetId,
                        new ColumnSet(
                            "msdyn_completedon",
                            "msdyn_status",
                            "msdyn_psserrorlog"));

                OperationSetState state =
                    new OperationSetState
                    {
                        CompletedOn =
                            GetDateTime(
                                entity,
                                "msdyn_completedon"),

                        ErrorLog =
                            entity.GetAttributeValue<EntityReference>(
                                "msdyn_psserrorlog")
                    };

                if (entity.FormattedValues.Contains(
                    "msdyn_status"))
                {
                    state.StatusText =
                        entity.FormattedValues[
                            "msdyn_status"];
                }
                else
                {
                    OptionSetValue status =
                        entity.GetAttributeValue<OptionSetValue>(
                            "msdyn_status");

                    if (status != null)
                    {
                        state.StatusText =
                            status.Value.ToString(
                                CultureInfo.InvariantCulture);
                    }
                }

                return state;
            }
            catch
            {
                return null;
            }
        }

        private string ReadPssError(
            EntityReference errorReference)
        {
            if (errorReference == null)
            {
                return
                    "PSS hata kaydı bulunamadı.";
            }

            try
            {
                Entity error =
                    _client.Retrieve(
                        "msdyn_psserrorlog",
                        errorReference.Id,
                        new ColumnSet(
                            "msdyn_errorcode",
                            "msdyn_log"));

                string code =
                    GetString(
                        error,
                        "msdyn_errorcode");

                string log =
                    GetString(
                        error,
                        "msdyn_log");

                return
                    (
                        string.IsNullOrWhiteSpace(
                            code)
                            ? ""
                            : "Error Code: " +
                              code +
                              "\r\n"
                    ) +
                    (
                        string.IsNullOrWhiteSpace(
                            log)
                            ? "PSS hata logu boş."
                            : log
                    );
            }
            catch (System.Exception ex)
            {
                return
                    "PSS hata logu okunamadı: " +
                    ex.Message;
            }
        }

        // =========================================================
        // AUDIT
        // =========================================================

        private string GetAuditLogPath()
        {
            return Path.Combine(
                Environment.GetFolderPath(
                    Environment.SpecialFolder.LocalApplicationData),
                "ProjectHub",
                "Logs",
                "publish-audit.log");
        }

        private void WriteAuditSuccess(
            Guid projectId,
            CrmPublishResult result)
        {
            WriteAudit(
                "SUCCESS",
                projectId,
                result.OperationSetId,
                "Verified=" +
                result.Verified +
                " | ProjectProgressUpdated=" +
                result.UpdatedProjectFields +
                " | CreatedTasks=" +
                result.CreatedTasks +
                " | UpdatedTasks=" +
                result.UpdatedTasks);
        }

        private void WriteAuditWarning(
            Guid projectId,
            CrmPublishResult result)
        {
            WriteAudit(
                result.Blocked
                    ? "BLOCKED"
                    : result.Failed
                        ? "FAILED"
                        : "WARNING",
                projectId,
                result.OperationSetId,
                NormalizeAuditText(
                    result.Failed
                        ? result.ErrorMessage
                        : result.VerificationMessage));
        }

        private void WriteAuditFailure(
            Guid projectId,
            string fingerprint,
            System.Exception ex)
        {
            WriteAudit(
                "FAILED",
                projectId,
                Guid.Empty,
                NormalizeAuditText(
                    ex != null
                        ? ex.Message
                        : ""));
        }

        private void WriteAudit(
            string status,
            Guid projectId,
            Guid operationSetId,
            string message)
        {
            try
            {
                string path =
                    GetAuditLogPath();

                Directory.CreateDirectory(
                    Path.GetDirectoryName(
                        path));

                File.AppendAllText(
                    path,

                    DateTime.Now.ToString(
                        "yyyy-MM-dd HH:mm:ss") +
                    " | " +
                    status +
                    " | CRM=" +
                    (ConnectedUrl ?? "") +
                    " | Project=" +
                    projectId +
                    " | OperationSet=" +
                    operationSetId +
                    " | " +
                    (message ?? "") +
                    Environment.NewLine,

                    Encoding.UTF8);
            }
            catch
            {
            }
        }

        // =========================================================
        // HELPERS
        // =========================================================

        private static int? GetPercentValue(
            Entity entity,
            string field)
        {
            double? number =
                GetNumber(
                    entity,
                    field);

            if (!number.HasValue)
                return null;

            return ClampPercent(
                Convert.ToInt32(
                    Math.Round(
                        number.Value,
                        MidpointRounding.AwayFromZero)));
        }

        private static bool SameProjectMinute(
            DateTime? left,
            DateTime? right)
        {
            if (!left.HasValue &&
                !right.HasValue)
            {
                return true;
            }

            if (!left.HasValue ||
                !right.HasValue)
            {
                return false;
            }

            return
                TrimToMinute(
                    left.Value)
                ==
                TrimToMinute(
                    right.Value);
        }

        private static bool SameCalendarDate(
            DateTime? projectValue,
            DateTime? crmValue)
        {
            if (!projectValue.HasValue &&
                !crmValue.HasValue)
            {
                return true;
            }

            if (!projectValue.HasValue ||
                !crmValue.HasValue)
            {
                return false;
            }

            DateTime a =
                projectValue.Value;

            DateTime b =
                crmValue.Value;

            if (a.Kind ==
                DateTimeKind.Utc)
            {
                a =
                    a.ToLocalTime();
            }

            if (b.Kind ==
                DateTimeKind.Utc)
            {
                b =
                    b.ToLocalTime();
            }

            return
                a.Date ==
                b.Date;
        }

        private static bool SameDuration(
            double projectDays,
            double crmDays)
        {
            return
                Math.Abs(
                    projectDays -
                    crmDays)
                <= 0.01;
        }

        private static DateTime TrimToMinute(
            DateTime value)
        {
            return
                new DateTime(
                    value.Year,
                    value.Month,
                    value.Day,
                    value.Hour,
                    value.Minute,
                    0,
                    value.Kind);
        }

        private static string DisplayDate(
            DateTime? value)
        {
            if (!value.HasValue)
                return "(boş)";

            return
                value.Value.ToString(
                    "yyyy-MM-dd HH:mm:ss") +
                " [" +
                value.Value.Kind +
                "]";
        }

        private static DateTime? GetTaskStart(
            Entity task)
        {
            return
                GetDateTime(
                    task,
                    "msdyn_start")
                ??
                GetDateTime(
                    task,
                    "msdyn_scheduledstart");
        }

        private static DateTime? GetTaskFinish(
            Entity task)
        {
            return
                GetDateTime(
                    task,
                    "msdyn_finish")
                ??
                GetDateTime(
                    task,
                    "msdyn_scheduledend");
        }

        private static Guid GetDefaultBucketId(
            CrmScheduleData schedule)
        {
            if (schedule.Buckets == null ||
                schedule.Buckets.Entities.Count == 0)
            {
                throw new InvalidOperationException(
                    "CRM projesinde Project Bucket bulunamadı.");
            }

            return
                schedule.Buckets.Entities[0].Id;
        }

        private static int ConvertLinkType(
            MSProject.PjTaskLinkType type)
        {
            switch (type)
            {
                case MSProject.PjTaskLinkType.pjStartToStart:

                    return 192350001;

                case MSProject.PjTaskLinkType.pjFinishToFinish:

                    return 192350002;

                case MSProject.PjTaskLinkType.pjStartToFinish:

                    return 192350004;

                default:

                    return 192350000;
            }
        }

        private static string NormalizeUrl(
            string url)
        {
            if (string.IsNullOrWhiteSpace(
                url))
            {
                return url;
            }

            url =
                url.Trim();

            if (!url.StartsWith(
                    "http://",
                    StringComparison.OrdinalIgnoreCase) &&
                !url.StartsWith(
                    "https://",
                    StringComparison.OrdinalIgnoreCase))
            {
                url =
                    "https://" +
                    url;
            }

            Uri uri;

            if (!Uri.TryCreate(
                url,
                UriKind.Absolute,
                out uri))
            {
                return url;
            }

            return
                uri.GetLeftPart(
                    UriPartial.Authority);
        }

        private static string GetString(
            Entity entity,
            string field)
        {
            if (entity == null ||
                !entity.Attributes.Contains(
                    field) ||
                entity[field] == null)
            {
                return "";
            }

            return
                Convert.ToString(
                    entity[field],
                    CultureInfo.InvariantCulture)
                ?? "";
        }

        private static DateTime? GetDateTime(
            Entity entity,
            string field)
        {
            if (entity == null ||
                !entity.Attributes.Contains(
                    field) ||
                entity[field] == null)
            {
                return null;
            }

            if (entity[field] is DateTime)
            {
                return
                    (DateTime)entity[field];
            }

            DateTime value;

            if (DateTime.TryParse(
                entity[field].ToString(),
                out value))
            {
                return value;
            }

            return null;
        }

        private static double? GetNumber(
            Entity entity,
            string field)
        {
            if (entity == null ||
                !entity.Attributes.Contains(
                    field) ||
                entity[field] == null)
            {
                return null;
            }

            object value =
                entity[field];

            try
            {
                if (value is Money)
                {
                    return
                        Convert.ToDouble(
                            ((Money)value).Value,
                            CultureInfo.InvariantCulture);
                }

                if (value is OptionSetValue)
                {
                    return
                        ((OptionSetValue)value).Value;
                }

                return
                    Convert.ToDouble(
                        value,
                        CultureInfo.InvariantCulture);
            }
            catch
            {
                return null;
            }
        }

        private static string NormalizeText(
            string value)
        {
            return
                (value ?? "")
                    .Replace(
                        "\r\n",
                        "\n")
                    .Replace(
                        "\r",
                        "\n")
                    .Trim();
        }

        private static string NormalizeFingerprintText(
            string value)
        {
            return
                NormalizeText(
                    value)
                    .Replace(
                        "|",
                        "\\|");
        }

        private static string NormalizeAuditText(
            string value)
        {
            return
                (value ?? "")
                    .Replace(
                        "\r",
                        " ")
                    .Replace(
                        "\n",
                        " ")
                    .Trim();
        }

        private static string NullableGuidText(
            Guid? value)
        {
            return
                value.HasValue
                    ? value.Value.ToString()
                    : "";
        }

        private static string DateText(
            DateTime? value)
        {
            if (!value.HasValue)
                return "";

            return
                TrimToMinute(
                    value.Value)
                    .ToString(
                        "yyyy-MM-dd HH:mm",
                        CultureInfo.InvariantCulture);
        }

        private static string ComputeSha256(
            string value)
        {
            using (SHA256 sha =
                   SHA256.Create())
            {
                byte[] hash =
                    sha.ComputeHash(
                        Encoding.UTF8.GetBytes(
                            value ?? ""));

                return
                    BitConverter
                        .ToString(
                            hash)
                        .Replace(
                            "-",
                            "");
            }
        }

        private static string GetServiceExceptionMessage(
            System.Exception ex)
        {
            if (ex == null)
                return "Bilinmeyen CRM hatası.";

            StringBuilder builder =
                new StringBuilder();

            System.Exception current =
                ex;

            int level =
                0;

            while (current != null &&
                   level < 5)
            {
                if (!string.IsNullOrWhiteSpace(
                    current.Message))
                {
                    if (builder.Length > 0)
                    {
                        builder.AppendLine();
                    }

                    builder.Append(
                        current.Message);
                }

                current =
                    current.InnerException;

                level++;
            }

            return
                builder.ToString();
        }

        private void EnsureConnected()
        {
            if (_client == null ||
                !_client.IsReady)
            {
                throw new InvalidOperationException(
                    "Dataverse bağlantısı hazır değil.");
            }
        }

        // =========================================================
        // DISPOSE
        // =========================================================

        public void Dispose()
        {
            if (_client != null)
            {
                try
                {
                    _client.Dispose();
                }
                catch
                {
                }
            }

            _client =
                null;

            _msalClient =
                null;

            _account =
                null;

            ConnectedUrl =
                null;

            _cachedProjects.Clear();

            lock (_baselineLock)
            {
                _loadedCrmBaselines.Clear();
                _loadedProjectBaselines.Clear();
                _pendingPublishes.Clear();
            }

            lock (_metadataLock)
            {
                _percentMetadata.Clear();
            }

            _publishGate.Dispose();
        }

        // =========================================================
        // INTERNAL MODELS
        // =========================================================

        private class ProjectSnapshot
        {
            public Guid ProjectId { get; set; }

            public DateTime SnapshotTakenOn { get; set; }

            public double HoursPerDay { get; set; }

            public int ProjectPercentComplete { get; set; }

            public Dictionary<int, Guid> TaskIds =
                new Dictionary<int, Guid>();

            public List<TaskSnapshot> Tasks =
                new List<TaskSnapshot>();

            public List<DependencySnapshot> Dependencies =
                new List<DependencySnapshot>();

            public List<AssignmentSnapshot> Assignments =
                new List<AssignmentSnapshot>();
        }

        private class TaskSnapshot
        {
            public Guid Id { get; set; }

            public string Name { get; set; }

            public DateTime? Start { get; set; }

            public DateTime? Finish { get; set; }

            public double DurationMinutes { get; set; }

            public double DurationDays { get; set; }

            public int PercentComplete { get; set; }

            public Guid? ParentId { get; set; }

            public Guid? BucketId { get; set; }

            public bool IsSummary { get; set; }

            public int OutlineLevel { get; set; }
        }

        private class TaskUpdatePlan
        {
            public TaskSnapshot Task { get; set; }

            public bool NameChanged { get; set; }

            public bool StartChanged { get; set; }

            public bool FinishChanged { get; set; }

            public bool DurationChanged { get; set; }

            public bool PercentCompleteChanged { get; set; }

            public bool HasChanges
            {
                get
                {
                    return
                        NameChanged ||
                        StartChanged ||
                        FinishChanged ||
                        DurationChanged ||
                        PercentCompleteChanged;
                }
            }
        }

        private class TaskDeletePlan
        {
            public Guid TaskId { get; set; }

            public string TaskName { get; set; }

            public int OutlineLevel { get; set; }
        }

        private class DependencySnapshot
        {
            public Guid PredecessorId { get; set; }

            public Guid SuccessorId { get; set; }

            public int LinkType { get; set; }

            public string Key
            {
                get
                {
                    return
                        PredecessorId +
                        "|" +
                        SuccessorId +
                        "|" +
                        LinkType;
                }
            }
        }

        private class DependencyDeletePlan
        {
            public Guid Id { get; set; }

            public string Key { get; set; }
        }

        private class AssignmentSnapshot
        {
            public Guid TaskId { get; set; }

            public Guid? TeamMemberId { get; set; }

            public Guid? BookableResourceId { get; set; }

            public string TaskName { get; set; }

            public string ResourceName { get; set; }

            public string Key
            {
                get
                {
                    return
                        TaskId +
                        "|" +
                        TeamMemberId;
                }
            }
        }

        private class AssignmentDeletePlan
        {
            public Guid Id { get; set; }

            public string Key { get; set; }
        }

        private class PublishPlan
        {
            public bool ProjectPercentChanged;

            public List<TaskSnapshot> TaskCreates =
                new List<TaskSnapshot>();

            public List<TaskUpdatePlan> TaskUpdates =
                new List<TaskUpdatePlan>();

            public List<TaskDeletePlan> TaskDeletes =
                new List<TaskDeletePlan>();

            public List<DependencySnapshot> DependencyCreates =
                new List<DependencySnapshot>();

            public List<DependencyDeletePlan> DependencyDeletes =
                new List<DependencyDeletePlan>();

            public int UnchangedDependencies;

            public List<AssignmentSnapshot> AssignmentCreates =
                new List<AssignmentSnapshot>();

            public List<AssignmentDeletePlan> AssignmentDeletes =
                new List<AssignmentDeletePlan>();

            public int UnchangedAssignments;

            public int OperationCount
            {
                get
                {
                    return
                        (
                            ProjectPercentChanged
                                ? 1
                                : 0
                        ) +
                        TaskCreates.Count +
                        TaskUpdates.Count +
                        TaskDeletes.Count +
                        DependencyCreates.Count +
                        DependencyDeletes.Count +
                        AssignmentCreates.Count +
                        AssignmentDeletes.Count;
                }
            }
        }

        private class PendingPublishState
        {
            public Guid ProjectId { get; set; }

            public Guid OperationSetId { get; set; }

            public ProjectSnapshot Snapshot { get; set; }

            public PublishPlan Plan { get; set; }

            public DateTime StartedOn { get; set; }
        }

        private class VerificationOutcome
        {
            public bool Verified { get; set; }

            public bool TerminalFailure { get; set; }

            public string Message { get; set; }

            public CrmScheduleData Schedule { get; set; }
        }

        private class OperationSetState
        {
            public DateTime? CompletedOn { get; set; }

            public string StatusText { get; set; }

            public EntityReference ErrorLog { get; set; }
        }
    }
}