using System;

namespace ProjectHub.AddIn
{
    public enum PublishStatusLevel
    {
        Info,
        Success,
        Warning,
        Error
    }

    public class PublishProgressInfo
    {
        public DateTime Time { get; set; }

        public PublishStatusLevel Level { get; set; }

        public string Title { get; set; }

        public string Message { get; set; }
    }

    public class CrmPublishResult
    {
        public Guid OperationSetId { get; set; }

        public bool NoChanges { get; set; }

        public bool Verified { get; set; }

        public bool VerificationWarning { get; set; }

        public bool Failed { get; set; }

        public bool Blocked { get; set; }

        public string VerificationMessage { get; set; }

        public string ErrorMessage { get; set; }

        public DateTime SnapshotTakenOn { get; set; }

        public DateTime? VerifiedOn { get; set; }

        public string PublishedProjectFingerprint { get; set; }

        public string AuditLogPath { get; set; }

        public int UpdatedProjectFields { get; set; }

        public int CreatedTasks { get; set; }

        public int UpdatedTasks { get; set; }

        public int DeletedTasks { get; set; }

        public int CreatedDependencies { get; set; }

        public int DeletedDependencies { get; set; }

        public int UnchangedDependencies { get; set; }

        public int CreatedAssignments { get; set; }

        public int DeletedAssignments { get; set; }

        public int UnchangedAssignments { get; set; }
    }
}