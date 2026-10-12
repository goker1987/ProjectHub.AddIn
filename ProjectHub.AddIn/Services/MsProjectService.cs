using Microsoft.Xrm.Sdk;
using System;
using System.Collections.Generic;
using System.Globalization;
using System.Linq;

using MSProject = Microsoft.Office.Interop.MSProject;

namespace ProjectHub.AddIn
{
    public static class MsProjectService
    {
        private const string PercentCompletedField =
            "tt_percentcompleted";

        // =========================================================
        // IMPORT
        // =========================================================

        public static void ImportProject(
            CrmScheduleData schedule)
        {
            if (schedule == null)
            {
                throw new ArgumentNullException(
                    "schedule");
            }

            MSProject.Application app =
                Globals.ThisAddIn.Application;

            MSProject.Project project =
                app.ActiveProject;

            if (project == null)
            {
                throw new InvalidOperationException(
                    "Aktif Microsoft Project bulunamadı.");
            }

            bool oldScreenUpdating =
                app.ScreenUpdating;

            try
            {
                app.ScreenUpdating =
                    false;

                try
                {
                    app.Calculation =
                        MSProject.PjCalculation
                            .pjAutomatic;
                }
                catch
                {
                }

                // Project Operations ile aynı çalışma düzeni:
                // Pazartesi-Cuma 09:00-17:00.
                ConfigureProjectCalendar(
                    project);

                ClearProject(
                    project);

                SetProjectMetadata(
                    project,
                    schedule);

                Dictionary<Guid, Entity> bucketMap =
                    schedule.Buckets != null
                        ? schedule.Buckets.Entities
                            .ToDictionary(
                                x => x.Id)
                        : new Dictionary<Guid, Entity>();

                List<Entity> crmTasks =
                    schedule.Tasks != null
                        ? schedule.Tasks.Entities
                            .OrderBy(
                                x =>
                                    GetNumber(
                                        x,
                                        "msdyn_displaysequence")
                                    ?? decimal.MaxValue)
                            .ThenBy(
                                x => x.Id)
                            .ToList()
                        : new List<Entity>();

                SetProjectStart(
                    project,
                    schedule,
                    crmTasks);

                Dictionary<Guid, MSProject.Task> taskMap =
                    CreateTasks(
                        project,
                        crmTasks,
                        bucketMap);

                // Önce hierarchy.
                ApplyHierarchy(
                    crmTasks,
                    taskMap);

                // İstenen Task Mode:
                //
                // Summary = Auto Scheduled
                // Leaf    = Manually Scheduled
                ApplyTaskModes(
                    project,
                    taskMap);

                // Dependency'leri kur.
                ApplyDependencies(
                    schedule.Dependencies,
                    taskMap);

                // Resource'ları ekle.
                Dictionary<Guid, MSProject.Resource> resourceMap =
                    CreateResources(
                        project,
                        schedule);

                ApplyResourceAssignments(
                    schedule.ResourceAssignments,
                    taskMap,
                    resourceMap);

                // Önce Project mevcut yapıyı hesaplasın.
                try
                {
                    app.CalculateProject();
                }
                catch
                {
                }

                // CRM leaf task değerleri authoritative.
                //
                // Start
                // Finish
                // Duration
                // tt_percentcompleted
                //
                // birebir uygulanır.
                ApplyLeafValues(
                    project,
                    crmTasks,
                    taskMap);

                // Summary task'ları tekrar Auto garanti et.
                ApplyTaskModes(
                    project,
                    taskMap);

                // Project Summary Task da Auto.
                ConfigureProjectSummaryTask(
                    project);

                // Bir kez hesaplat.
                try
                {
                    app.CalculateProject();
                }
                catch
                {
                }

                // Calculate sırasında dependency/scheduling engine
                // leaf task'a dokunduysa CRM değerlerini SON KEZ
                // geri bas.
                ApplyLeafValues(
                    project,
                    crmTasks,
                    taskMap);

                // Summary task'ları yeniden Auto garanti et.
                ApplyTaskModes(
                    project,
                    taskMap);

                ConfigureProjectSummaryTask(
                    project);
            }
            finally
            {
                app.ScreenUpdating =
                    oldScreenUpdating;
            }
        }

        // =========================================================
        // PROJECT CALENDAR
        // =========================================================

        private static void ConfigureProjectCalendar(
            MSProject.Project project)
        {
            if (project == null)
                return;

            try
            {
                project.HoursPerDay =
                    8.0;

                project.HoursPerWeek =
                    40.0;
            }
            catch
            {
            }

            try
            {
                MSProject.Calendar calendar =
                    project.Calendar;

                if (calendar == null)
                    return;

                ConfigureWorkingDay(
                    calendar,
                    MSProject.PjWeekday.pjMonday);

                ConfigureWorkingDay(
                    calendar,
                    MSProject.PjWeekday.pjTuesday);

                ConfigureWorkingDay(
                    calendar,
                    MSProject.PjWeekday.pjWednesday);

                ConfigureWorkingDay(
                    calendar,
                    MSProject.PjWeekday.pjThursday);

                ConfigureWorkingDay(
                    calendar,
                    MSProject.PjWeekday.pjFriday);

                ConfigureNonWorkingDay(
                    calendar,
                    MSProject.PjWeekday.pjSaturday);

                ConfigureNonWorkingDay(
                    calendar,
                    MSProject.PjWeekday.pjSunday);
            }
            catch
            {
            }
        }

        private static void ConfigureWorkingDay(
            MSProject.Calendar calendar,
            MSProject.PjWeekday day)
        {
            try
            {
                MSProject.WeekDay weekDay =
                    calendar.WeekDays[day];

                weekDay.set_Working(
                    true);

                weekDay.Shift1.Start =
                    DateTime.Today
                        .AddHours(9);

                weekDay.Shift1.Finish =
                    DateTime.Today
                        .AddHours(17);

                try
                {
                    weekDay.Shift2.Clear();
                }
                catch
                {
                }

                try
                {
                    weekDay.Shift3.Clear();
                }
                catch
                {
                }
            }
            catch
            {
            }
        }

        private static void ConfigureNonWorkingDay(
            MSProject.Calendar calendar,
            MSProject.PjWeekday day)
        {
            try
            {
                MSProject.WeekDay weekDay =
                    calendar.WeekDays[day];

                weekDay.set_Working(
                    false);
            }
            catch
            {
            }
        }

        // =========================================================
        // PROJECT START
        // =========================================================

        private static void SetProjectStart(
            MSProject.Project project,
            CrmScheduleData schedule,
            List<Entity> crmTasks)
        {
            DateTime? projectStart =
                null;

            if (schedule != null &&
                schedule.Project != null)
            {
                projectStart =
                    GetDateTime(
                        schedule.Project,
                        "msdyn_tzascheduledstart")
                    ??
                    GetDateTime(
                        schedule.Project,
                        "msdyn_scheduledstart");
            }

            if (!projectStart.HasValue &&
                crmTasks != null &&
                crmTasks.Count > 0)
            {
                List<DateTime> starts =
                    new List<DateTime>();

                foreach (Entity source
                         in crmTasks)
                {
                    DateTime? start =
                        GetTaskStart(
                            source);

                    if (start.HasValue)
                    {
                        starts.Add(
                            start.Value);
                    }
                }

                if (starts.Count > 0)
                {
                    projectStart =
                        starts.Min();
                }
            }

            if (!projectStart.HasValue)
                return;

            try
            {
                project.ProjectStart =
                    NormalizeStart(
                        projectStart.Value);
            }
            catch
            {
            }
        }

        // =========================================================
        // PROJECT METADATA
        // =========================================================

        private static void SetProjectMetadata(
            MSProject.Project project,
            CrmScheduleData schedule)
        {
            try
            {
                MSProject.Task summary =
                    project.ProjectSummaryTask;

                if (summary == null)
                    return;

                if (schedule.Project != null)
                {
                    string projectName =
                        GetString(
                            schedule.Project,
                            "msdyn_subject");

                    if (!string.IsNullOrWhiteSpace(
                        projectName))
                    {
                        summary.Name =
                            projectName;
                    }
                }

                summary.Text30 =
                    schedule.ProjectId.ToString();

                summary.Text29 =
                    Globals.ThisAddIn.Crm != null
                        ? Globals.ThisAddIn.Crm.ConnectedUrl
                        : "";
            }
            catch
            {
            }
        }

        private static void ConfigureProjectSummaryTask(
            MSProject.Project project)
        {
            if (project == null)
                return;

            try
            {
                MSProject.Task summary =
                    project.ProjectSummaryTask;

                if (summary != null)
                {
                    summary.Manual =
                        false;
                }
            }
            catch
            {
            }
        }

        // =========================================================
        // CLEAR
        // =========================================================

        private static void ClearProject(
            MSProject.Project project)
        {
            if (project == null)
                return;

            for (int i = project.Tasks.Count;
                 i >= 1;
                 i--)
            {
                try
                {
                    MSProject.Task task =
                        project.Tasks[i];

                    if (task != null)
                    {
                        task.Delete();
                    }
                }
                catch
                {
                }
            }

            for (int i = project.Resources.Count;
                 i >= 1;
                 i--)
            {
                try
                {
                    MSProject.Resource resource =
                        project.Resources[i];

                    if (resource != null)
                    {
                        resource.Delete();
                    }
                }
                catch
                {
                }
            }
        }

        // =========================================================
        // CREATE TASKS
        // =========================================================

        private static Dictionary<Guid, MSProject.Task> CreateTasks(
            MSProject.Project project,
            List<Entity> crmTasks,
            Dictionary<Guid, Entity> bucketMap)
        {
            Dictionary<Guid, MSProject.Task> result =
                new Dictionary<Guid, MSProject.Task>();

            foreach (Entity source
                     in crmTasks)
            {
                string name =
                    GetString(
                        source,
                        "msdyn_subject");

                if (string.IsNullOrWhiteSpace(
                    name))
                {
                    name =
                        "(İsimsiz Görev)";
                }

                MSProject.Task task =
                    project.Tasks.Add(
                        name,
                        Type.Missing);

                // CRM Task ID
                task.Text30 =
                    source.Id.ToString();

                // =================================================
                // BUCKET
                // =================================================

                EntityReference bucket =
                    GetReference(
                        source,
                        "msdyn_projectbucket");

                if (bucket != null)
                {
                    task.Text29 =
                        bucket.Id.ToString();

                    Entity bucketEntity;

                    if (bucketMap.TryGetValue(
                        bucket.Id,
                        out bucketEntity))
                    {
                        task.Text27 =
                            GetString(
                                bucketEntity,
                                "msdyn_name");
                    }
                }

                // CRM Client ID
                task.Text28 =
                    GetString(
                        source,
                        "msdyn_msprojectclientid");

                // =================================================
                // DESCRIPTION
                // =================================================

                string description =
                    GetString(
                        source,
                        "msdyn_description");

                if (!string.IsNullOrWhiteSpace(
                    description))
                {
                    try
                    {
                        task.Notes =
                            description;
                    }
                    catch
                    {
                    }
                }

                // =================================================
                // FLAGS
                // =================================================

                try
                {
                    task.Flag19 =
                        GetBoolean(
                            source,
                            "msdyn_summary")
                        ?? false;
                }
                catch
                {
                }

                bool? critical =
                    GetBoolean(
                        source,
                        "msdyn_critical");

                if (critical.HasValue)
                {
                    try
                    {
                        task.Flag20 =
                            critical.Value;
                    }
                    catch
                    {
                    }
                }

                // =================================================
                // CRM TECHNICAL DATA
                // =================================================

                SetNumberField(
                    delegate
                    {
                        task.Number20 =
                            ToDouble(
                                GetNumber(
                                    source,
                                    "msdyn_effort"));
                    });

                SetNumberField(
                    delegate
                    {
                        task.Number19 =
                            ToDouble(
                                GetNumber(
                                    source,
                                    "msdyn_effortcompleted"));
                    });

                SetNumberField(
                    delegate
                    {
                        task.Number18 =
                            ToDouble(
                                GetNumber(
                                    source,
                                    "msdyn_effortremaining"));
                    });

                SetNumberField(
                    delegate
                    {
                        task.Number17 =
                            ToDouble(
                                GetNumber(
                                    source,
                                    "msdyn_duration"));
                    });

                SetNumberField(
                    delegate
                    {
                        task.Number16 =
                            ToDouble(
                                GetNumber(
                                    source,
                                    "msdyn_displaysequence"));
                    });

                SetNumberField(
                    delegate
                    {
                        task.Number15 =
                            ToDouble(
                                GetNumber(
                                    source,
                                    "msdyn_outlinelevel"));
                    });

                result[
                    source.Id] =
                    task;
            }

            return result;
        }

        // =========================================================
        // HIERARCHY
        // =========================================================

        private static void ApplyHierarchy(
            List<Entity> crmTasks,
            Dictionary<Guid, MSProject.Task> taskMap)
        {
            foreach (Entity source
                     in crmTasks)
            {
                EntityReference parentReference =
                    GetReference(
                        source,
                        "msdyn_parenttask");

                if (parentReference == null)
                    continue;

                MSProject.Task child;
                MSProject.Task parent;

                if (!taskMap.TryGetValue(
                        source.Id,
                        out child) ||
                    !taskMap.TryGetValue(
                        parentReference.Id,
                        out parent))
                {
                    continue;
                }

                try
                {
                    int desiredLevel =
                        parent.OutlineLevel + 1;

                    int guard =
                        0;

                    while (child.OutlineLevel <
                               desiredLevel &&
                           guard < 20)
                    {
                        child.OutlineIndent();

                        guard++;
                    }
                }
                catch
                {
                }
            }
        }

        // =========================================================
        // TASK MODE
        //
        // SUMMARY = AUTO
        // LEAF    = MANUAL
        // =========================================================

        private static void ApplyTaskModes(
            MSProject.Project project,
            Dictionary<Guid, MSProject.Task> taskMap)
        {
            foreach (KeyValuePair<Guid, MSProject.Task> pair
                     in taskMap)
            {
                MSProject.Task task =
                    pair.Value;

                if (task == null)
                    continue;

                try
                {
                    if (task.Summary)
                    {
                        task.Manual =
                            false;
                    }
                    else
                    {
                        task.Manual =
                            true;

                        // Resource ataması duration'ı değiştirmesin.
                        task.Type =
                            MSProject.PjTaskFixedType
                                .pjFixedDuration;
                    }
                }
                catch
                {
                }
            }

            ConfigureProjectSummaryTask(
                project);
        }

        // =========================================================
        // EXACT CRM LEAF VALUES
        //
        // CRM:
        //
        // Start
        // Finish
        // Duration
        // tt_percentcompleted
        //
        // -> Project leaf task birebir.
        //
        // Summary alanlarına DOKUNMUYORUZ.
        // Summary'yi Project hesaplıyor.
        // =========================================================

        private static void ApplyLeafValues(
            MSProject.Project project,
            List<Entity> crmTasks,
            Dictionary<Guid, MSProject.Task> taskMap)
        {
            double hoursPerDay =
                8.0;

            try
            {
                if (project.HoursPerDay > 0)
                {
                    hoursPerDay =
                        project.HoursPerDay;
                }
            }
            catch
            {
            }

            foreach (Entity source
                     in crmTasks)
            {
                MSProject.Task task;

                if (!taskMap.TryGetValue(
                        source.Id,
                        out task) ||
                    task == null)
                {
                    continue;
                }

                bool isSummary =
                    false;

                try
                {
                    isSummary =
                        task.Summary;
                }
                catch
                {
                }

                // Summary:
                //
                // Start
                // Finish
                // Duration
                // Percent Complete
                //
                // Project tarafından roll-up.
                if (isSummary)
                {
                    try
                    {
                        task.Manual =
                            false;
                    }
                    catch
                    {
                    }

                    continue;
                }

                // Leaf = Manual.
                try
                {
                    task.Manual =
                        true;

                    task.Type =
                        MSProject.PjTaskFixedType
                            .pjFixedDuration;
                }
                catch
                {
                }

                DateTime? crmStart =
                    GetTaskStart(
                        source);

                DateTime? crmFinish =
                    GetTaskFinish(
                        source);

                decimal? durationDays =
                    GetNumber(
                        source,
                        "msdyn_duration");

                bool milestone =
                    GetBoolean(
                        source,
                        "msdyn_ismilestone")
                    ?? false;

                if (durationDays.HasValue &&
                    durationDays.Value == 0m)
                {
                    milestone =
                        true;
                }

                // -------------------------------------------------
                // MILESTONE
                // -------------------------------------------------

                try
                {
                    task.Milestone =
                        milestone;
                }
                catch
                {
                }

                // -------------------------------------------------
                // DURATION
                //
                // CRM gün -> Project dakika
                // -------------------------------------------------

                if (milestone)
                {
                    try
                    {
                        task.Duration =
                            0d;
                    }
                    catch
                    {
                    }
                }
                else if (durationDays.HasValue)
                {
                    try
                    {
                        double days =
                            Convert.ToDouble(
                                durationDays.Value,
                                CultureInfo.InvariantCulture);

                        if (days < 0)
                        {
                            days =
                                0;
                        }

                        task.Duration =
                            days *
                            hoursPerDay *
                            60.0;
                    }
                    catch
                    {
                    }
                }

                // -------------------------------------------------
                // START
                //
                // CRM Grid tarihini birebir korumak için
                // timezone-agnostic tarih kullanılır.
                //
                // Çalışma başlangıcı 09:00.
                // -------------------------------------------------

                if (crmStart.HasValue)
                {
                    try
                    {
                        task.Start =
                            NormalizeStart(
                                crmStart.Value);
                    }
                    catch
                    {
                    }
                }

                // -------------------------------------------------
                // FINISH
                //
                // CRM Grid tarihini birebir koruyoruz.
                //
                // Normal task: 17:00
                // Milestone    : 09:00
                // -------------------------------------------------

                if (crmFinish.HasValue)
                {
                    try
                    {
                        task.Finish =
                            milestone
                                ? NormalizeStart(
                                    crmFinish.Value)
                                : NormalizeFinish(
                                    crmFinish.Value);
                    }
                    catch
                    {
                    }
                }

                // -------------------------------------------------
                // % COMPLETE
                //
                // CRM tt_percentcompleted
                // ->
                // Project built-in % Complete
                // -------------------------------------------------

                int? percent =
                    GetPercent(
                        source,
                        PercentCompletedField);

                if (percent.HasValue)
                {
                    try
                    {
                        task.PercentComplete =
                            percent.Value;
                    }
                    catch
                    {
                    }
                }
            }
        }

        // =========================================================
        // CRM DATE HELPERS
        // =========================================================

        private static DateTime? GetTaskStart(
            Entity source)
        {
            return
                GetDateTime(
                    source,
                    "msdyn_tzascheduledstart")
                ??
                GetDateTime(
                    source,
                    "msdyn_start")
                ??
                GetDateTime(
                    source,
                    "msdyn_scheduledstart");
        }

        private static DateTime? GetTaskFinish(
            Entity source)
        {
            return
                GetDateTime(
                    source,
                    "msdyn_tzascheduledend")
                ??
                GetDateTime(
                    source,
                    "msdyn_finish")
                ??
                GetDateTime(
                    source,
                    "msdyn_scheduledend");
        }

        private static DateTime NormalizeStart(
            DateTime value)
        {
            return
                value.Date
                    .AddHours(9);
        }

        private static DateTime NormalizeFinish(
            DateTime value)
        {
            return
                value.Date
                    .AddHours(17);
        }

        // =========================================================
        // DEPENDENCIES
        // =========================================================

        private static void ApplyDependencies(
            EntityCollection dependencies,
            Dictionary<Guid, MSProject.Task> taskMap)
        {
            if (dependencies == null)
                return;

            foreach (Entity dependency
                     in dependencies.Entities)
            {
                EntityReference predecessorReference =
                    GetReference(
                        dependency,
                        "msdyn_predecessortask");

                EntityReference successorReference =
                    GetReference(
                        dependency,
                        "msdyn_successortask");

                if (predecessorReference == null ||
                    successorReference == null)
                {
                    continue;
                }

                MSProject.Task predecessor;
                MSProject.Task successor;

                if (!taskMap.TryGetValue(
                        predecessorReference.Id,
                        out predecessor) ||
                    !taskMap.TryGetValue(
                        successorReference.Id,
                        out successor))
                {
                    continue;
                }

                try
                {
                    successor.LinkPredecessors(
                        predecessor,
                        GetProjectLinkType(
                            dependency),
                        Type.Missing);
                }
                catch
                {
                }
            }
        }

        private static MSProject.PjTaskLinkType GetProjectLinkType(
            Entity entity)
        {
            OptionSetValue value =
                entity.GetAttributeValue<OptionSetValue>(
                    "msdyn_linktype");

            int raw =
                value != null
                    ? value.Value
                    : 192350000;

            switch (raw)
            {
                case 192350001:

                    return
                        MSProject.PjTaskLinkType
                            .pjStartToStart;

                case 192350002:

                    return
                        MSProject.PjTaskLinkType
                            .pjFinishToFinish;

                case 192350003:
                case 192350004:

                    return
                        MSProject.PjTaskLinkType
                            .pjStartToFinish;

                default:

                    return
                        MSProject.PjTaskLinkType
                            .pjFinishToStart;
            }
        }

        // =========================================================
        // RESOURCES
        // =========================================================

        private static Dictionary<Guid, MSProject.Resource> CreateResources(
            MSProject.Project project,
            CrmScheduleData schedule)
        {
            Dictionary<Guid, MSProject.Resource> result =
                new Dictionary<Guid, MSProject.Resource>();

            Dictionary<Guid, string> bookableNames =
                new Dictionary<Guid, string>();

            if (schedule.BookableResources != null)
            {
                foreach (Entity resource
                         in schedule.BookableResources.Entities)
                {
                    bookableNames[
                        resource.Id] =
                        GetString(
                            resource,
                            "name");
                }
            }

            if (schedule.TeamMembers == null)
                return result;

            foreach (Entity teamMember
                     in schedule.TeamMembers.Entities)
            {
                EntityReference bookable =
                    GetReference(
                        teamMember,
                        "msdyn_bookableresourceid");

                string resourceName =
                    null;

                if (bookable != null)
                {
                    bookableNames.TryGetValue(
                        bookable.Id,
                        out resourceName);
                }

                if (string.IsNullOrWhiteSpace(
                    resourceName))
                {
                    resourceName =
                        GetString(
                            teamMember,
                            "msdyn_name");
                }

                if (string.IsNullOrWhiteSpace(
                    resourceName))
                {
                    resourceName =
                        "Resource";
                }

                MSProject.Resource projectResource =
                    project.Resources.Add(
                        resourceName,
                        Type.Missing);

                projectResource.Text29 =
                    teamMember.Id.ToString();

                if (bookable != null)
                {
                    projectResource.Text30 =
                        bookable.Id.ToString();
                }

                result[
                    teamMember.Id] =
                    projectResource;
            }

            return result;
        }

        // =========================================================
        // RESOURCE ASSIGNMENTS
        // =========================================================

        private static void ApplyResourceAssignments(
            EntityCollection assignments,
            Dictionary<Guid, MSProject.Task> taskMap,
            Dictionary<Guid, MSProject.Resource> resourceMap)
        {
            if (assignments == null)
                return;

            foreach (Entity source
                     in assignments.Entities)
            {
                EntityReference taskReference =
                    GetReference(
                        source,
                        "msdyn_taskid");

                EntityReference teamReference =
                    GetReference(
                        source,
                        "msdyn_projectteamid");

                if (taskReference == null ||
                    teamReference == null)
                {
                    continue;
                }

                MSProject.Task task;
                MSProject.Resource resource;

                if (!taskMap.TryGetValue(
                        taskReference.Id,
                        out task) ||
                    !resourceMap.TryGetValue(
                        teamReference.Id,
                        out resource))
                {
                    continue;
                }

                try
                {
                    MSProject.Assignment assignment =
                        task.Assignments.Add(
                            Type.Missing,
                            resource.ID,
                            Type.Missing);

                    if (assignment == null)
                        continue;

                    assignment.Text30 =
                        source.Id.ToString();

                    assignment.Text29 =
                        teamReference.Id.ToString();

                    assignment.Text28 =
                        NumberText(
                            source,
                            "msdyn_effort");

                    assignment.Text27 =
                        NumberText(
                            source,
                            "msdyn_effortcompleted");

                    assignment.Text26 =
                        NumberText(
                            source,
                            "msdyn_effortremaining");

                    DateTime? start =
                        GetDateTime(
                            source,
                            "msdyn_start");

                    DateTime? finish =
                        GetDateTime(
                            source,
                            "msdyn_finish");

                    assignment.Text25 =
                        start.HasValue
                            ? start.Value.ToString(
                                "o",
                                CultureInfo.InvariantCulture)
                            : "";

                    assignment.Text24 =
                        finish.HasValue
                            ? finish.Value.ToString(
                                "o",
                                CultureInfo.InvariantCulture)
                            : "";
                }
                catch
                {
                }
            }
        }

        // =========================================================
        // HELPERS
        // =========================================================

        private static int? GetPercent(
            Entity entity,
            string field)
        {
            decimal? value =
                GetNumber(
                    entity,
                    field);

            if (!value.HasValue)
                return null;

            int result =
                Convert.ToInt32(
                    Math.Round(
                        value.Value,
                        0,
                        MidpointRounding.AwayFromZero));

            if (result < 0)
            {
                result =
                    0;
            }

            if (result > 100)
            {
                result =
                    100;
            }

            return result;
        }

        private static double ToDouble(
            decimal? value)
        {
            if (!value.HasValue)
            {
                return 0d;
            }

            return Convert.ToDouble(
                value.Value,
                CultureInfo.InvariantCulture);
        }

        private static string NumberText(
            Entity entity,
            string field)
        {
            decimal? number =
                GetNumber(
                    entity,
                    field);

            return number.HasValue
                ? number.Value.ToString(
                    CultureInfo.InvariantCulture)
                : "";
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

            return Convert.ToString(
                       entity[field],
                       CultureInfo.InvariantCulture)
                   ?? "";
        }

        private static EntityReference GetReference(
            Entity entity,
            string field)
        {
            if (entity == null)
                return null;

            return entity
                .GetAttributeValue<EntityReference>(
                    field);
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
                CultureInfo.InvariantCulture,
                DateTimeStyles.None,
                out value))
            {
                return value;
            }

            return null;
        }

        private static bool? GetBoolean(
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

            if (value is bool)
            {
                return
                    (bool)value;
            }

            bool result;

            if (bool.TryParse(
                value.ToString(),
                out result))
            {
                return result;
            }

            return null;
        }

        private static decimal? GetNumber(
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
                        ((Money)value).Value;
                }

                if (value is OptionSetValue)
                {
                    return
                        ((OptionSetValue)value).Value;
                }

                return Convert.ToDecimal(
                    value,
                    CultureInfo.InvariantCulture);
            }
            catch
            {
                return null;
            }
        }

        private static void SetNumberField(
            Action setter)
        {
            try
            {
                setter();
            }
            catch
            {
            }
        }
    }
}