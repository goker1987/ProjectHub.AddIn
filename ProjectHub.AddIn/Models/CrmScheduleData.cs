using Microsoft.Xrm.Sdk;
using System;

namespace ProjectHub.AddIn
{
    public class CrmScheduleData
    {
        public Guid ProjectId { get; set; }

        public Entity Project { get; set; }

        public EntityCollection Tasks { get; set; }

        public EntityCollection Dependencies { get; set; }

        public EntityCollection ResourceAssignments { get; set; }

        public EntityCollection TeamMembers { get; set; }

        public EntityCollection BookableResources { get; set; }

        public EntityCollection Buckets { get; set; }

        public CrmScheduleData()
        {
            Tasks = new EntityCollection();
            Dependencies = new EntityCollection();
            ResourceAssignments = new EntityCollection();
            TeamMembers = new EntityCollection();
            BookableResources = new EntityCollection();
            Buckets = new EntityCollection();
        }
    }
}