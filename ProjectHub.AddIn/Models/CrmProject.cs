using System;

namespace ProjectHub.AddIn
{
    public class CrmProject
    {
        public Guid Id { get; set; }

        public string Name { get; set; }

        public string ProjectManager { get; set; }

        public DateTime? Start { get; set; }

        public DateTime? Finish { get; set; }

        public int? PercentComplete { get; set; }

        public override string ToString()
        {
            return Name ?? "";
        }
    }
}