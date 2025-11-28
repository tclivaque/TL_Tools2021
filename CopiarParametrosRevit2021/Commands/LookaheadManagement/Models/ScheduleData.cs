using System.Collections.Generic;

namespace CopiarParametrosRevit2021.Commands.LookaheadManagement.Models
{
    public class ScheduleData
    {
        public string ActivityName { get; set; }
        public ActivityRule MatchingRule { get; set; }
        public List<ScheduleGroup> Groups { get; set; } = new List<ScheduleGroup>();
    }
}