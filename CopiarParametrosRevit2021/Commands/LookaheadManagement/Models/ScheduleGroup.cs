using System.Collections.Generic;

namespace CopiarParametrosRevit2021.Commands.LookaheadManagement.Models
{
    public class ScheduleGroup
    {
        public string Sector { get; set; }
        public string Nivel { get; set; }
        public bool IsSitioLogic { get; set; }
        public Dictionary<string, int> WeekCounts { get; set; } = new Dictionary<string, int>();
    }
}