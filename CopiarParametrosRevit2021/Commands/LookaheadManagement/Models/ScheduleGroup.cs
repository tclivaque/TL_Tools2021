using System.Collections.Generic;

namespace TL_Tools2021.Commands.LookaheadManagement.Models
{
    public class ScheduleGroup
    {
        public bool IsSitioLogic { get; set; }
        public string Sector { get; set; }
        public string Nivel { get; set; }
        public Dictionary<string, int> WeekCounts { get; set; }

        public ScheduleGroup()
        {
            WeekCounts = new Dictionary<string, int>();
        }
    }
}
