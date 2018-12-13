using System.ComponentModel;

namespace Scheduling_Demo.Models
{
    public class Appointments
    {

        public string Facility { get; set; }
        [DisplayName("MR#")]
        public string MRNumber { get; set; }

        [DisplayName("Group Topic")]
        public string GroupTopic { get; set; }

        public string Date { get; set; }
        public string LOC { get; set; }
        public int Episode { get; set; }

        [DisplayName("Checked In?")]
        public bool CheckedIn { get; set; }
    }
}