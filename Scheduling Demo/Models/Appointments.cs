using System.ComponentModel.DataAnnotations;

namespace Scheduling_Demo.Models
{
    public class Appointments
    {

        public string Facility { get; set; }

        [Display(Name = "MR#")]
        public string MRNumber { get; set; }

        [Display(Name = "Group Topic")]
        public string GroupTopic { get; set; }

        
        public string LOC { get; set; }

        [Display(Name = "Checked In?")]
        public bool CheckedIn { get; set; }

        [Display(Name = "Group Date")]
        public string GroupDate { get; set; }
    }
}