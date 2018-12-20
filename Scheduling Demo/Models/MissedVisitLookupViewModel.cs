using System;
using System.ComponentModel.DataAnnotations;

namespace Scheduling_Demo.Models
{
    public class MissedVisitLookupViewModel
    {
        [DataType(DataType.DateTime, ErrorMessage = "Please enter a date time in the format of MM/DD/YYYY")]
        [RegularExpression(@"^((0|1)\d{1})\/((0|1|2)\d{1})\/((19|20)\d{2})", ErrorMessage = "Date must be in the format of MM/DD/YYYY")]
        [Required(AllowEmptyStrings = false, ErrorMessage = "Group Date Required")]
        [Display(Name = "Group Date")]
        public DateTime? MissedVisitDate { get; set; }
    }
}