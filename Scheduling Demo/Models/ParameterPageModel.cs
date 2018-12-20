using System;
using System.Collections.Generic;
using System.ComponentModel.DataAnnotations;
using System.Web.Mvc;

namespace Scheduling_Demo.Models
{
    public class ParameterPageModel
    {
        [Required]
        public List<SelectListItem> Facility { get; set; }
        [Display(Name = "Facility")]
        public string SelectedFacility { get; set; }

        public List<SelectListItem> GroupTopics { get; set; }
        [Display(Name = "Group Topic")]
        public string SelectedGroupTopic { get; set; }


        
        [DataType(DataType.DateTime, ErrorMessage ="Please enter a date time in the format of MM/DD/YYYY")]
        [RegularExpression(@"^((0|1)\d{1})\/((0|1|2)\d{1})\/((19|20)\d{2})", ErrorMessage = "Date must be in the format of MM/DD/YYYY")]
        [Required(AllowEmptyStrings = false, ErrorMessage = "Group Date Required")]
        [Display(Name = "Group Date")]
        public DateTime? GroupDate { get; set; }

        public List<SelectListItem> LOCs { get; set; }
        [Display(Name = "LOC")]
        public string SelectedLOC { get; set; }
    }
}