using System;
using System.Collections.Generic;
using System.ComponentModel.DataAnnotations;
using System.Linq;
using System.Web;

namespace CRUDWithDapper.Models
{
    public class Student
    {
        [Display(Name = "Id")]
        public int StudentId { get; set; }
        [Required(ErrorMessage = "First name is required.")]
        public string StudName { get; set; }
        [Required(ErrorMessage = "Age is required.")]
        public int StudAge { get; set; }
        [EmailAddress(ErrorMessage = "Email is required.")]
        public string StudEmail { get; set; }
        [Required(ErrorMessage = "Department name is required.")]
        public string StudDepartment { get; set; }
        public List<string> SelectedSkills { get; set; }
        //[Required(ErrorMessage = "Skills is required.")]
        public string Skills { get; set; }
    }
}