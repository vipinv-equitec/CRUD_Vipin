using ClosedXML.Excel;
using CRUDWithDapper.Models;
using CRUDWithDapper.Repository;
using PagedList;
using System.IO;
using System.Linq;
using System.Web.Mvc;

namespace CRUDWithDapper.Controllers
{
    public class StudentController : Controller
    {
        private StudentRepository _studentRepository;
        public StudentController()
        {
            string c = @"Data Source=DESKTOP-FMGSCAM;Initial Catalog=student;User Id=sa;Password=123";
            _studentRepository = new StudentRepository(c);
        }
        public ActionResult GetAllStudents(int page = 1, int pageSize = 5)
        {
            if (Request.QueryString["pageSize"] != null && int.TryParse(Request.QueryString["pageSize"], out int selectedPageSize))
            {
                pageSize = selectedPageSize;
            }
            var students = _studentRepository.GetStudents(page, pageSize);
            var totalStudentCount = _studentRepository.GetTotalStudentCount();
            var pagedList = new StaticPagedList<Student>(students, page, pageSize, totalStudentCount);
            return View(pagedList);
        }
        public ActionResult GetAllStudent()
        {
            return RedirectToAction("GetAllStudents");
        }
        public ActionResult AddStudent()
        {
            return View();
        }
        [HttpPost]
        public ActionResult AddStudent(Student student)
        {
            try
            {
                if (ModelState.IsValid)
                {
                    var selectedSkills = Request.Form.GetValues("SelectedSkills");
                    student.Skills = selectedSkills != null ? string.Join(", ", selectedSkills) : string.Empty;
                    _studentRepository.AddStudent(student.StudentId, student.StudName, student.StudAge, student.StudEmail, student.StudDepartment, student.Skills);
                    ViewBag.Message = "Records added successfully.";
                }
                return RedirectToAction("GetAllStudent");
            }
            catch
            {
                return View();
            }
        }
        public ActionResult UpdateStudent(int id)
        {
            // return View(_studentRepository.GetAllStudent().Find(Student => Student.StudentId == id));
            var student = _studentRepository.GetAllStudent().Find(s => s.StudentId == id);
            student.SelectedSkills = student.Skills?.Split(',').ToList();
            return View(student);
        }
        [HttpPost]
        public ActionResult UpdateStudent(int id, Student student)
        {
            try
            {
                // _studentRepository.UpdateStudent(student);
                // return RedirectToAction("GetAllStudent");
                if (ModelState.IsValid)
                {
                    var selectedSkills = Request.Form.GetValues("SelectedSkills");
                    student.Skills = selectedSkills != null ? string.Join(", ", selectedSkills) : string.Empty;

                    _studentRepository.UpdateStudent(student.StudentId, student.StudName, student.StudAge, student.StudEmail, student.StudDepartment, student.Skills);

                    return RedirectToAction("GetAllStudent");
                }
                return View(student);
            }
            catch
            {
                return View();
            }
        }
        public ActionResult DeleteStudentFromList(int id)
        {
            try
            {
                _studentRepository.DeleteStudent(id);
                ViewBag.AlertMsg = "Employee details deleted successfully";
                return RedirectToAction("GetAllStudents");
            }
            catch
            {
                return RedirectToAction("GetAllStudents");
            }
        }
        public ActionResult DetailStudent(int id)
        {
            return View(_studentRepository.GetAllStudent().Find(Student => Student.StudentId == id));
        }
        public ActionResult DeleteSoft(int id)
        {
            try
            {
                _studentRepository.DeleteSoft(id);
                ViewBag.AlertMsg = "Employee details deleted successfully";
                return RedirectToAction("GetAllStudent");
            }
            catch
            {
                return RedirectToAction("GetAllStudent");
            }
        }
        public ActionResult RestoreStudent()
        {
            return View(_studentRepository.RestoreStudent());
        }
        public ActionResult BackStudent(int id)
        {
            try
            {
                _studentRepository.BackStudent(id);
                ViewBag.AlertMsg = "Restore back successfully";
                return RedirectToAction("GetAllStudent");
            }
            catch
            {
                return RedirectToAction("GetAllStudent");
            }
        }
        public ActionResult DownloadExcel()
        {
            var products = _studentRepository.GetAllStudent();
            using (var workbook = new XLWorkbook())
            {
                var worksheet = workbook.Worksheets.Add("Student");
                var properties = typeof(Student).GetProperties();
                for (int i = 0; i < properties.Length; i++)
                {
                    worksheet.Cell(1, i + 1).Value = properties[i].Name;
                }
                for (int row = 2; row <= products.Count() + 1; row++)
                {
                    for (int col = 1; col <= properties.Length; col++)
                    {
                        var propertyValue = properties[col - 1].GetValue(products.ElementAt(row - 2));
                        worksheet.Cell(row, col).Value = propertyValue != null ? propertyValue.ToString() : null;
                    }
                }
                var memoryStream = new MemoryStream();
                workbook.SaveAs(memoryStream);
                return File(memoryStream.ToArray(), "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet", "Student.xlsx");
            }
        }
        public ActionResult PrintData()
        {
            string htmlContent = _studentRepository.PrintData().ToString();
            htmlContent += "<script>window.onload = function() { window.print();}</script>";
            return Content(htmlContent, "text/html");
        }
    }
}
