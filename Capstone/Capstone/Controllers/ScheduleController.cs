using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;
using System.Web.Mvc;
using System.Data.OleDb;
using System.Threading.Tasks;
using System.IO;
using System.Data;
using Newtonsoft.Json;
using System.Diagnostics;
using Capstone.Classes.GeneratorClasses;
using Capstone.Classes;
using ClosedXML.Excel;

namespace Capstone.Controllers
{
    public class ScheduleController : Controller
    {
        // GET: ScheduleGenerator
        public ActionResult Index()
        {
            //If for ever any reason we return here ... just go back to the main page!
            return RedirectToAction("Index", "Home", null);
        }

        [HttpPost]
        public ActionResult ImportExcelFromFile(HttpPostedFileBase sheetFile)
        {
            if (sheetFile == null || sheetFile.ContentLength == 0)
            {
                TempData["UploadError"] = "You must upload a file.";
            }
            if (sheetFile.FileName.EndsWith(".xls") || sheetFile.FileName.EndsWith(".xlsx"))
            {
                TempData["UploadSuccess"] = "Upload successful!";
                try
                {
                    string path = AppDomain.CurrentDomain.GetData("DataDirectory").ToString() + "/sheetStorage/" + sheetFile.FileName;
                    if (System.IO.File.Exists(path))
                        System.IO.File.Delete(path);
                    sheetFile.SaveAs(path);

                    APIManager drive = new APIManager(Session["USER"].ToString());
                    string returnURL = Task.Run(() => drive.UploadSheet(HttpContext, path, sheetFile.FileName)).Result;

                    ProcessSchedule(path, sheetFile.FileName, returnURL);
                }
                catch (Exception e)
                {
                    TempData["UploadError"] = "Please upload a valid file.";
                }
            }
            else
            {
                TempData["UploadError"] = "Please upload a valid file.";
            }
            return RedirectToAction("History", "Home", null);
        }
        [HttpPost]
        public ActionResult ImportExcelFromDrive(string driveItemID, string fileName, string returnURL)
        {
            string path = AppDomain.CurrentDomain.GetData("DataDirectory").ToString() + "/sheetStorage/" + fileName;
            if (System.IO.File.Exists(path))
                System.IO.File.Delete(path);

            APIManager drive = new APIManager(Session["USER"].ToString());
            var task = Task.Run(() => drive.DownloadSheet(HttpContext, driveItemID, path)).Result;

            ProcessSchedule(path, fileName, returnURL);

            return RedirectToAction("History", "Home", null);
        }

        public void ProcessSchedule(string path, string fileName, string fileURL)
        {
            // Connect to recently saved excel sheet
            OleDbConnection conn = null;
            if (fileName.EndsWith(".xls"))
                conn = new OleDbConnection(@"Provider=Microsoft.Jet.OLEDB.4.0; data source=" + path + ";Extended Properties=\"Excel 8.0;HDR=1;IMEX=1\";");
            if (fileName.EndsWith(".xlsx"))
                conn = new OleDbConnection(@"Provider=Microsoft.ACE.OLEDB.12.0;Data Source=" + path + ";Extended Properties=\"Excel 12.0 Xml;HDR=1;IMEX=1\";");

            conn.Open();
            DataTable data = conn.GetOleDbSchemaTable(OleDbSchemaGuid.Tables, null);
            string sheet = data.Rows[0]["Table_Name"].ToString();

            OleDbCommand sheetCommand = new OleDbCommand(@"SELECT * FROM [" + sheet + @"]", conn);
            OleDbDataAdapter sheetAdapter = new OleDbDataAdapter(sheetCommand);

            DataSet sheetData = new DataSet();
            sheetAdapter.Fill(sheetData);
            conn.Close();

            string examURL = "";
            string calendarURL = "";
            try
            {
                Object[] examSchedule = GenerateExamSchedule(sheetData);
                SaveExamSheet(examSchedule, fileName);

                APIManager drive = new APIManager(Session["USER"].ToString());

                System.IO.File.Delete(path);
                path = AppDomain.CurrentDomain.GetData("DataDirectory").ToString() + "/sheetStorage/Exam_" + fileName;

                examURL = Task.Run(() => drive.UploadSheet(HttpContext, path, "Exam_" + fileName, true)).Result;
                calendarURL = Task.Run(() => drive.CreateCalendarEvent(HttpContext, "Exam_" + fileName, (List<Exam>)examSchedule[0])).Result;

                System.IO.File.Delete(path);

            } catch(Exception e)
            {
                TempData["ERROR"] = "The provided file was incompatible. It may be missing required value columns. Please refer to the template sheet for comparison!";
            }

            DBManager db = new DBManager();
            db.NewHistoryEntry(fileName, fileURL, "Exam_"+ fileName, examURL, calendarURL, Session["USER"].ToString());
        }

        public Object[] GenerateExamSchedule(DataSet classData)
        {
            // Data Ref Points
            // Indexing ensures that any sheet will function appropriately even if the headers aren't exact, so long as
            // they're in the right place. In future development, we could use a series of value comparisons to dynamically
            // assign indexes to allow a sheet to have varying column placement. Quality of life.
            int CODE = 0;
            int SECTION = 1;
            int NAME = 2;
            int FACULTY = 3;
            int ROOM = 4;
            int DAY = 5;
            int START = 6;
            int END = 7;
            int DUR = 8;

            // weekdays logs available rooms in a series of dictionary layers to optimizate iteration patterns.
            // [Day][StartTime][Room, EndTime]
            Dictionary<string, Dictionary<string, Dictionary<string, string>>>
                rooms = new Dictionary<string, Dictionary<string, Dictionary<string, string>>>();
            // classes logs every unique class by program code, and then by section.
            // [Code][Section][Name, FACULTY, Duration]
            Dictionary<string, Dictionary<string, Dictionary<string, string>>>
                classes = new Dictionary<string, Dictionary<string, Dictionary<string, string>>>();
            // faculty is a list of all faculty discovered in the submission sheet
            List<string> faculty = new List<string>();
            // facultySchedule keeps track of all available faculty by for every weekday.
            // [Day][Name, Available]
            Dictionary<string, Dictionary<string, DateTime>> facultySchedule = new Dictionary<string, Dictionary<string, DateTime>>();

            foreach (DataRow row in classData.Tables[0].Rows)
            {
                // Establish Room Slot
                if (!rooms.ContainsKey(row[DAY].ToString()))
                    rooms[row[DAY].ToString()] = new Dictionary<string, Dictionary<string, string>>();
                if (!rooms[row[DAY].ToString()].ContainsKey(row[START].ToString()))
                    rooms[row[DAY].ToString()][row[START].ToString()] = new Dictionary<string, string>();
                rooms[row[DAY].ToString()][row[START].ToString()][row[ROOM].ToString()] = row[END].ToString();
                // Establish Classes
                if (!classes.ContainsKey(row[CODE].ToString()))
                    classes[row[CODE].ToString()] = new Dictionary<string, Dictionary<string, string>>();
                if (!classes[row[CODE].ToString()].ContainsKey(row[SECTION].ToString()))
                    classes[row[CODE].ToString()][row[SECTION].ToString()] = new Dictionary<string, string>()
                    {
                        { "NAME", row[NAME].ToString() },
                        { "FACULTY", row[FACULTY].ToString() },
                        { "DUR", row[DUR].ToString() }
                    };
                if (!faculty.Contains(row[FACULTY].ToString()))
                    faculty.Add(row[FACULTY].ToString());
                // Establish faculty schedule weekdays
                if (!facultySchedule.ContainsKey(row[DAY].ToString())) { }
                    facultySchedule[row[DAY].ToString()] = new Dictionary<string, DateTime>();
            }
            // Develop list of faculty within scheduling weekday
            List<string> scheduleKeys = new List<string>(facultySchedule.Keys);
            foreach (string weekday in scheduleKeys)
            {
                foreach (string person in faculty)
                    facultySchedule[weekday].Add(person, new DateTime());
            }
            int iteration = 0;

            List<Exam> exams = new List<Exam>();
            List<string> classKeys = new List<string>(classes.Keys);
            List<string> roomKeys = new List<string>(rooms.Keys);

            // This loop restricts exams from being placed in time slots that might need to accomodate larger section splits
            while (classKeys.Count > 0 && iteration <= 1)
            {
                foreach (string progCode in classKeys)
                {
                    // All weekdays.
                    foreach (string weekday in roomKeys)
                    {
                        // All time entries within a week day. For example, 12:00. Contains all rooms and end times for rooms that start at 12.
                        List<string> timeKeys = new List<string>(rooms[weekday].Keys);
                        foreach (string timeOfDay in timeKeys)
                        {
                            // This condition verifies whether or not this specific timeOfDay (E.G. 12:00) contains as many available rooms as there are
                            // sections in a class.
                            if (rooms[weekday][timeOfDay].Count() == classes[progCode].Count()
                                || (iteration == 1 && rooms[weekday][timeOfDay].Count() >= classes[progCode].Count()))
                            {
                                // Next, we need to verify that this room slot is large enough to accomodate our exam duration
                                bool valid = true;
                                DateTime start = DateTime.Now;
                                DateTime end = DateTime.Now;
                                try
                                {
                                    start = Convert.ToDateTime(timeOfDay);

                                    List<string> spefRoomKeys = new List<string>(rooms[weekday][timeOfDay].Keys);
                                    foreach (string room in spefRoomKeys)
                                    {
                                        if (rooms[weekday][timeOfDay][room] == null)
                                            continue;
                                        try
                                        {
                                            end = Convert.ToDateTime(rooms[weekday][timeOfDay][room]);
                                            TimeSpan roomAvailability = end.Subtract(start);
                                            List<string> sectionKeys = new List<string>(classes[progCode].Keys);
                                            foreach (string section in sectionKeys)
                                            {
                                                if (int.Parse(classes[progCode][section]["DUR"].ToString()) > roomAvailability.Hours)
                                                    valid = false;
                                            }
                                        }
                                        catch (Exception e) { valid = false; } // TO DO if this catches, we should return a row value through an error message
                                    }
                                }
                                catch (Exception e) { valid = false; } // TO DO if this catches, we should return a row value through an error message

                                // If this room slot is valid, it means all of our sections can be accomodated here! Yipee!!
                                if (valid)
                                {
                                    List<string> sectionKeys = new List<string>(classes[progCode].Keys);
                                    foreach (string section in sectionKeys)
                                    {
                                        List<string> spefRoomKeys = new List<string>(rooms[weekday][timeOfDay].Keys);
                                        foreach (string room in spefRoomKeys)
                                        {
                                            Exam sectionExam = new Exam();
                                            sectionExam.Code = progCode;
                                            sectionExam.Name = classes[progCode][section]["NAME"];
                                            sectionExam.Section = section;
                                            sectionExam.Faculty = classes[progCode][section]["FACULTY"];
                                            sectionExam.Proctor = null;
                                            sectionExam.Room = room;
                                            sectionExam.Day = weekday;
                                            sectionExam.Start = string.Format("{0:hh:mm tt}", start);
                                            sectionExam.End = string.Format("{0:hh:mm tt}", end);
                                            sectionExam.Duration = classes[progCode][section]["DUR"] + " hour(s)";
                                            exams.Add(sectionExam);

                                            rooms[weekday][timeOfDay].Remove(room);
                                            break;
                                        }
                                        classes[progCode].Remove(section);
                                    }
                                }
                            }
                        }
                    }
                }
                iteration++; // iteration counter limits potential looping locking caused by inadequate room availibility. If we reach that, we should output an error
                classKeys = new List<string>(classes.Keys);
                roomKeys = new List<string>(rooms.Keys);
            }
            // With all exams scheduled, we can now assign proctors!
            List<string> weekdays = new List<string>(facultySchedule.Keys);
            for (int i = 0; i < exams.Count(); i++)
            {
                DateTime start = Convert.ToDateTime(exams[i].Start);
                DateTime timeAvailable = Convert.ToDateTime(exams[i].End).AddMinutes(30);
                List<string> facultyAvailable = new List<string>(facultySchedule[exams[i].Day].Keys);
                foreach (string proctor in facultyAvailable)
                {
                    if (exams[i].Proctor != null)
                        break;
                    if (facultySchedule[exams[i].Day][exams[i].Faculty] == new DateTime() || facultySchedule[exams[i].Day][exams[i].Faculty] < start)
                    {
                        facultySchedule[exams[i].Day][exams[i].Faculty] = timeAvailable;
                        exams[i].Proctor = exams[i].Faculty;
                        break;
                    }
                    else if (facultySchedule[exams[i].Day][proctor] == new DateTime() || facultySchedule[exams[i].Day][proctor] < start)
                    {
                        facultySchedule[exams[i].Day][proctor] = timeAvailable;
                        exams[i].Proctor = proctor;
                        break;
                    }
                }
            }
            // Check for unassigned classes & send them to an unassigned list!
            classKeys = new List<string>(classes.Keys);
            List<Exam> unassigned = new List<Exam>();
            foreach (string program in classKeys)
            {
                if (classes[program].Count == 0)
                    continue;
                List<string> sections = new List<string>(classes[program].Keys);
                foreach (string section in sections)
                {
                    Exam un = new Exam();
                    un.Code = program;
                    un.Name = classes[program][section]["NAME"];
                    un.Section = section;
                    un.Faculty = classes[program][section]["FACULTY"];
                    unassigned.Add(un);
                }
            }
            // We don't need to list the rooms still available if we don't have unassigned classes.
            List<Exam> availableRooms = new List<Exam>();
            if (unassigned.Count != 0)
            {
                roomKeys = new List<string>(rooms.Keys);
                foreach (string weekday in roomKeys)
                {
                    List<string> times = new List<string>(rooms[weekday].Keys);
                    foreach (string time in times)
                    {
                        if (rooms[weekday][time].Count == 0)
                            continue;
                        List<string> roomNums = new List<string>(rooms[weekday][time].Keys);
                        foreach (string room in roomNums)
                        {
                            Exam availableRoom = new Exam();
                            availableRoom.Room = room;
                            availableRoom.Day = weekday;
                            availableRoom.Start = string.Format("{0:hh:mm tt}", Convert.ToDateTime(time));
                            availableRoom.End = string.Format("{0:hh:mm tt}", Convert.ToDateTime(rooms[weekday][time][room]));
                            availableRooms.Add(availableRoom);
                        }
                    }
                }
            }

            Object[] output = {exams, unassigned, availableRooms};
            if (unassigned.Count != 0)
                return output;
            return output;
        }
        public void SaveExamSheet(Object[] examGeneration, string fileName)
        {
            List<Exam> exams = (List<Exam>)examGeneration[0];
            List<Exam> unassigned = (List<Exam>)examGeneration[1];
            List<Exam> availableRooms = (List<Exam>)examGeneration[2];

            XLWorkbook doc = new XLWorkbook();
            IXLWorksheet sheet = doc.Worksheets.Add("Exam Schedule");

            sheet.Column("A").Width = 15;
            sheet.Column("B").Width = 60;
            sheet.Column("C").Width = 10;
            sheet.Columns("D:I").Width = 20;

            int row = 1;
            if(unassigned.Count != 0)
            {
                sheet.Cells("A" + row + ":I" + row).Style.Fill.SetBackgroundColor(XLColor.Crimson);
                sheet.Row(row).Style.Font.SetFontColor(XLColor.White);

                sheet.Cell("A" + row).Value = "There weren't enough rooms available to fit these exams.";

                row++;
                sheet.Cells("A" + row + ":I" + row).Style.Fill.SetBackgroundColor(XLColor.Crimson);
                sheet.Row(row).Style.Font.SetFontColor(XLColor.White);
                sheet.Row(row).Style.Font.SetBold();

                sheet.Cell("A" + row).Value = "Code";
                sheet.Cell("B" + row).Value = "Course Name";
                sheet.Cell("C" + row).Value = "Section";
                sheet.Cell("D" + row).Value = "Faculty";
                foreach (Exam datarow in unassigned)
                {
                    row++;
                    sheet.Cell("A" + row).Value = datarow.Code;
                    sheet.Cell("B" + row).Value = datarow.Name;
                    sheet.Cell("C" + row).Value = datarow.Section;
                    sheet.Cell("D" + row).Value = datarow.Faculty;
                }
                row++;
                sheet.Cells("A" + row + ":I" + row).Style.Fill.SetBackgroundColor(XLColor.Crimson);
                sheet.Row(row).Style.Font.SetFontColor(XLColor.White);

                sheet.Cell("A" + row).Value = "Here's a list of available rooms.";

                row++;
                sheet.Cells("A" + row + ":I" + row).Style.Fill.SetBackgroundColor(XLColor.Crimson);
                sheet.Row(row).Style.Font.SetFontColor(XLColor.White);
                sheet.Row(row).Style.Font.SetBold();

                sheet.Cell("D" + row).Value = "Faculty";
                sheet.Cell("F" + row).Value = "Room";
                sheet.Cell("G" + row).Value = "Day";
                sheet.Cell("H" + row).Value = "Time";
                foreach (Exam datarow in availableRooms)
                {
                    row++;
                    sheet.Cell("D" + row).Value = datarow.Faculty;
                    sheet.Cell("F" + row).Value = datarow.Room;
                    sheet.Cell("G" + row).Value = datarow.Day;
                    sheet.Cell("H" + row).Value = datarow.Start + " - " + datarow.End;
                }
                row++;
            }
            
            sheet.Cells("A" + row + ":I" + row).Style.Fill.SetBackgroundColor(XLColor.Charcoal);
            sheet.Row(row).Style.Font.SetFontColor(XLColor.White);
            sheet.Row(row).Style.Font.SetBold();

            sheet.Cell("A" + row).Value = "Code";
            sheet.Cell("B" + row).Value = "Course Name";
            sheet.Cell("C" + row).Value = "Section";
            sheet.Cell("D" + row).Value = "Faculty";
            sheet.Cell("E" + row).Value = "Proctor";
            sheet.Cell("F" + row).Value = "Room";
            sheet.Cell("G" + row).Value = "Day";
            sheet.Cell("H" + row).Value = "Time";
            sheet.Cell("I" + row).Value = "Duration";
            foreach (Exam datarow in exams)
            {
                row++;
                sheet.Cell("A" + row).Value = datarow.Code;
                sheet.Cell("B" + row).Value = datarow.Name;
                sheet.Cell("C" + row).Value = datarow.Section;
                sheet.Cell("D" + row).Value = datarow.Faculty;
                sheet.Cell("E" + row).Value = datarow.Proctor;
                sheet.Cell("F" + row).Value = datarow.Room;
                sheet.Cell("G" + row).Value = datarow.Day;
                sheet.Cell("H" + row).Value = datarow.Start + " - " + datarow.End;
                sheet.Cell("I" + row).Value = datarow.Duration;
            }

            doc.SaveAs(AppDomain.CurrentDomain.GetData("DataDirectory").ToString() + "/sheetStorage/Exam_" + fileName);
        }
        public string GenerateCalendarEvents(List<Exam> exams)
        {

            return null;
        }

        public bool DownloadFileFromPath(string path)
        {
            FileInfo downloadFile = new FileInfo(path);
            if (downloadFile.Exists)
            {
                Response.Clear();
                Response.ClearHeaders();
                Response.ClearContent();
                Response.AddHeader("content-disposition", "attachment; filename=" + downloadFile.Name);
                Response.AddHeader("Content-Type", "application/Excel");
                Response.ContentType = "application/vnd.xls";
                Response.AddHeader("Content-Length", downloadFile.Length.ToString());
                Response.WriteFile(downloadFile.FullName);
                Response.End();
                return true;
            }
            else
            {
                return false;
            }
        }
    }
}