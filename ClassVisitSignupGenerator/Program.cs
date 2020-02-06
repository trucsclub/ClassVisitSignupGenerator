using Newtonsoft.Json;
using System;
using System.Drawing;
using System.IO;
using System.Linq;
using System.Text.RegularExpressions;

namespace ClassVisitSignupGenerator
{
    class Program
    {
        static void Main(string[] args)
        {
            //At this time, you need to go to "Browse classes", open the Network inspector, search for COMP, and open it in a new URL. Edit "pageMaxSize" to something like 150. Then save in the same folder as "classes.json"
            var file = File.ReadAllText(Path.Combine(Environment.CurrentDirectory, "classes.json"));

            var data = JsonConvert.DeserializeObject<BaseObject>(file);

            var pack = new OfficeOpenXml.ExcelPackage();
            pack.Workbook.Properties.Title = "Class Visit Signup Form";
            pack.Workbook.Properties.Created = DateTime.Now;

            var worksheet = pack.Workbook.Worksheets.Add("Signup");
            worksheet.CreateTitleRow(1, 0, "Course Number", "Course Title", "Faculty Member", "Room", "Monday", "Tuesday", "Wednesday", "Thursday", "Friday", "Notes");

            worksheet.Column(1).Width = 17.43d;
            worksheet.Column(2).Width = 60.0d;
            worksheet.Column(3).Width = 26.0d;
            worksheet.Column(4).Width = 10.0d;
            for (int i = 5; i <= 10; ++i)
            {
                worksheet.Column(i).Width = 12.0d;
                worksheet.Column(i).Style.Border.BorderAround(OfficeOpenXml.Style.ExcelBorderStyle.Thin, Color.FromArgb(150, 150, 150));
            }
            worksheet.Column(11).Width = 60.0d;
            
            worksheet.Row(1).Style.Border.Bottom.Style = OfficeOpenXml.Style.ExcelBorderStyle.Thick; //create separator between title row and the rest of the data
            worksheet.Column(4).Style.Border.Right.Style = OfficeOpenXml.Style.ExcelBorderStyle.Thin; //create a separator between the day cells and the descriptions

            int row = 2; //keeps track of current row to write data to. Start at 2 because 1 is used by the titles and API is 1-based

            //This massive ass LINQ statement will group everything by course number + section number, that way 
            //Also, we want to filter out COMP 4910 (and any additional non-campus/non-classroom courses!) so we check for anything without a BeginTime.
            foreach (var course in data.Data.GroupBy(x => x.CourseNumber + Regex.Replace(x.SequenceNumber, "[^0-9]+", string.Empty)).Select(x => new { MainClass = x.First(y => char.IsDigit(y.SequenceNumber[0])), MeetingsFaculty = x.SelectMany(y => y.MeetingsFaculty).Where(z => !string.IsNullOrEmpty(z.MeetingTime.BeginTime)).OrderBy(z => TimeSpan.Parse(z.MeetingTime.BeginTime)).ToList() }))
            {                
                if (course.MeetingsFaculty.Count == 0) continue;

                var firstMeeting = course.MeetingsFaculty.First();
                //note: the following line is a terrible, lazy hack that would not be performant in a real-world use case.
                worksheet.CreateRow(row, 0, Color.FromArgb(217, 217, 217), new string[] { $"{course.MainClass.Subject} {course.MainClass.CourseNumber} {course.MainClass.SequenceNumber}", course.MainClass.CourseTitle, course.MainClass.Faculty.FirstOrDefault()?.DisplayName ?? "UNKNOWN", firstMeeting.MeetingTime.Building + firstMeeting.MeetingTime.Room }.Concat(Enumerable.Range(1, 5).Select(x => firstMeeting.MeetingTime.MakeDayCell((DayOfWeek)x))).ToArray());
                row++;
                worksheet.CreateRow(row, 0, Color.FromArgb(242, 242, 242), "Availability");
                row++;
                foreach (var meetingTime in course.MeetingsFaculty.Skip(1))
                {
                    worksheet.CreateRow(row, 3, Color.FromArgb(217, 217, 217), new string[] { meetingTime.MeetingTime.Building.ToString() + meetingTime.MeetingTime.Room }.Concat(Enumerable.Range(1, 5).Select(x => meetingTime.MeetingTime.MakeDayCell((DayOfWeek)x))).ToArray());
                    row++;
                    worksheet.CreateRow(row, 0, Color.FromArgb(242, 242, 242), "Availability");
                    row++;
                }
                worksheet.Row(row).Style.Border.Top.Style = OfficeOpenXml.Style.ExcelBorderStyle.Thin;
            }

            worksheet.CreateRow(row, 0, "Automatically generated using a tool created by Tylor Pater.");
            pack.SaveAs(new FileInfo(Path.Combine(Environment.CurrentDirectory, $"Class Visit Signup - {data.Data.First().TermDesc}.xlsx")));

        }
    }

    public static class Extensions
    {
        public static OfficeOpenXml.ExcelRange MakeBold(this OfficeOpenXml.ExcelRange range)
        {
            range.Style.Font.Bold = true;
            return range;
        }

        public static string MakeDayCell(this MeetingTime time, DayOfWeek day) => day switch
            {
                DayOfWeek.Monday => time.Monday,
                DayOfWeek.Tuesday => time.Tuesday,
                DayOfWeek.Wednesday => time.Wednesday,
                DayOfWeek.Thursday => time.Thursday,
                DayOfWeek.Friday => time.Friday
            } ? $"{time.BeginTime.Insert(2, ":")} - {time.EndTime.Insert(2, ":")}" : ""; //have to insert colon manually
        
        public static OfficeOpenXml.ExcelWorksheet CreateTitleRow(this OfficeOpenXml.ExcelWorksheet worksheet, int row, int offset = 0, params string[] titles)
        {
            for (int i = 0; i < titles.Length; i++)
                worksheet.Cells[row, offset + i + 1].MakeBold().Value = titles[i];
            return worksheet;
        }

        public static OfficeOpenXml.ExcelWorksheet CreateRow(this OfficeOpenXml.ExcelWorksheet worksheet, int row, int offset, params string[] data)
        {
            for (int i = 0; i < data.Length; i++)
                worksheet.Cells[row, offset + i + 1].Value = data[i];
            return worksheet;
        }
        public static OfficeOpenXml.ExcelWorksheet CreateRow(this OfficeOpenXml.ExcelWorksheet worksheet, int row, int offset, Color fillColor, params string[] data)
        {
            worksheet.Row(row).Style.Fill.PatternType = OfficeOpenXml.Style.ExcelFillStyle.Solid;
            worksheet.Row(row).Style.Fill.BackgroundColor.SetColor(fillColor);
            return CreateRow(worksheet, row, offset, data);
        }
    }
}
