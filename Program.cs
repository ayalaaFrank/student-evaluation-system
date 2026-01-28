using System;
using System.IO;
using System.Collections.Generic;
using System.Linq;
using GemBox.Document;
using StudentEvaluationSystem;  // ← התקנת מ-NuGet: GemBox.Document

// --- קריאה של הקובץ ---
string filePath = "Data/Data.csv";
if (!File.Exists(filePath))
{
    Console.WriteLine("קובץ לא נמצא!");
    return;
}

string[] lines = File.ReadAllLines(filePath);
Console.WriteLine($"סך הכל שורות (כולל כותרת): {lines.Length}");

// --- יצירת רשימת סטודנטים ---
List<Student> students = new List<Student>();
for (int i = 1; i < lines.Length; i++)
{
    string[] parts = lines[i].Split(',');
    if (parts.Length < 7) continue;

    Student student = new Student
    {
        FirstName = string.IsNullOrWhiteSpace(parts[0]) ? "Unknown" : parts[0].Trim(),
        LastName = string.IsNullOrWhiteSpace(parts[1]) ? "Unknown" : parts[1].Trim(),
        Department = string.IsNullOrWhiteSpace(parts[2]) ? "Unknown" : parts[2].Trim(),
        TheoryScore = int.TryParse(parts[3], out int theory) ? theory : 0,
        PracticalScore = int.TryParse(parts[4], out int practical) ? practical : 0,
        Phone = parts.Length > 5 && !string.IsNullOrWhiteSpace(parts[5]) ? parts[5].Trim() : "Unknown",
        Email = parts.Length > 6 && !string.IsNullOrWhiteSpace(parts[6]) ? parts[6].Trim() : "Unknown"
    };
    students.Add(student);
}

// --- תיקון שמות (Capitalize) ---
foreach (var s in students)
{
    s.FirstName = FixName(s.FirstName);
    s.LastName = FixName(s.LastName);
}

// --- הסרת כפילויות (על שם + משפחה + מחלקה) ---
students = students
    .GroupBy(s => $"{s.FirstName}-{s.LastName}-{s.Department}")
    .Select(g => g.First())
    .ToList();

// --- חישוב ציון סופי + הודעה (רק למי שעבר ≥70) ---
foreach (var s in students)
{
    s.FinalScore = (int)Math.Round(s.PracticalScore * 0.6 + s.TheoryScore * 0.4, 0);

    if (s.FinalScore >= 70)
    {
        s.Message = GetMessage(s);
    }
    else
    {
        s.Message = null;
    }
}

// --- שמירת CSV מעובד (אופציונלי, אבל שמרת אותו) ---
string outputDir = Path.Combine(AppContext.BaseDirectory, "..\\..\\..\\Output");
Directory.CreateDirectory(outputDir);

string outputPath = Path.Combine(outputDir, "ProcessedStudents.csv");
using (StreamWriter writer = new StreamWriter(outputPath, false))
{
    writer.WriteLine("FirstName,LastName,Department,TheoryScore,PracticalScore,Phone,Email,FinalScore,Message");
    foreach (var s in students.Where(x => x.Message != null))
    {
        string safeMessage = (s.Message ?? "").Replace("\r", "").Replace("\n", " ");
        if (safeMessage.Contains(",") || safeMessage.Contains("\""))
        {
            safeMessage = $"\"{safeMessage.Replace("\"", "\"\"")}\"";
        }
        writer.WriteLine($"{s.FirstName},{s.LastName},{s.Department},{s.TheoryScore},{s.PracticalScore},{s.Phone},{s.Email},{s.FinalScore},{safeMessage}");
    }
}
Console.WriteLine($"\nנתונים מעובדים נשמרו: {outputPath}");

// --- הפקת PDF עם GemBox (חינמי + Mail Merge) ---
ComponentInfo.SetLicense("FREE-LIMITED-KEY");  // ← חובה!

string templatePath = Path.Combine(AppContext.BaseDirectory, "..\\..\\..\\Templates\\Letters.docx");

if (!File.Exists(templatePath))
{
    Console.WriteLine("תבנית Word לא נמצאה! שים Letters.docx בתיקיית Templates");
    return;
}

foreach (var s in students.Where(x => x.FinalScore >= 70))
{
    var document = DocumentModel.Load(templatePath);

    // Mail Merge – השדות בתבנית חייבים להיות MergeFields עם השמות האלה (בלי « »)
    document.MailMerge.Execute(
        new
        {
            FirstName = s.FirstName,
            LastName = s.LastName,
            Department = s.Department,
            Phone = s.Phone,
            Email = s.Email,
            FinalScore = s.FinalScore.ToString(),
            Message = s.Message?.Replace("\n", "\n\n") ?? ""
        });

    string pdfPath = Path.Combine(outputDir, $"{s.FirstName}_{s.LastName}.pdf");
    document.Save(pdfPath);

    Console.WriteLine($"נוצר PDF: {pdfPath}");
}

// --- פונקציות עזר ---
static string FixName(string name)
{
    if (string.IsNullOrWhiteSpace(name)) return "Unknown";
    name = name.Trim().ToLower();
    return char.ToUpper(name[0]) + name.Substring(1);
}

static string GetMessage(Student s)
{
    if (s.FinalScore >= 90)
    {
        return $"הרינו להודיעך כי עברת בהצלחה את ההכשרה. הציון הסופי שלך הינו {s.FinalScore}.\nנמצאת מתאימ/ה לתפקיד מוביל/ה טכנולוגי מחלקתית.";
    }
    else  // 70 ≤ ציון < 90
    {
        return "הרינו להודיעך כי לא עברת את ההכשרה אך לצערנו לא נמצא תפקיד מתאים עבורך.";
    }
}