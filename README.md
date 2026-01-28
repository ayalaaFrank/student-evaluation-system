# מערכת הערכת סטודנטים (Student Evaluation System)

תוכנית C# קונסולה שמעבדת נתוני סטודנטים מקובץ CSV, מחשבת ציון סופי, מסננת מי שעברו (ציון ≥70), ומפיקה מכתבי הסמכה אישיים בפורמט PDF על בסיס תבנית Word אחת.

## לוגיקה עיקרית

1. **טעינת נתונים**  
   קוראת קובץ CSV (Data/Data.csv) עם פורמט:  
   FirstName,LastName,Department,TheoryScore,PracticalScore,Phone,Email  
   מדלגת על כותרת, מטפלת בשדות ריקים וממירה ציונים למספרים.

2. **ניקוי הנתונים**  
   - תיקון שמות: Capitalize ראשון של כל מילה (FixName).  
   - הסרת כפילויות: GroupBy על שם פרטי + משפחה + מחלקה (לוקחת את הרשומה הראשונה).

3. **חישוב ציון סופי**  
   FinalScore = (60% PracticalScore) + (40% TheoryScore), מעוגל למספר שלם.

4. **סינון והודעה**  
   - רק סטודנטים עם FinalScore ≥ 70 מקבלים הודעה ומכתב PDF.  
   - אם FinalScore ≥ 90: הודעה חיובית + "נמצאת מתאימ/ה לתפקיד מוביל/ה טכנולוגי מחלקתית".  
   - אם 70 ≤ FinalScore < 90: הודעה שלילית ("לא עברת... לא נמצא תפקיד מתאים").

5. **שמירת נתונים מעובדים**  
   יוצרת ProcessedStudents.csv (רק מי שעברו) עם כל השדות + Message (בטיפול בפסיקים ומרכאות).

6. **הפקת מכתבי PDF**  
   - משתמשת בתבנית Word אחת (Templates/Letters.docx) עם Mail Merge Fields.  
   - ממלאת את השדות (FirstName, LastName, Department, Phone, Email, FinalScore, Message).  
   - שומרת PDF אישי לכל סטודנט (שם_משפחה.pdf).  
   - ספרייה: GemBox.Document (חינמי עם FREE-LIMITED-KEY).

זה הכל – הקוד המלא נמצא כאן ב-GitHub.
