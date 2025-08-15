### ورود فایل زمان‌بندی به MS Project

- **فایل**: `msproject_schedule.csv`
- **محتوا**: WBS، مدت‌زمان‌ها (روز)، وابستگی‌ها (FS) و توضیحات

#### مراحل ورود (Import)
1) در MS Project یک پروژه خالی بسازید.
2) از مسیر File > Open > Browse فایل `msproject_schedule.csv` را انتخاب کنید (نوع فایل Text/CSV).
3) در Import Wizard گزینه New map را بزنید.
4) تیک "First row contains headers" را بزنید.
5) Map را این‌طور تنظیم کنید (Data Type: Tasks):
   - `Task Name` → Name
   - `Duration` → Duration
   - `Predecessors` → Predecessors
   - `Outline Level` → Outline Level
   - `Notes` → Notes
6) Next → As a new project → Finish.

#### تنظیمات پروژه پس از ورود
- از Project > Project Information تاریخ شروع پروژه را تعیین کنید.
- در Project > Change Working Time در صورت نیاز تقویم کاری را به **شنبه تا پنج‌شنبه** (یا مدنظر شما) تغییر دهید.
- با دابل‌کلیک روی Summary Tasks (سطح‌های 1 و 2) ساختار WBS را بررسی کنید.

#### نکات
- مدت‌زمان‌ها قابل ویرایش هستند. با تغییر مدت‌زمان، برنامه به‌صورت خودکار باززمان‌بندی می‌شود.
- وابستگی‌ها بر اساس FS تعریف شده‌اند؛ می‌توانید Lag/Lead اضافه کنید (مثلاً 10d+، 2d-).
- اگر مقادیر اجرایی و راندمان اکیپ‌ها را دارید، می‌توانم مدت‌زمان‌ها را دقیق بر اساس آن محاسبه و به‌روزرسانی کنم.