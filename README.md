
# Field Timesheet PWA Ready (Single User)

نسخة فردية مناسبة للموبايل واللاب، بدون manager flow.

## التشغيل
```bash
py -m venv .venv
.venv\Scriptsctivate
pip install -r requirements.txt
python -m uvicorn app.main:app --reload
```

ثم افتح:
`http://127.0.0.1:8000`

## الميزات الحالية
- Setup Profile مرة واحدة
- Time Sheet + Evaluation في خطوتين
- منع تكرار نفس اليوم
- تجميع السجلات حسب الشهر تلقائيًا
- تعديل / حذف / تصدير Excel
- manifest + service worker كأساس PWA

## ملاحظة
هذه نسخة PWA-ready. التثبيت على الشاشة الرئيسية متاح من المتصفح، لكن تحسين الأوفلاين الكامل كنسخة بدون سيرفر ما يزال خطوة لاحقة.
