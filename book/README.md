# كتاب مقرر إدارة اتخاذ القرار على أساس البيانات (MIS6004)

هذا المجلد يحوي مولّد الفصول 10–13 ومخرجاتها، بالتصميم نفسه المستعمل في الفصول 1–9 المحفوظة في Google Drive (مجلد «مشروع الكتاب»).

## المخرجات (`out/`)

| الفصل | الكتاب (Word) | الشرائح (PowerPoint) | التمرين (Excel) |
|---|---|---|---|
| 10 التحيّزات والضجيج في القرار | `out/chapter/Ch10_Book.docx` | `out/slides/Ch10_Slides.pptx` | `out/excel/Ch10_BiasNoise.xlsx` |
| 11 سرد القصة بالبيانات | `out/chapter/Ch11_Book.docx` | `out/slides/Ch11_Slides.pptx` | `out/excel/Ch11_DataStorytelling.xlsx` |
| 12 التنفيذ والتجريب والحوكمة | `out/chapter/Ch12_Book.docx` | `out/slides/Ch12_Slides.pptx` | `out/excel/Ch12_ExperimentGovernance.xlsx` |
| 13 قرار متكامل | `out/chapter/Ch13_Book.docx` | `out/slides/Ch13_Slides.pptx` | `out/excel/Ch13_IntegratedDecision.xlsx` |

كل فصل يتبع بنية الفصول السابقة: سؤال الفصل، حالة افتتاحية، مخرجات التعلم، خريطة بأربعة محاور، أقسام مرقمة (فقرات + جدول أو بطاقتان أو خطوات + خلاصة + مثال محلول)، التطبيق في Excel، حالة للنقاش ومهمة بحثية، الخلاصة، أسئلة المراجعة، المصطلحات، المراجع، ملحق الإجابات النموذجية.

ملفات Excel: «التعليمات» و«شرح_الحالة» وتمرينان محسوبان بالصيغ وورقة بطاقة بتقييم ذاتي (0/1/2 لخمسة معايير). الخلايا الصفراء مدخلات والزرقاء صيغ. جميع الصيغ تُحسب دون أخطاء (تحقق آلي عبر LibreOffice).

## إعادة البناء

```bash
cd book/gen
python3 illustrations.py ../img      # رسوم الأغلفة والحالات (Node Playwright)
python3 build.py 10 11 12 13         # يكتب إلى ../out
python3 verify_xlsx.py ../out/excel/Ch10_BiasNoise.xlsx   # فحص الصيغ
```

- `gen/content/chNN.py`: نص الفصل كاملا (يخدم الكتاب والشرائح معا) ودالة بناء ملف Excel.
- `gen/docgen.py`: الكتاب، ويستعمل `templates/Ch09_Book.docx` للأنماط والترقيم والتذييل فقط.
- `gen/slidegen.py`: الشرائح، ويستعمل `templates/Ch09_Slides.pptx` للخلفيات والقالب فقط.
- `gen/xlsxkit.py`: أنماط ملفات Excel المشتركة.

يحتاج البناء إلى خط Cairo لمعاينة دقيقة فقط (الملفات تحدد الخط ولا تضمّنه).

انظر `STATUS.md` لما تبقى من أعمال المشروع.
