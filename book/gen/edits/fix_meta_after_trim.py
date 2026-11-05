"""Update outcomes/axes/summary/cross-references after trim_and_merge.py."""
import os, re
GEN = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
R = {
 5: [("النموذج المرجعي – السلاسل الزمنية", "النموذج المرجعي – اختيار النموذج – قاعدة القرار")],
 6: [("الكلفة – المنحنيات – الخطأ العددي – المعايرة – الرفع", "الكلفة – مقارنة النماذج – الخطأ العددي – الثقة في الاحتمالات – الرفع"),
     ("وفي القرارات التي تحسب القيمة المتوقعة تُفحص معايرة الاحتمالات ودرجة بريير.", "وفي القرارات التي تحسب القيمة المتوقعة تُفحص معايرة الاحتمالات بجدول الفئات."),
     ("معايرة الاحتمالات ودرجة بريير قبل ضربها في العائد.", "معايرة الاحتمالات بجدول الفئات قبل ضربها في العائد."),
     ("الدقة تضلل في الأحداث النادرة، ويلزم الضبط والاسترجاع والصافي والمعايرة.", "الدقة تضلل في الأحداث النادرة، ويلزم الضبط والاسترجاع والصافي والثقة في الاحتمالات.")],
 7: [("بدرجة بريير التي وردت في الفصل السادس", "بدرجة بريير التي يعرضها ملحق الموضوعات للاستزادة")],
 8: [("نموذجا خطيا بنقاط الزوايا ويفسّر أسعاره الظلية، ويميّز مواضع البرمجة الصحيحة.", "نموذجا خطيا في جدول بيانات بأداة Solver، ويفسّر أسعاره الظلية، ويحوّل الحل الأمثل إلى خطة قابلة للتنفيذ."),
     ("البرمجة الخطية – الحل الهندسي – الأسعار الظلية – البرمجة الصحيحة – اختيار النموذج", "البرمجة الخطية – بناء النموذج في جدول بيانات – الأسعار الظلية – من الحل إلى الخطة – اختيار النموذج"),
     ("خطة الإنتاج – نقاط الزوايا – الأسعار الظلية", "خطة الإنتاج – الحل بأداة Solver – الأسعار الظلية"),
     ("البرمجة الصحيحة تتناول القرارات التي لا تقبل القسمة، ولا تضمن القواعد الجشعة الأمثل.", "الحل الأمثل يُقرأ من تقرير الحساسية، ويُحوَّل إلى خطة تراعي القيود العملية قبل التنفيذ.")],
}
import importlib, pprint, sys
sys.path.insert(0, os.path.join(GEN, "content"))


def walk(o, a, b, hits):
    if isinstance(o, str):
        if a in o:
            hits.append(1)
        return o.replace(a, b)
    if isinstance(o, dict):
        return {k: walk(v, a, b, hits) for k, v in o.items()}
    if isinstance(o, list):
        return [walk(v, a, b, hits) for v in o]
    if isinstance(o, tuple):
        return tuple(walk(v, a, b, hits) for v in o)
    return o


for n, pairs in R.items():
    ch = importlib.import_module(f"ch{n:02d}").CH
    for a, b in pairs:
        hits = []
        ch = walk(ch, a, b, hits)
        if not hits:
            print("  (already applied or missing)", n, a[:40])
    p = os.path.join(GEN, "content", f"ch{n:02d}.py")
    head = open(p, encoding="utf-8").read().partition("CH = ")[0]
    open(p, "w", encoding="utf-8").write(head + "CH = " + pprint.pformat(ch, width=150, sort_dicts=False) + "\n\nbuild_xlsx = None\n")
    print("ok", n)
