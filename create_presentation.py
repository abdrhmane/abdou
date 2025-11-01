#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
إنشاء عرض تقديمي PowerPoint حول مجمع سيم (المجمع الصناعي للحديد والصلب)
"""

try:
    from pptx import Presentation
    from pptx.util import Inches, Pt
    from pptx.enum.text import PP_ALIGN
    from pptx.dml.color import RGBColor
except ImportError:
    print("تثبيت المكتبات المطلوبة...")
    import subprocess
    import sys
    subprocess.check_call([sys.executable, "-m", "pip", "install", "python-pptx"])
    from pptx import Presentation
    from pptx.util import Inches, Pt
    from pptx.enum.text import PP_ALIGN
    from pptx.dml.color import RGBColor

def create_title_slide(prs, title, subtitle=""):
    """إنشاء شريحة عنوان"""
    slide_layout = prs.slide_layouts[0]
    slide = prs.slides.add_slide(slide_layout)
    
    title_shape = slide.shapes.title
    subtitle_shape = slide.placeholders[1]
    
    title_shape.text = title
    subtitle_shape.text = subtitle
    
    # تنسيق العنوان
    title_frame = title_shape.text_frame
    title_frame.paragraphs[0].font.size = Pt(44)
    title_frame.paragraphs[0].font.bold = True
    title_frame.paragraphs[0].font.color.rgb = RGBColor(0, 51, 102)
    
    return slide

def create_content_slide(prs, title, content_points):
    """إنشاء شريحة محتوى"""
    slide_layout = prs.slide_layouts[1]
    slide = prs.slides.add_slide(slide_layout)
    
    title_shape = slide.shapes.title
    title_shape.text = title
    
    # تنسيق العنوان
    title_shape.text_frame.paragraphs[0].font.size = Pt(32)
    title_shape.text_frame.paragraphs[0].font.bold = True
    title_shape.text_frame.paragraphs[0].font.color.rgb = RGBColor(0, 51, 102)
    
    # إضافة المحتوى
    body_shape = slide.placeholders[1]
    text_frame = body_shape.text_frame
    text_frame.clear()
    
    for point in content_points:
        p = text_frame.add_paragraph()
        p.text = point
        p.level = 0
        p.font.size = Pt(18)
        p.space_before = Pt(12)
    
    return slide

def main():
    # إنشاء عرض تقديمي جديد
    prs = Presentation()
    prs.slide_width = Inches(10)
    prs.slide_height = Inches(7.5)
    
    # الشريحة 1: العنوان الرئيسي
    create_title_slide(
        prs,
        "مجمع سيم للحديد والصلب",
        "المجمع الصناعي الجزائري للحديد والصلب\nبحث شامل"
    )
    
    # الشريحة 2: مقدمة
    create_content_slide(
        prs,
        "مقدمة",
        [
            "مجمع سيم هو أحد أكبر المجمعات الصناعية في الجزائر",
            "يقع في مدينة عنابة شرق الجزائر",
            "تأسس في السبعينيات كجزء من التصنيع الثقيل",
            "يعتبر رمزاً للصناعة الوطنية الجزائرية",
            "يساهم في الاقتصاد الوطني وتوفير فرص العمل"
        ]
    )
    
    # الشريحة 3: التاريخ والتأسيس
    create_content_slide(
        prs,
        "التاريخ والتأسيس",
        [
            "تأسس المجمع في عام 1964 تحت اسم 'سونسيد' (SONSID)",
            "بدأ الإنتاج الفعلي في السبعينيات",
            "تم إنشاؤه بالتعاون مع خبرات دولية",
            "كان جزءاً من استراتيجية التصنيع الوطنية",
            "شهد عدة مراحل من التطوير والتحديث"
        ]
    )
    
    # الشريحة 4: الموقع الجغرافي
    create_content_slide(
        prs,
        "الموقع الجغرافي والأهمية",
        [
            "يقع في مدينة عنابة على الساحل الشرقي للجزائر",
            "قرب الميناء البحري لتسهيل الاستيراد والتصدير",
            "موقع استراتيجي قرب مناجم الحديد",
            "سهولة الوصول إلى الأسواق المحلية والدولية",
            "يغطي مساحة واسعة من المنطقة الصناعية"
        ]
    )
    
    # الشريحة 5: المنتجات والإنتاج
    create_content_slide(
        prs,
        "المنتجات الرئيسية",
        [
            "قضبان الحديد المسلح للبناء",
            "الصفائح الفولاذية",
            "الأنابيب الفولاذية",
            "الحديد الخام والصلب",
            "منتجات متنوعة للصناعات المختلفة",
            "طاقة إنتاجية تصل إلى مئات الآلاف من الأطنان سنوياً"
        ]
    )
    
    # الشريحة 6: العملية الإنتاجية
    create_content_slide(
        prs,
        "العملية الإنتاجية",
        [
            "استخراج ومعالجة خام الحديد",
            "عملية الصهر في الأفران العالية",
            "التحويل إلى صلب في المحولات",
            "الدرفلة والتشكيل",
            "المعالجة النهائية والتشطيب",
            "مراقبة الجودة والاختبارات"
        ]
    )
    
    # الشريحة 7: الأهمية الاقتصادية
    create_content_slide(
        prs,
        "الأهمية الاقتصادية",
        [
            "يوفر آلاف فرص العمل المباشرة وغير المباشرة",
            "يساهم في الناتج المحلي الإجمالي",
            "يقلل من الاستيراد ويعزز الاكتفاء الذاتي",
            "يدعم قطاع البناء والتشييد",
            "يصدر منتجاته إلى دول أفريقية ومتوسطية",
            "يساهم في التنمية الإقليمية لمدينة عنابة"
        ]
    )
    
    # الشريحة 8: التحديات
    create_content_slide(
        prs,
        "التحديات التي تواجه المجمع",
        [
            "المنافسة من الواردات الأجنبية",
            "الحاجة إلى تحديث التكنولوجيا والمعدات",
            "التكاليف التشغيلية المرتفعة",
            "تقلبات أسعار المواد الخام عالمياً",
            "الحاجة إلى تدريب وتأهيل العمالة",
            "التحديات البيئية ومعايير الاستدامة"
        ]
    )
    
    # الشريحة 9: خطط التطوير
    create_content_slide(
        prs,
        "خطط التطوير والتحديث",
        [
            "برامج تحديث المعدات والتكنولوجيا",
            "زيادة الطاقة الإنتاجية",
            "تحسين جودة المنتجات",
            "تطوير منتجات جديدة",
            "الاستثمار في التدريب والكفاءات",
            "تطبيق معايير بيئية حديثة",
            "توسيع الأسواق التصديرية"
        ]
    )
    
    # الشريحة 10: الأثر الاجتماعي
    create_content_slide(
        prs,
        "الأثر الاجتماعي والبيئي",
        [
            "توفير فرص عمل لآلاف الأسر",
            "تطوير البنية التحتية في المنطقة",
            "برامج التكوين والتدريب المهني",
            "المساهمة في الأنشطة الاجتماعية والثقافية",
            "جهود للحد من التلوث البيئي",
            "برامج إعادة التدوير والاستدامة"
        ]
    )
    
    # الشريحة 11: الشراكات والتعاون
    create_content_slide(
        prs,
        "الشراكات والتعاون",
        [
            "شراكات مع شركات عالمية للتكنولوجيا",
            "التعاون مع الجامعات ومراكز البحث",
            "عضوية في منظمات صناعية دولية",
            "تبادل الخبرات مع مجمعات مماثلة",
            "برامج تدريب بالتعاون مع خبراء دوليين"
        ]
    )
    
    # الشريحة 12: الإحصائيات الرئيسية
    create_content_slide(
        prs,
        "إحصائيات ومعلومات رئيسية",
        [
            "عدد العمال: آلاف الموظفين",
            "الطاقة الإنتاجية: مئات الآلاف من الأطنان سنوياً",
            "المساحة: عشرات الهكتارات",
            "عدد الأفران والوحدات الإنتاجية",
            "حجم الاستثمارات على مر السنين",
            "نسبة المساهمة في السوق المحلي"
        ]
    )
    
    # الشريحة 13: مستقبل المجمع
    create_content_slide(
        prs,
        "رؤية المستقبل",
        [
            "التحول نحو الصناعة الذكية والرقمنة",
            "زيادة الاعتماد على الطاقات المتجددة",
            "تطوير منتجات عالية القيمة المضافة",
            "التوسع في الأسواق الأفريقية",
            "تحقيق الاستدامة البيئية الكاملة",
            "أن يصبح مرجعاً إقليمياً في صناعة الحديد والصلب"
        ]
    )
    
    # الشريحة 14: الخلاصة
    create_content_slide(
        prs,
        "الخلاصة",
        [
            "مجمع سيم ركيزة أساسية للصناعة الجزائرية",
            "يلعب دوراً حيوياً في الاقتصاد الوطني",
            "يواجه تحديات تتطلب التحديث والتطوير",
            "لديه إمكانيات كبيرة للنمو والتوسع",
            "يمثل نموذجاً للصناعة الوطنية الطموحة",
            "المستقبل واعد مع الاستثمار المناسب"
        ]
    )
    
    # الشريحة 15: شكراً
    create_title_slide(
        prs,
        "شكراً لحسن الاستماع",
        "مجمع سيم - فخر الصناعة الجزائرية"
    )
    
    # حفظ العرض التقديمي
    filename = "بحث_مجمع_سيم.pptx"
    prs.save(filename)
    print(f"✓ تم إنشاء العرض التقديمي بنجاح: {filename}")
    print(f"✓ عدد الشرائح: {len(prs.slides)}")
    print(f"✓ يمكنك الآن فتح الملف باستخدام PowerPoint أو LibreOffice Impress")

if __name__ == "__main__":
    main()
