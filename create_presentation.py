#!/usr/bin/env python3
# -*- coding: utf-8 -*-

from pptx import Presentation
from pptx.util import Inches, Pt
from pptx.enum.text import PP_ALIGN, MSO_ANCHOR
from pptx.dml.color import RGBColor

def create_title_slide(prs, title, subtitle=""):
    """Create a title slide"""
    slide_layout = prs.slide_layouts[0]
    slide = prs.slides.add_slide(slide_layout)
    
    title_shape = slide.shapes.title
    subtitle_shape = slide.placeholders[1]
    
    title_shape.text = title
    if subtitle:
        subtitle_shape.text = subtitle
    
    return slide

def create_content_slide(prs, title, content_items):
    """Create a content slide with bullet points"""
    slide_layout = prs.slide_layouts[1]
    slide = prs.slides.add_slide(slide_layout)
    
    title_shape = slide.shapes.title
    title_shape.text = title
    
    body_shape = slide.placeholders[1]
    text_frame = body_shape.text_frame
    text_frame.clear()
    
    for item in content_items:
        p = text_frame.add_paragraph()
        p.text = item
        p.level = 0
        p.font.size = Pt(18)
    
    return slide

def create_section_slide(prs, title):
    """Create a section divider slide"""
    slide_layout = prs.slide_layouts[2]
    slide = prs.slides.add_slide(slide_layout)
    
    title_shape = slide.shapes.title
    title_shape.text = title
    
    return slide

def main():
    # Create presentation
    prs = Presentation()
    prs.slide_width = Inches(10)
    prs.slide_height = Inches(7.5)
    
    # Slide 1: Title
    create_title_slide(prs, 
                      "بحث حول شركة مجمع سيم الجزائرية",
                      "Groupe SIM")
    
    # Slide 2: Introduction
    create_content_slide(prs, "المقدمة", [
        "مجمع سيم من أبرز الشركات الجزائرية الرائدة في مجال الصناعات الغذائية",
        "احتل مكانة مرموقة في السوق الجزائرية والإفريقية",
        "منتجات متنوعة وجودة عالية"
    ])
    
    # Slide 3: Section - History
    create_section_slide(prs, "1. تاريخ ونشأة مجمع سيم")
    
    # Slide 4: Beginnings
    create_content_slide(prs, "1.1 البدايات", [
        "تأسست الشركة: عام 1998",
        "المؤسس: رجل الأعمال الجزائري سليم عثماني",
        "الموقع: مقرها الرئيسي في الجزائر العاصمة",
        "البداية: شركة صغيرة متخصصة في إنتاج المواد الغذائية الأساسية"
    ])
    
    # Slide 5: Growth
    create_content_slide(prs, "1.2 التطور والنمو", [
        "التوسع السريع خلال العقدين الماضيين",
        "التنويع: من منتج واحد إلى مجموعة واسعة",
        "الاستثمارات في أحدث التقنيات والمعدات",
        "توفير آلاف فرص العمل للجزائريين"
    ])
    
    # Slide 6: Current Position
    create_content_slide(prs, "1.3 المكانة الحالية", [
        "الريادة الوطنية: من أكبر المجموعات الصناعية الخاصة",
        "التوسع الإفريقي: دخول أسواق عدة دول إفريقية",
        "الاعتراف الدولي: شهادات جودة عالمية (ISO)"
    ])
    
    # Slide 7: Section - Products
    create_section_slide(prs, "2. المنتجات الرئيسية")
    
    # Slide 8: Dairy Products
    create_content_slide(prs, "2.1 منتجات الألبان", [
        "الحليب المبستر والمعقم",
        "الأجبان بأنواعها (جبن طري، مطبوخ، أبيض)",
        "الزبادي والياغورت",
        "الزبدة والقشدة",
        "العلامات التجارية: Candia و Soummam"
    ])
    
    # Slide 9: Other Products
    create_content_slide(prs, "2.2 منتجات أخرى", [
        "عصائر الفواكه الطبيعية",
        "المشروبات الغازية",
        "المياه المعدنية",
        "المعجنات والحلويات",
        "المواد الغذائية المصنعة"
    ])
    
    # Slide 10: Section - Marketing
    create_section_slide(prs, "3. استراتيجيات التسويق")
    
    # Slide 11: Traditional Marketing
    create_content_slide(prs, "3.1 التسويق التقليدي", [
        "شبكة توزيع ضخمة تغطي كامل التراب الوطني",
        "نقاط البيع في جميع المحلات والسوبرماركت",
        "الشاحنات المبردة لضمان الجودة",
        "أسعار معقولة تناسب مختلف الشرائح",
        "العروض الترويجية والخصومات الموسمية"
    ])
    
    # Slide 12: Digital Marketing
    create_content_slide(prs, "3.2 التسويق الرقمي", [
        "صفحات نشطة على Facebook, Instagram, YouTube",
        "محتوى تفاعلي: وصفات طبخ، نصائح صحية، مسابقات",
        "الإعلانات التلفزيونية والإذاعية",
        "رعاية الفعاليات الرياضية",
        "الشراكات الاجتماعية والمبادرات الخيرية"
    ])
    
    # Slide 13: Quality Focus
    create_content_slide(prs, "3.3 التركيز على الجودة والصحة", [
        "شعارات تسويقية: الجودة، الطبيعة، والصحة",
        "إبراز الشهادات الدولية والمحلية",
        "الشفافية في نشر معلومات المواد الأولية"
    ])
    
    # Slide 14: Section - Campaigns
    create_section_slide(prs, "4. أشهر الإشهارات والحملات")
    
    # Slide 15: Campaign 1
    create_content_slide(prs, "4.1 حملة 'Soummam - طبيعي 100%'", [
        "الفكرة: منتجات من حليب طبيعي 100%",
        "استخدام صور الطبيعة الجزائرية والمزارع المحلية",
        "الرسالة: 'من الطبيعة الجزائرية إلى مائدتك'",
        "التأثير: تعزيز ثقة المستهلك والهوية الوطنية"
    ])
    
    # Slide 16: Campaign 2
    create_content_slide(prs, "4.2 حملة رمضان السنوية", [
        "إعلانات عاطفية تركز على قيم الأسرة والتضامن",
        "استخدام ممثلين ومشاهير جزائريين",
        "موسيقى تصويرية وأغاني رمضانية",
        "الشعار: 'Soummam - رفيق موائدكم في رمضان'"
    ])
    
    # Slide 17: Campaign 3
    create_content_slide(prs, "4.3 حملة 'Candia - للكبار والصغار'", [
        "استهداف جميع أفراد الأسرة",
        "التركيز على الفوائد الصحية للحليب",
        "مشاهد عائلية: أطفال يشربون الحليب",
        "رسالة صحية: 'الحليب لعظام قوية وجسم سليم'"
    ])
    
    # Slide 18: Campaign 4
    create_content_slide(prs, "4.4 حملة 'Soummam Danino'", [
        "الاستهداف: الأطفال والأمهات",
        "شخصيات كرتونية محببة للأطفال",
        "الرسالة: 'غذاء صحي ولذيذ لأطفالك'",
        "مسابقات للأطفال على العبوات",
        "الشعار: 'Danino - الطعم اللي يحبوه، الصحة اللي تحبيها'"
    ])
    
    # Slide 19: Social Responsibility
    create_content_slide(prs, "4.5 حملة المسؤولية الاجتماعية", [
        "دعم التعليم: توزيع منتجات مجانية في المدارس",
        "الصحة: حملات توعية حول التغذية السليمة",
        "البيئة: مبادرات إعادة التدوير",
        "تعزيز الصورة الإيجابية للعلامة التجارية"
    ])
    
    # Slide 20: Section - Success Factors
    create_section_slide(prs, "5. عوامل النجاح")
    
    # Slide 21: Success Factors
    create_content_slide(prs, "عوامل النجاح", [
        "الجودة: معايير صارمة ومختبرات متطورة",
        "الابتكار: منتجات جديدة وتغليف عملي",
        "الاستثمار في الموارد البشرية والتدريب",
        "التسويق الذكي وفهم السوق الجزائري"
    ])
    
    # Slide 22: Section - Challenges
    create_section_slide(prs, "6. التحديات والمنافسة")
    
    # Slide 23: Challenges
    create_content_slide(prs, "التحديات والاستجابة", [
        "التحديات: المنافسة الأجنبية، تقلبات الأسعار، اللوجستيك",
        "الاستجابة: التطوير المستمر للمنتجات",
        "التوسع: فتح أسواق جديدة في إفريقيا",
        "الشراكات: تعاون مع علامات عالمية"
    ])
    
    # Slide 24: Section - Future Vision
    create_section_slide(prs, "7. الرؤية المستقبلية")
    
    # Slide 25: Future Goals
    create_content_slide(prs, "الأهداف والمشاريع القادمة", [
        "التصدير: زيادة الصادرات إلى الدول الإفريقية والعربية",
        "التنويع: دخول قطاعات جديدة",
        "الاستدامة: ممارسات صديقة للبيئة",
        "مصانع جديدة: توسيع الطاقة الإنتاجية",
        "البحث والتطوير والرقمنة"
    ])
    
    # Slide 26: Conclusion
    create_content_slide(prs, "الخاتمة", [
        "مجمع سيم نموذج ناجح للمقاولاتية الجزائرية",
        "بناء إمبراطورية صناعية بفضل الجودة والابتكار",
        "نجاح في فرض المنتجات محليًا والتوسع خارجيًا",
        "حملات إعلانية مؤثرة تلامس الوجدان الجزائري",
        "جزء لا يتجزأ من الحياة اليومية للأسر الجزائرية"
    ])
    
    # Slide 27: References
    create_content_slide(prs, "المراجع والمصادر", [
        "الموقع الرسمي لمجمع سيم",
        "تقارير إعلامية جزائرية",
        "دراسات السوق الجزائرية",
        "مقابلات مع مسؤولي الشركة",
        "تحليلات الحملات الإعلانية"
    ])
    
    # Slide 28: Thank You
    create_title_slide(prs, "شكراً لحسن الاستماع", "تاريخ إعداد البحث: نوفمبر 2025")
    
    # Save presentation
    prs.save('بحث_مجمع_سيم_الجزائرية.pptx')
    print("✓ تم إنشاء العرض التقديمي بنجاح!")
    print("✓ اسم الملف: بحث_مجمع_سيم_الجزائرية.pptx")

if __name__ == "__main__":
    main()
