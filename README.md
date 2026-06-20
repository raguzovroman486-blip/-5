from pptx import Presentation
from pptx.util import Inches, Pt
from pptx.enum.text import PP_ALIGN
from pptx.enum.shapes import MSO_SHAPE

def add_title_slide(prs, title, subtitle=""):
    s = prs.slides.add_slide(prs.slide_layouts[0])
    title_shape = s.shapes.title
    subtitle_shape = s.placeholders[1] if len(s.placeholders) > 1 else None
    title_shape.text = title
    if subtitle_shape:
        subtitle_shape.text = subtitle
    return s

def add_content_slide(prs, heading, items, image_path=None, page_num=None):
    s = prs.slides.add_slide(prs.slide_layouts[1])
    shapes = s.shapes
    shapes.title.text = heading

    body = s.placeholders[1]
    tf = body.text_frame
    tf.clear()

    for it in items:
        p = tf.add_paragraph()
        p.text = it
        p.level = 0

    if image_path:
        left = Inches(5)
        top = Inches(2)
        height = Inches(3)
        shapes.add_picture(image_path, left, top, height=height)

    # Номер страницы
    if page_num is not None:
        textbox = shapes.add_textbox(Inches(9), Inches(7.5), Inches(1.5), Inches(0.4))
        tf2 = textbox.text_frame
        tf2.text = str(page_num)
        for p in tf2.paragraphs:
            p.alignment = PP_ALIGN.RIGHT

def main():
 prs = Presentation()

    # Слайд 1: Оглавление
    s1 = prs.slides.add_slide(prs.slide_layouts[5])
    s1.shapes.title.text = "Название презентации"
    body = s1.shapes.placeholders[1]
    tf = body.text_frame
    tf.text = "Оглавление"
    p = tf.add_paragraph()
    p.text = "1. Введение"
    p = tf.add_paragraph()
    p.text = "2. Раздел 1 — Тема"
    p = tf.add_paragraph()
    p.text = "3. Раздел 2 — Тема"
    p = tf.add_paragraph()
    p.text = "4. Раздел 3 — Практические рекомендации"
    # Титульная страница оглавления может быть адаптирована

    # Слайды по плану (примерная структура)
    add_content_slide(prs, "Введение",
                      ["Цель презентации", "Контекст задачи", "Чего достичь"] ,
                      image_path=None, page_num=2)

    add_content_slide(prs, "Раздел 1 — Тема",
                      ["Ключевые идеи 1", "Ключевая идея 2", "Ключевые выводы"],
                      image_path=None, page_num=3)

    add_content_slide(prs, "Подтема 1.1",
                      ["Пункт 1", "Пункт 2", "Пункт 3"],
                      image_path=None, page_num=4)

    add_content_slide(prs, "Раздел 2 — Тема",
                      ["Основные моменты", "Пример использования", "Итоги"],
                      image_path=None, page_num=6)

    add_content_slide(prs, "Пример/Кейс",
                      ["Описание кейса", "Выводы по кейсу"],
                      image_path=None, page_num=7)

    add_content_slide(prs, "Раздел 3 — Практические рекомендации",
                      ["Шаги внедрения", "Проверки и риски"],
                      image_path=None, page_num=8)

    add_content_slide(prs, "Резюме",
                      ["Ключевые выводы", "Важно помнить"],
                      image_path=None, page_num=9)

    add_content_slide(prs, "Вопросы и обсуждение",
                      ["Контакты", "Спасибо"],
                      image_path=None, page_num=10)

    # Сохранение
    prs.save("presentation_plan.pptx")

if __name__ == "__main__":
    main()
