from datetime import datetime
import subprocess
from pptx import Presentation
from pptx.dml.color import RGBColor
from pptx.enum.dml import MSO_THEME_COLOR
from pptx.util import Pt
from pptx.enum.text import PP_ALIGN
from pptx.util import Inches
from pptx.enum.text import MSO_ANCHOR, MSO_AUTO_SIZE
from flask import Flask, request, send_file
from flask import render_template, send_from_directory
import json



def paginate_presentation(prs: Presentation):
    for i, slide in enumerate(prs.slides, start=1):
        slide_number_box = slide.shapes.add_textbox(Inches(9.45), Inches(6.90), Inches(0.5), Inches(0.5))
        slide_number_frame = slide_number_box.text_frame
        slide_number = slide_number_frame.add_paragraph()
        slide_number.text = "Page " + str(i)
        slide_number.font.size = Pt(12)
        slide_number.font.name = 'Calibri'
        slide_number.alignment = PP_ALIGN.RIGHT

def set_footer_logo_slides(prs: Presentation, enterprise_logo):
    for i, slide in enumerate(prs.slides, start=1):
        slide.shapes.add_picture(enterprise_logo, Inches(0), Inches(6.90), Inches(0.60), Inches(0.60))


def add_end_slide(prs: Presentation, end_title):
    blank_slide_layout = prs.slide_layouts[6]  # 6 est une diapositive vide
    slide6 = prs.slides.add_slide(blank_slide_layout)

    # left = top = width = height = Inches(1)
    title_slide6_shape = slide6.shapes.add_textbox(Inches(2.8), Inches(2.75), Inches(4.5), Inches(2))

    p = title_slide6_shape.text_frame.paragraphs[0] 

    p.text = end_title
    p.font.color.rgb = RGBColor(0x00, 0x70, 0xc0)
    p.font.bold = True
    p.font.size = Pt(115)


def add_appendix_slide(prs: Presentation, text):
    title_slide_layout = prs.slide_layouts[0]
    slide5 = prs.slides.add_slide(title_slide_layout)

    title_slide5_shape = slide5.placeholders[0]

    title_slide5_shape.text = text
    title_slide5_shape.text_frame.paragraphs[0].font.color.rgb = RGBColor(0xB2, 0x22, 0x22)


def add_resume_slide(prs: Presentation, slide_title):

    # Ajouter une diapositive vide (la septième diapositive)
    slide_layout = prs.slide_layouts[6]  # 6 est une diapositive vide
    slide9 = prs.slides.add_slide(slide_layout)

    # Ajouter un titre à la diapositive
    title_box = slide9.shapes.add_textbox(Inches(0), Inches(0), Inches(10), Inches(1))
    title_frame = title_box.text_frame
    title_frame.word_wrap = True  # Wrap text in shape
    title = title_frame.paragraphs[0]
    title.text = slide_title
    title.font.size = Pt(30)
    title.font.name = 'Calibri'
    title.font.color.rgb = RGBColor(0xB2, 0x22, 0x22)  # Dark Red
    title.alignment = PP_ALIGN.CENTER

    slide9.shapes.add_textbox(Inches(0.25), Inches(1), Inches(9.5), Inches(6))


def add_scat_table_slide(prs: Presentation, event_title, event_table_titles):
    # Ajouter une diapositive vide (la septième diapositive)
    slide_layout = prs.slide_layouts[6]  # 6 est une diapositive vide
    slide8 = prs.slides.add_slide(slide_layout)

    # Ajouter un titre à la diapositive
    title_box = slide8.shapes.add_textbox(Inches(0), Inches(0), Inches(10), Inches(1))
    title_frame = title_box.text_frame
    title_frame.word_wrap = True  # Wrap text in shape
    title = title_frame.paragraphs[0]
    title.text = "Evènement : " + event_title
    title.font.size = Pt(30)
    title.font.name = 'Calibri'
    title.font.color.rgb = RGBColor(0xB2, 0x22, 0x22)  # Dark Red
    title.alignment = PP_ALIGN.CENTER

    # Ajouter un tableau en bas de la diapositive
    table = slide8.shapes.add_table(2, 5, Inches(0.25), Inches(1.15), Inches(9.5), Inches(10.75)).table  # Added 0.5 inch margin to the left, right and bottom

    # Définir les titres des colonnes
    for i in range(5):
        cell = table.cell(0, i)
        cell.text_frame.auto_size = True
        cell.text = event_table_titles[i]
        cell.text_frame.paragraphs[0].alignment = PP_ALIGN.CENTER  # Center column titles

    table.rows[0].height = Inches(0.5)  # Définir la hauteur à 1 pouce


def add_presentation_slide(prs: Presentation, slide_title):
    slide_layout = prs.slide_layouts[6]  # 6 is a blank slide
    slide7 = prs.slides.add_slide(slide_layout)

    # Add a title to the slide
    title_box = slide7.shapes.add_textbox(Inches(2), Inches(0.25), Inches(6), Inches(1))
    title_frame = title_box.text_frame
    title = title_frame.paragraphs[0]
    title.text = slide_title
    title.font.size = Pt(48)
    title.font.name = 'Tahoma'
    title.font.color.rgb = RGBColor(0xB2, 0x22, 0x22)  # Bright Red

    # left = top = width = height = Inches(1)

    # Add the images to the slide
    image1 = slide7.shapes.add_picture('Picture1.png', Inches(4), Inches(1.75), Inches(6), Inches(2))
    image2 = slide7.shapes.add_picture('Picture2.png', Inches(0), Inches(1), Inches(4.5), Inches(5.5))

    # Add text below each image
    text1_box = slide7.shapes.add_textbox(Inches(0.5), Inches(6), Inches(5), Inches(1))
    text1_frame = text1_box.text_frame
    text1 = text1_frame.add_paragraph()
    text1.text = "SCAT – isrs7"
    text1.font.size = Pt(44)
    text1.font.name = 'Tahoma'
    text1.font.color.rgb = RGBColor(0xB2, 0x22, 0x22)  # Bright Red

    text2_box = slide7.shapes.add_textbox(Inches(4.25), Inches(3.5), Inches(6), Inches(1))
    text2_frame = text2_box.text_frame
    text2 = text2_frame.add_paragraph()
    text2.text = "ARBRE DES CAUSES"
    text2.font.size = Pt(44)
    text2.font.name = 'Tahoma'
    text2.font.color.rgb = RGBColor(0xB2, 0x22, 0x22)  # Bright Red

def add_team_slide(prs: Presentation, slide_title, team_members):
    # prs.slide_layouts[6] => blank_slide_layout
    blank_slide_layout = prs.slide_layouts[6]

    slide4 = prs.slides.add_slide(blank_slide_layout)
    slide4_shapes = slide4.shapes

    left = top = width = height = Inches(1)
    txBox2 = slide4.shapes.add_textbox(Inches(0.3), Inches(0), Inches(4), Inches(0.4))
    tf = txBox2.text_frame
    p = tf.paragraphs[0]

    p.text = slide_title
    p.font.color.rgb = RGBColor(0xB2, 0x22, 0x22)
    p.font.alignment = PP_ALIGN.RIGHT
    p.font.bold = True
    p.font.size = Pt(36)

    # left = top = width = height = Inches(1)
    # Ajouter du contenu à la diapositive sous forme de deux colonnes
    txBox1 = slide4.shapes.add_textbox(Inches(0), Inches(0.5), Inches(4.5), Inches(6))
    tf1 = txBox1.text_frame
    tf1.word_wrap = True
    tf1.auto_size = True

    txBox2 = slide4.shapes.add_textbox(Inches(4), Inches(0.5), Inches(6), Inches(6))
    tf2 = txBox2.text_frame
    tf2.word_wrap = True
    tf2.auto_size = True

    # Ajouter un nouveau paragraphe pour chaque membre de l'équipe
    for i, member in enumerate(team_members):
        if i < len(team_members) / 2:
            tf = tf1
        else:
            tf = tf2
        p = tf.add_paragraph()
        p.text = "• " + member  # Ajouter une puce
        p.font.size = Pt(16)
        p.level = 1  # Ajouter une puce
        p.space_before = Pt(21)

def add_context_description_slide(prs: Presentation, slide_title, slide_content):
    blank_slide_layout = prs.slide_layouts[6]
    slide = prs.slides.add_slide(blank_slide_layout)

    left = top = width = height = Inches(1)
    txBox = slide.shapes.add_textbox(Inches(0.3), Inches(0), Inches(2), Inches(0.4))
    tf = txBox.text_frame

    p = tf.paragraphs[0]
    tf.alignment = PP_ALIGN.CENTER

    p.text = slide_title
    p.font.size = Pt(36)
    p.font.bold = True
    p.font.name = 'Calibri'
    p.font.color.rgb = RGBColor(0xB2, 0x22, 0x22)


    txBox = slide.shapes.add_textbox(Inches(0.25), Inches(0.5), Inches(9.5), Inches(6.4))
    tf = txBox.text_frame
    tf.word_wrap = True

    p = tf.add_paragraph()

    p.text = slide_content
    p.font.size = Pt(16)
    p.alignment = PP_ALIGN.JUSTIFY
    txBox.word_wrap = MSO_AUTO_SIZE.SHAPE_TO_FIT_TEXT


def add_summary_slide(prs: Presentation, summary_title, parts):
    # prs.slide_layouts[1] => bullet_slide_layout
    bullet_slide_layout = prs.slide_layouts[1]

    slide2 = prs.slides.add_slide(bullet_slide_layout)
    slide2_shapes = slide2.shapes

    txBox2 = slide2.shapes.add_textbox(Inches(0.3), Inches(0), Inches(4), Inches(0.4))


    title_slide2_shape = slide2_shapes.placeholders[0]
    title_slide2_shape.text = summary_title
    title_slide2_shape.text_frame.paragraphs[0].font.color.rgb = RGBColor(0xB2, 0x22, 0x22)

    slide2_shape_content = slide2_shapes.placeholders[1]

    slide2_shape_content_text_frame = slide2_shape_content.text_frame
    # slide2_shape_content_text_frame.auto_size = MSO_AUTO_SIZE.SHAPE_TO_FIT_TEXT
    # slide2_shape_content_text_frame.alignment = PP_ALIGN.CENTER

    # Obtenir la largeur et la hauteur de la diapositive
    slide_width =  prs.slide_width
    slide_height = prs.slide_height

    # Obtenir la largeur et la hauteur de la forme
    shape_width = slide2_shape_content.width
    shape_height = slide2_shape_content.height

    # Redimensionner la TextFrame
    slide2_shape_content.width = 5000000
    slide2_shape_content.height = 5000000

    # Déplacer la TextFrame
    slide2_shape_content.left = 2500000
    slide2_shape_content.top  = 1500000

    for i, member in enumerate(parts):
        p2 = slide2_shape_content_text_frame.add_paragraph()
        p2.text = member
        p2.font.color.rgb = RGBColor(0xB2, 0x22, 0x22)


    # Supprimer uniquement les puces mais garder le texte
    for paragraph in slide2_shape_content_text_frame.paragraphs:
        paragraph.level = 0
        paragraph.space_before = None
        paragraph.space_after = None
        paragraph.line_spacing = None
        paragraph.alignment = None

def set_front_cover_presentation(prs: Presentation, incident_site, incident_report_edition_date, incident_title, incident_title_sub_text, image_uri = ''):
    # prs.slide_layouts[8] => picture_with_caption_layout

    placeholder_id = 2

    if image_uri != '' :
        selected_layout = prs.slide_layouts[8]
    else :
        selected_layout = prs.slide_layouts[0]
        placeholder_id = 1

        
    slide1 = prs.slides.add_slide(selected_layout)
    slide1_shapes = slide1.shapes

    title_box = slide1.shapes.add_textbox(Inches(2.5), Inches(0.10), Inches(5), Inches(0.5))
    title_frame = title_box.text_frame
    title_frame.word_wrap = True  # Wrap text in shape
    title = title_frame.paragraphs[0]
    title.text = incident_site + " LE " + incident_report_edition_date
    title.font.size = Pt(18)
    title.font.name = 'Calibri'
    title.font.color.rgb = RGBColor(0x00, 0x70, 0xc0)  # Blue
    title.alignment = PP_ALIGN.CENTER
    title.font.bold = True


    # picture_with_caption_layout: placeholders[0] => Title
    title_shape = slide1_shapes.placeholders[0]

    tf = title_shape.text_frame.add_paragraph()
    tf.alignment = PP_ALIGN.CENTER

    run = tf.add_run()
    run.text = incident_title

    font = run.font
    font.color.rgb = RGBColor(0xB2, 0x22, 0x22)

    # picture_with_caption_layout: placeholders[2] => Text Placeholder
    text_shape = slide1_shapes.placeholders[placeholder_id]

    tf = text_shape.text_frame.add_paragraph()
    tf.alignment = PP_ALIGN.CENTER

    run = tf.add_run()
    run.text = incident_title_sub_text

    font = run.font
    font.name = 'Calibri'
    font.size = Pt(18)
    font.bold = True
    font.italic = None  # cause value to be inherited from theme
    font.color.rgb = RGBColor(0xB2, 0x22, 0x22)

    # txBox2 = slide1.shapes.add_textbox(Inches(0.3), Inches(0), Inches(4), Inches(0.4))


def convert_pptx_to_pdf(file_name):
    try:
        subprocess.run(['libreoffice', '--headless', '--convert-to', 'pdf', file_name, '--outdir', app.root_path + '/files/pdf'], check=True)
        # print(f"Le fichier {file_name} a été converti avec succès en PDF.")
    except subprocess.CalledProcessError as e:
        print(f"Une erreur s'est produite lors de la conversion du fichier {file_name} : {e}")

    # return '/files/pdf/' + file_name


app = Flask(__name__)
@app.route('/api/<string:filetype>', methods=['POST'])
def create_presentation(filetype):
    if filetype == 'pptx' or filetype == 'pdf':

        prs = Presentation()
            
        enterprise_logo = 'Picture3.png'

        incident_site = 'SOCAVER'
        incident_report_edition_date = '01/06/2022'
        incident_title = 'ACCIDENT GRAVE DE CHUTE DE CUBITAINER D’UN CHARIOT SUR M. MABIA DU 30/05/2022'

        slide2_parts = [
            'Le contexte',
            'L’équipe',
            'La méthodologie',
            'Evènement',
            'Les recommandations',
        ]

        slide3_content = "Suite à la mise en service du nouveau réseau d’eau de refroidissement 30°C par le service Maintenance, il y ‘a eu une baisse de niveau d’eau dans la bâche qui a nécessité un appoint d’eau traitée (1000 litres d’eau et 25 litres de NALCO trac 102 ). C’est dans cette optique que deux cubitainers d’eau traitée ont été déposés dans la salle machine à coté de la station des pompes. Un des deux cubitainers a été utilisé samedi et l’autre(qui n’a pas de support en palette bois en dessous) n’a pas été utilisé en attente d’un éventuel appoint. "

        slide4_team_members = [
            "KAMENI Vincent: Directeur des ventes",
            "MPONDO MBOKA Régis: Directeur QHSE",
            "KAPING Stephan: Chef Division Production",
            "NGUIMKENG Arnold: Chef de Division RH",
            "POUTCHEU Emmanuel: Chef d’atelier Préforme et Casier",
            "WANDJI Valery: Responsable HSSE",
            "KOSSI David: Chef de groupe presse",
            "MANI Joseph: Cariste",
            "TAMO Florence: Infirmière Chef",
            "TAMEFO Gabriel: Chef de Groupe laboratoire préforme",
            "NDJELLE Lavie: Chef de Groupe presse",
            "KOUATCHO: Laborantin",
            "GUIDIONGUENA Albert: Cariste",
            "WAFO Freddy: Chef de service Senior Maintenance",
            "TEMFACK Rodrigue: Chef de service Pyrométrie et maintenance des presses",
            "Dr TCHIMOU: Médecin référent"
        ]

        slide6_event_title = 'Heurt et Ecrasement de M. MABIA par un Cubitainer avec une charge d’environ une tonne'
        slide6_event_table_titles = ["Causes immédiates", "Causes fondamentales", "Actions", "Responsabilités", "Délais"]

        # Récupérer les données JSON de la requête
        data = request.get_json()

        # Slide 1
        set_front_cover_presentation(prs, incident_site, incident_report_edition_date, incident_title, 'SYNTHESE DE LA RECHERCHE DES CAUSES')

        # Slide 2
        add_summary_slide(prs, 'Recherche des causes', slide2_parts)

        # Slide 3
        add_context_description_slide(prs, "Contexte", slide3_content)

        # Slide 4
        add_team_slide(prs, "L’équipe", slide4_team_members)

        # Slide 5
        add_presentation_slide(prs, "La méthodologie:")

        # Slide 6
        add_scat_table_slide(prs, slide6_event_title, slide6_event_table_titles)

        # Slide 7                                                                                                                                                                                                                                                                     
        add_resume_slide(prs, "Les principales recommandations")

        # Slide 8
        add_appendix_slide(prs, "ANNEXES")

        # Slide 9
        add_end_slide(prs, "MERCI")

        # Footer
        paginate_presentation(prs)
        set_footer_logo_slides(prs, enterprise_logo)

        output_filename = incident_title.replace(' ', '-').replace('/', '-') + '-' + datetime.now().strftime('%H-%M-%S')

        # print(app.static_folder + '/' + output_filename + '.pptx')
        file_out = app.root_path + '/files/pptx/' + output_filename + '.pptx'

        if filetype == 'pdf':
            prs.save(file_out)
            convert_pptx_to_pdf(file_out)
            # Instancier un objet Presentation qui représente un fichier PPTX
            return send_file(app.root_path + '/files/pdf' + '/' + output_filename + ".pdf", as_attachment=True)
        elif filetype == 'pptx':
            prs.save(file_out)
            return send_file(file_out, as_attachment=True)
    else:
        return "Type de fichier non pris en charge", 400
