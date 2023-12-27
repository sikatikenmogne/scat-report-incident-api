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
from IncidentReportPresentation import IncidentReportPresentation
import json


def convert_file_to_pdf(file_name):
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

        enterprise_logo = 'Picture3.png'

        incident_presentation = IncidentReportPresentation(Presentation(), enterprise_logo)

        incident_site = 'SOCAVER'
        incident_report_edition_date = '01/06/2022'
        incident_title = 'ACCIDENT GRAVE DE CHUTE DE CUBITAINER D’UN CHARIOT SUR M. MABIA DU 30/05/2022'

        summary = [
            'Le contexte',
            'L’équipe',
            'La méthodologie',
            'Evènement',
            'Les recommandations',
        ]

        slide3_content = "Suite à la mise en service du nouveau réseau d’eau de refroidissement 30°C par le service Maintenance, il y ‘a eu une baisse de niveau d’eau dans la bâche qui a nécessité un appoint d’eau traitée (1000 litres d’eau et 25 litres de NALCO trac 102 ). C’est dans cette optique que deux cubitainers d’eau traitée ont été déposés dans la salle machine à coté de la station des pompes. Un des deux cubitainers a été utilisé samedi et l’autre(qui n’a pas de support en palette bois en dessous) n’a pas été utilisé en attente d’un éventuel appoint. "

        team_members = [
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
        incident_presentation.set_front_page(incident_site, incident_report_edition_date, incident_title, 'SYNTHESE DE LA RECHERCHE DES CAUSES')

        # Slide 2
        incident_presentation.set_summary_slide('Recherche des causes', summary)

        # Slide 3
        incident_presentation.set_context_slide("Contexte", slide3_content)

        # Slide 4
        incident_presentation.set_team_slide("L’équipe", team_members)

        # Slide 5
        incident_presentation.set_methodology_illustration("La méthodologie:")

        # Slide 6
        incident_presentation.add_event_slide(slide6_event_title, slide6_event_table_titles)

        # Slide 7                                                                                                                                                                                                                                                                     
        incident_presentation.set_resume_slide("Les principales recommandations")

        # Slide 8
        incident_presentation.set_appendix("ANNEXES")

        # Slide 9
        incident_presentation.set_end_slide("MERCI")

        # Footer
        incident_presentation.paginate_presentation()
        incident_presentation.set_footer_logo_to_slides()

        output_filename = incident_title.replace(' ', '-').replace('/', '-').replace('.', '-') + '-' + datetime.now().strftime('%H-%M-%S')

        file_path = app.root_path + '/files/pptx/' + output_filename + '.pptx'

        if filetype == 'pdf':
            incident_presentation.save(file_path)            
            convert_file_to_pdf(file_path)

            return send_file(app.root_path + '/files/pdf' + '/' + output_filename + ".pdf", as_attachment=True)
        elif filetype == 'pptx':
            incident_presentation.save(file_path)

            return send_file(file_path, as_attachment=True)
    else:
        return "Type de fichier non pris en charge", 400
