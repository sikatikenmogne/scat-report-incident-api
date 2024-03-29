from datetime import datetime
import os
import subprocess
from pptx import Presentation
from pptx.dml.color import RGBColor
from flask import Flask, request, send_file
from flask import after_this_request
from IncidentReportPresentation import IncidentReportPresentation
import json


def convert_file_to_pdf(file_name):
    try:
        subprocess.run(['libreoffice', '--headless', '--convert-to', 'pdf:writer_pdf_Export', file_name, '--outdir', app.root_path + '/files/pdf'], check=True)
        # print(f"Le fichier {file_name} a été converti avec succès en PDF.")
    except subprocess.CalledProcessError as e:
        print(f"Une erreur s'est produite lors de la conversion du fichier {file_name} : {e}")

    # return '/files/pdf/' + file_name

def check_if_colorable(events_data, direct_causes_security_limit = 14):
    count = 0
    for event_data in events_data:
        if event_data['directCauses']:
            for direct_cause in event_data['directCauses']:
                count += 1

    is_ok = count <= direct_causes_security_limit

    # print(f"{count} <= {direct_causes_security_limit} ==> {is_ok}")

    return is_ok


app = Flask(__name__)
@app.route('/api/<string:filetype>', methods=['POST'])
def create_presentation(filetype):
                
    @after_this_request
    def delete_file(response):
        try:
            # print(file_path)
            os.remove(file_path)
        except Exception as error:
            app.logger.error("Erreur lors de la suppression du fichier : %s", error)
        return response


    if filetype == 'pptx' or filetype == 'pdf':

        enterprise_logo = 'logo-boissons-du-cameroun.png'

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

        incident_context = " "

        event_table_headers = ["Causes immédiates", "Causes fondamentales", "Actions", "Responsabilités", "Délais"]


        # Récupérer les données JSON de la requête
        json_data = request.get_json()

        data = json.loads(json_data)[0]

        incident_site = data['site']['name']

        now = datetime.now()
        incident_report_edition_date = now.strftime('%d/%m/%Y')

        events_data = data['events']

        incident_presentation = IncidentReportPresentation(Presentation(), enterprise_logo, event_table_headers)


        if type(data['description']) == str:
            incident_title = data['description']
        # Slide 1
        incident_presentation.set_front_page(incident_site, incident_report_edition_date, incident_title, 'SYNTHESE DE LA RECHERCHE DES CAUSES')

        # Slide 2
        incident_presentation.set_summary_slide('Recherche des causes', summary)
        
        if type(data['context']) == str:
            incident_context = data['context']

            line_count = incident_presentation.calculate_line_count(incident_context, 110)

            if(line_count > 27):
                split_context_values = incident_presentation.split_string(incident_context, 3000) 

                for split_context_value in split_context_values:
                    incident_presentation.set_context_slide("Contexte", split_context_value)
            else:       
                # Slide 3
                incident_presentation.set_context_slide("Contexte", incident_context)

        # Slide 4
        if(data['team'] != None and data['team'] != [] ):
            team = list(data['team'].values())
            incident_presentation.set_team_slide("L’équipe", team)

        # Slide 5
        incident_presentation.set_methodology_illustration("La méthodologie:")

        for event_data in events_data:
            if event_data['directCauses']:
                for direct_cause in event_data['directCauses']:
                        incident_presentation.add_one_direct_cause_event_slide(direct_cause, event_data)


        # Slide 7                                                                                                                                                                                                                                                                     
        # incident_presentation.set_resume_slide("Les principales recommandations")

        # Slide 8
        incident_presentation.set_appendix("ANNEXES")

        # Slide 9
        incident_presentation.set_end_slide("MERCI")

        # Footer
        incident_presentation.paginate_presentation()
        incident_presentation.set_footer_logo_to_slides()

        output_filename = incident_title.replace(' ', '-').replace('/', '-').replace('.', '-').upper() + '-' + datetime.now().strftime('%H-%M-%S')

        file_path = app.root_path + '/files/pptx/' + output_filename + '.pptx'

        if filetype == 'pdf':
            incident_presentation.save(file_path)            
            convert_file_to_pdf(file_path)
            
            os.remove(file_path)
            file_path = app.root_path + '/files/pdf' + '/' + output_filename + ".pdf"
            return send_file(file_path, as_attachment=True)
        elif filetype == 'pptx':
            incident_presentation.save(file_path)

            return send_file(file_path, as_attachment=True)
    else:
        return "Type de fichier non pris en charge", 400
