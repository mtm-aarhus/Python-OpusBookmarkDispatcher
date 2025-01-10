#Connect to orchestrator
from OpenOrchestrator.orchestrator_connection.connection import OrchestratorConnection
from OpenOrchestrator.database.queues import QueueElement
import os
import json

orchestrator_connection = OrchestratorConnection("OpusBookMarkPerformer", os.getenv('OpenOrchestratorSQL'),os.getenv('OpenOrchestratorKey'), None)
queue_json = {"Bookmark": "3HXTWFDL2HCN3HDY0K5B6SKWB", 
              "Filnavn": "FOA132 - Fremregnet løn - Natur og Miljø", 
              "SharePointMappeLink": "https://aarhuskommune.sharepoint.com/Teams/tea-teamsite11333/Delte%20dokumenter/Forms/AllItems.aspx?id=%2FTeams%2Ftea%2Dteamsite11333%2FDelte%20dokumenter%2FOPUS%20L%C3%B8nrapporter%20%2D%20MTM%20p%C3%A5%20tv%C3%A6rs&viewid=646a2b8d%2D2266%2D41ed%2D9593%2D549d6e10ea70", 
              "Dagligt (Ja/Nej)": "ja", 
              "MånedsSlut (Ja/Nej)": None, 
              "MånedsStart (Ja/Nej)": None, 
              "Årlig (Ja/Nej)": None, 
              "Ansvarlig i økonomi": "Martin Pedersen", 
              "Rapportype": "FOA132", 
              "Fulde link til Opus": "=CONCATENATE(\"https://portal-k1-nc-21.kmd.dk/irj/servlet/prt/portal/prtroot/pcd!3aportal_content!2fcom.sap.pct!2fplatform_add_ons!2fcom.sap.ip.bi!2fiViews!2fcom.sap.ip.bi.bex?BOOKMARK=\",A44)", 
              "Kolonne1": "=HYPERLINK(Tabel1[[#This Row],[Fulde link til Opus]],\"LINK\")", 
              "Kommentar": None}
orchestrator_connection.create_queue_element("OpusBookmarkQueue", "Test", json.dumps(queue_json))


