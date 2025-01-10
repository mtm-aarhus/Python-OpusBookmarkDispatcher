#Connect to orchestrator
from OpenOrchestrator.orchestrator_connection.connection import OrchestratorConnection
from OpenOrchestrator.database.queues import QueueElement
import os
import json

orchestrator_connection = OrchestratorConnection("OpusBookMarkPerformer", os.getenv('OpenOrchestratorSQL'),os.getenv('OpenOrchestratorKey'), None)
queue_json = {"Bookmark": "5JMRTEJLJFVFG81GOU3C0GV6M", 
              "Filnavn": "FIB045 - Natur og Miljø - Forbrug 2023 og 2024", 
              "SharePointMappeLink": "https://aarhuskommune.sharepoint.com/Teams/tea-teamsite11333/Delte%20dokumenter/Forms/AllItems.aspx?id=%2FTeams%2Ftea%2Dteamsite11333%2FDelte%20dokumenter%2FNatur%20og%20Milj%C3%B8%20%2D%20BI%20rapporter&viewid=646a2b8d%2D2266%2D41ed%2D9593%2D549d6e10ea70", 
              "Dagligt (Ja/Nej)": "Ja", 
              "MånedsSlut (Ja/Nej)": None, 
              "MånedsStart (Ja/Nej)": None, 
              "Årlig (Ja/Nej)": None, 
              "Ansvarlig i økonomi": "Martin Pedersen", 
              "Rapportype": "FIB045", 
              "Fulde link til Opus": "https://portal-k1-nc-21.kmd.dk/irj/servlet/prt/portal/prtroot/pcd!3aportal_content!2fcom.sap.pct!2fplatform_add_ons!2fcom.sap.ip.bi!2fiViews!2fcom.sap.ip.bi.bex?BOOKMARK=5JMRTEJLJFVFG81GOU3C0GV6M", 
              "Kolonne1": "=HYPERLINK(Tabel1[[#This Row],[Fulde link til Opus]],\"LINK\")", 
              "Kommentar": "Adgang"}
orchestrator_connection.create_queue_element("OpusBookmarkQueue", "Test", json.dumps(queue_json))


