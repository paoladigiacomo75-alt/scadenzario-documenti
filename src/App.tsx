import { useState, useMemo } from "react";
import * as XLSX from "xlsx";

const INIT_DOCS = [{"id":"ASS-001","cat":"Assicurazioni","nome":"Polizza RCT/RCO","durata":6,"dataEmissione":"2026-01-31"},{"id":"ASS-002","cat":"Assicurazioni","nome":"Polizza RCT/RCO - Secondo Rischio / Following Form","durata":6,"dataEmissione":"2025-12-13"},{"id":"ASS-003","cat":"Assicurazioni","nome":"Polizza Professionale","durata":12,"dataEmissione":"2025-12-13"},{"id":"ASS-004","cat":"Assicurazioni","nome":"Polizze mezzi","durata":12,"dataEmissione":"2025-09-30"},{"id":"ASS-005","cat":"Assicurazioni","nome":"Polizza Infortuni","durata":12,"dataEmissione":"2025-05-12"},{"id":"ASS-006","cat":"Assicurazioni","nome":"Dichiarazione di Assicurazione (RCT/RCO)","durata":6,"dataEmissione":"2026-02-16"},{"id":"ASS-007","cat":"Assicurazioni","nome":"Declaratoria rinuncia alla rivalsa ACS DOBFAR","durata":6,"dataEmissione":"2025-07-22"},{"id":"BAN-001","cat":"Bancario/Finanziario","nome":"Lettera di attestazione del Conto corrente bancario BPER","durata":6,"dataEmissione":"2025-02-04"},{"id":"BAN-002","cat":"Bancario/Finanziario","nome":"Lettera di attestazione del Conto corrente bancario INTESA","durata":6,"dataEmissione":"2025-02-04"},{"id":"BAN-003","cat":"Bancario/Finanziario","nome":"Lettera di attestazione del Conto corrente bancario BNL","durata":6,"dataEmissione":"2023-05-16"},{"id":"CERT-001","cat":"Certificazioni","nome":"ISO 9001","durata":36,"dataEmissione":"2025-11-26"},{"id":"CERT-002","cat":"Certificazioni","nome":"ISO 14001","durata":36,"dataEmissione":"2025-12-13"},{"id":"CERT-003","cat":"Certificazioni","nome":"ISO 45001","durata":36,"dataEmissione":"2025-12-13"},{"id":"CERT-004","cat":"Certificazioni","nome":"SOA","durata":60,"dataEmissione":"2025-09-30"},{"id":"CERT-005","cat":"Certificazioni","nome":"Attestato RLS DE FLORIO ADOLFO","durata":0,"dataEmissione":"2023-07-28"},{"id":"CERT-005_1","cat":"Certificazioni","nome":"Attestato RLS MARTANO STEFANO","durata":0,"dataEmissione":"2026-01-07"},{"id":"CERT-005_2","cat":"Certificazioni","nome":"Attestato RLS PELLICORO MARCO","durata":0,"dataEmissione":"2026-01-07"},{"id":"CERT-006","cat":"Certificazioni","nome":"Attestato RSPP","durata":0,"dataEmissione":"2026-01-07"},{"id":"CERT-007","cat":"Certificazioni","nome":"Valutazione di sostenibilità (ESG SCORECARD)","durata":12,"dataEmissione":"2025-04-04"},{"id":"CERT-008","cat":"Certificazioni","nome":"Valutazione di sostenibilità (OPEN-ES CARD)","durata":12,"dataEmissione":"2026-02-10"},{"id":"DOC-001","cat":"Legale/Societario","nome":"Dichiarazione sulla catena di controllo e titolarità effettiva","durata":6,"dataEmissione":"2025-10-03"},{"id":"DOC-002","cat":"Legale/Societario","nome":"Certificato Attribuzione P.IVA","durata":6,"dataEmissione":"2025-12-02"},{"id":"DOC-003","cat":"Legale/Societario","nome":"Statuto Societario","durata":0,"dataEmissione":"2025-12-02"},{"id":"DOC-004","cat":"Legale/Societario","nome":"Documento Identità M. La Gioia","durata":130,"dataEmissione":"2015-10-09"},{"id":"DOC-005","cat":"Legale/Societario","nome":"Documento Identità S. Schiavoni","durata":0,"dataEmissione":"2026-01-20"},{"id":"DOC-006","cat":"Legale/Societario","nome":"Brochure di presentazione ITA","durata":0,"dataEmissione":"2025-11-14"},{"id":"DOC-006_1","cat":"Legale/Societario","nome":"Brochure di presentazione ENG","durata":0,"dataEmissione":"2025-11-14"},{"id":"DOC-007","cat":"Legale/Societario","nome":"C.C.I.A.A. & R.E.A. (Visura Camerale)","durata":6,"dataEmissione":"2026-01-07"},{"id":"DOC-008","cat":"Legale/Societario","nome":"CERTIFICATO CCIAA","durata":6,"dataEmissione":"2026-10-03"},{"id":"DOC-009","cat":"Legale/Societario","nome":"Certificato dei Carichi pendenti Massimo La Gioia","durata":6,"dataEmissione":"2025-12-02"},{"id":"DOC-010","cat":"Legale/Societario","nome":"Certificato del Casellario Giudiziale Massimo La Gioia","durata":6,"dataEmissione":"2025-12-02"},{"id":"DOC-011","cat":"Legale/Societario","nome":"Carichi pendenti Stefano Schiavoni","durata":6,"dataEmissione":"2026-01-16"},{"id":"DOC-012","cat":"Legale/Societario","nome":"Casellario Giudiziale Stefano Schiavoni","durata":6,"dataEmissione":"2026-01-16"},{"id":"DOC-013","cat":"Legale/Societario","nome":"Dichiarazione  Carichi pendenti (LR dichiara per tutti gli apicali)","durata":6,"dataEmissione":""},{"id":"DOC-014","cat":"Legale/Societario","nome":"Dichiarazione Casellario Giudiziale","durata":6,"dataEmissione":""},{"id":"DOC-015","cat":"Legale/Societario","nome":"Scheda OPC","durata":0,"dataEmissione":""},{"id":"DOC-015_1","cat":"Legale/Societario","nome":"Dichiarazione Parti Correlate","durata":0,"dataEmissione":""},{"id":"DOC-016","cat":"Legale/Societario","nome":"Analisi Controparte","durata":0,"dataEmissione":""},{"id":"DOC-017","cat":"Legale/Societario","nome":"Dichiarazione assenza conflitto di interessi","durata":0,"dataEmissione":""},{"id":"DOC-018","cat":"Legale/Societario","nome":"Rating di legalità","durata":0,"dataEmissione":""},{"id":"DOC-018_1","cat":"Legale/Societario","nome":"Credit Rating","durata":12,"dataEmissione":"2025-10-21"},{"id":"DOC-019","cat":"Legale/Societario","nome":"DGUE","durata":6,"dataEmissione":""},{"id":"DOC-020","cat":"Legale/Societario","nome":"Dichiarazione Antimafia (LR per tutti)","durata":6,"dataEmissione":"2025-07-11"},{"id":"DOC-020_1","cat":"Legale/Societario","nome":"Dichiarazione di esenzione dall’obbligo della presentazione della documentazione Antimafia","durata":6,"dataEmissione":"2025-10-23"},{"id":"DOC-020_10","cat":"Legale/Societario","nome":"Dichiarazione Antimafia SALVATORE VARRIALE","durata":6,"dataEmissione":"2025-10-03"},{"id":"DOC-020_11","cat":"Legale/Societario","nome":"Dichiarazione Antimafia PAOLO NAGAR","durata":6,"dataEmissione":"2025-07-29"},{"id":"DOC-020_12","cat":"Legale/Societario","nome":"Dichiarazione Antimafia GIOVANNI PELUSO","durata":6,"dataEmissione":"2025-10-02"},{"id":"DOC-020_13","cat":"Legale/Societario","nome":"Dichiarazione Antimafia MENA MENZIONE","durata":6,"dataEmissione":"2025-08-01"},{"id":"DOC-020_14","cat":"Legale/Societario","nome":"Dichiarazione Antimafia PIO QUARANTA","durata":6,"dataEmissione":"2025-08-04"},{"id":"DOC-020_15","cat":"Legale/Societario","nome":"Dichiarazione Antimafia MICHELE LIZZIO","durata":6,"dataEmissione":"2025-08-04"},{"id":"DOC-020_16","cat":"Legale/Societario","nome":"Dichiarazione Antimafia CORRADO GIANCASPRO","durata":6,"dataEmissione":"2025-07-04"},{"id":"DOC-020_17","cat":"Legale/Societario","nome":"Dichiarazione Antimafia MARCO GRAZIANO","durata":6,"dataEmissione":"2025-07-29"},{"id":"DOC-020_18","cat":"Legale/Societario","nome":"Dichiarazione Antimafia FRANCESCO CARUSO","durata":6,"dataEmissione":"2025-07-29"},{"id":"DOC-020_19","cat":"Legale/Societario","nome":"Dichiarazione Antimafia FRANCESCO MARTINO","durata":6,"dataEmissione":"2025-07-29"},{"id":"DOC-020_2","cat":"Legale/Societario","nome":"Dichiarazione Antimafia MASSIMO LA GIOIA","durata":6,"dataEmissione":"2025-07-29"},{"id":"DOC-020_20","cat":"Legale/Societario","nome":"Dichiarazione Antimafia VALENTINA GIANNI'","durata":6,"dataEmissione":"2025-08-04"},{"id":"DOC-020_21","cat":"Legale/Societario","nome":"Dichiarazione Antimafia IOLE ANNA SAVINI","durata":6,"dataEmissione":"2025-07-29"},{"id":"DOC-020_3","cat":"Legale/Societario","nome":"Dichiarazione Antimafia STEFANO SCHIAVONI","durata":6,"dataEmissione":"2025-07-29"},{"id":"DOC-020_4","cat":"Legale/Societario","nome":"Dichiarazione Antimafia DAVID GIORDANO ARMANINI","durata":6,"dataEmissione":"2025-07-29"},{"id":"DOC-020_5","cat":"Legale/Societario","nome":"Dichiarazione Antimafia MARIO PERCUOCO","durata":6,"dataEmissione":"2025-07-29"},{"id":"DOC-020_6","cat":"Legale/Societario","nome":"Dichiarazione Antimafia GIANLUCA SCUDIERO","durata":6,"dataEmissione":"2025-08-06"},{"id":"DOC-020_7","cat":"Legale/Societario","nome":"Dichiarazione Antimafia PIETRO SILVESTRO  CAMINITI","durata":6,"dataEmissione":"2025-07-30"},{"id":"DOC-020_8","cat":"Legale/Societario","nome":"Dichiarazione Antimafia ANNALISA EBREO","durata":6,"dataEmissione":"2025-08-05"},{"id":"DOC-020_9","cat":"Legale/Societario","nome":"Dichiarazione Antimafia GOFFREDO HINNA DANESI","durata":6,"dataEmissione":"2025-08-01"},{"id":"DOC-021","cat":"Legale/Societario","nome":"Dichiarazione iscrizione Whitelist","durata":6,"dataEmissione":"2026-02-17"},{"id":"FIS-001","cat":"Fiscale","nome":"DURC","durata":4,"dataEmissione":"2026-02-13"},{"id":"FIS-002","cat":"Fiscale","nome":"Dichiarazione di residenza Fiscale","durata":6,"dataEmissione":"2026-01-20"},{"id":"FIS-002_1","cat":"Fiscale","nome":"Certificato di residenza Fiscale","durata":6,"dataEmissione":""},{"id":"FIS-003","cat":"Fiscale","nome":"Bilancio Ultimo Esercizio","durata":12,"dataEmissione":"2024-12-31"},{"id":"FIS-004","cat":"Fiscale","nome":"Bilanci ultimi 3 anni","durata":0,"dataEmissione":""},{"id":"FIS-004_1","cat":"Fiscale","nome":"Dichiarazione fatturato","durata":6,"dataEmissione":""},{"id":"FIS-005","cat":"Fiscale","nome":"DURF","durata":4,"dataEmissione":"2025-11-14"},{"id":"PRIV-001","cat":"Privacy/Compliance","nome":"Modello 231","durata":0,"dataEmissione":""},{"id":"PRIV-002","cat":"Privacy/Compliance","nome":"Codice Etico ITA","durata":0,"dataEmissione":""},{"id":"PRIV-003","cat":"Privacy/Compliance","nome":"Codice Etico ENG","durata":0,"dataEmissione":""},{"id":"PRIV-004","cat":"Privacy/Compliance","nome":"Informativa Privacy","durata":0,"dataEmissione":""},{"id":"PRIV-005","cat":"Privacy/Compliance","nome":"Nomina DPO","durata":0,"dataEmissione":""},{"id":"SIC-001","cat":"Sicurezza","nome":"DVR","durata":0,"dataEmissione":"2025-04-14"},{"id":"SIC-002","cat":"Sicurezza","nome":"Dichiarazione indici infortunistici","durata":6,"dataEmissione":""},{"id":"SIC-002_1","cat":"Sicurezza","nome":"Cruscotto Infortuni INAIL","durata":0,"dataEmissione":""},{"id":"SIC-003","cat":"Sicurezza","nome":"Nomina Medico Competente","durata":0,"dataEmissione":"2025-04-15"},{"id":"SIC-003_1","cat":"Sicurezza","nome":"Nomina Medici Coordinati","durata":0,"dataEmissione":""},{"id":"SIC-004","cat":"Sicurezza","nome":"Dichiarazione EX Art.14 dlgs 81/08 (di non essere oggetto di provvedimenti di sospensione)","durata":6,"dataEmissione":""},{"id":"SIC-005","cat":"Sicurezza","nome":"Dichiarazione organico medio annuo","durata":6,"dataEmissione":""},{"id":"SIC-006","cat":"Sicurezza","nome":"Dichiarazione contratto collettivo applicato","durata":6,"dataEmissione":""},{"id":"SIC-007","cat":"Sicurezza","nome":"Dichiarazione dei requisiti di idoneità tecnico-professionale dell’impresa appaltatrice (art. 26 del D.Lgs. 81/2008 e s.m.i.)","durata":6,"dataEmissione":"2025-12-23"},{"id":"SIC-008","cat":"Sicurezza","nome":"Nomina Addetto ANTINCENDIO","durata":0,"dataEmissione":""},{"id":"SIC-008","cat":"Sicurezza","nome":"Nomina RLS (Responsabile dei Lavoratori per la Sicurezza)","durata":0,"dataEmissione":""},{"id":"SIC-009","cat":"Sicurezza","nome":"Nomina Addetto PRIMO SOCCORSO","durata":0,"dataEmissione":""},{"id":"SIC-009","cat":"Sicurezza","nome":"Nomina RSPP (Responsabile del Servizio di Prevenzione e Protezione)","durata":0,"dataEmissione":""},{"id":"SIC-010","cat":"Sicurezza","nome":"Registro carico/scarico rifiuti","durata":0,"dataEmissione":""},{"id":"SIC-011","cat":"Sicurezza","nome":"Dichiarazione spazi confinati (DPR 177-11)","durata":0,"dataEmissione":""},{"id":"TEC-001","cat":"Tecnico/Operativo","nome":"Dichiarazione assenza cause di esclusione di cui agli art. 94, 95 e 98 del D.Lgs. 36/2023","durata":6,"dataEmissione":"2026-02-17"},{"id":"TEC-002","cat":"Tecnico/Operativo","nome":"Dichiarazione subappalti","durata":0,"dataEmissione":""},{"id":"TEC-003","cat":"Tecnico/Operativo","nome":"Lista Mezzi","durata":6,"dataEmissione":""},{"id":"TEC-004","cat":"Tecnico/Operativo","nome":"Lista referenze EPC/AT","durata":6,"dataEmissione":""},{"id":"TEC-005","cat":"Tecnico/Operativo","nome":"Lista referenze O&M/AT","durata":6,"dataEmissione":""}];
const INIT_CLIENTS = ["A2A","A5","Acciona","ACEA","ACHILLES","Acquedotto Pugliese","ACS DOBFAR","ADR (Aeoroporti di Roma)","AGCM","AGSM AIM Power","Alperia","Amazon (AWS)- Avetta","Amazon (AWS)- Salesforce","Amazon (AWS)- Lightsource","ANAC","Aquila Capital","AST Terni","Autorità di Sistema Portuale del Mare di Sardegna","Autorità di Sistema Portuale del Mare di Sicilia","Autostrade per l'Italia","Barilla","Bosch","Cabot Italia","Cefla","Cerved - Rating Agency","Contourglobal","Creditsafe","CVA Energie","CVL MANTOVA","DATA4","Deval","Dolomiti Energia","Duferco Energia","Ecovadis","Edison","EDP Renewables","EF Solare","ENEL","Engie","ENI","E-ON","EP Produzione (Biomasse Italia)","ERG","FAREVA","FER Servizi","Ferrari","Ferrari2","Ferrero","Fincantieri","Galileo","Google","Gori","Governo di Malta","GreenYellow","Heidelberg/Italcementi","Hera","Hyphen","Iberdrola","INFN - ISTITUTO NAZIONALE DI FISICA NUCLEARE","INGFERRARI","Iren","Italgas","Italgas3","Kimberly Clark","Lucchini RS","Max Mara Fashion Group","Mediaset","METINVEST","Microsoft","MM (Metropolitane Milanesi)","Nadara","Neoen","Open-es  Altenia","ORION CARBONS","Parmalat","Philipps Morris","Piaggio","Pilkington (NSG)","Pizzarotti","Pratiche.it","Proger - Alantra Solar","RAM","Renantis_NADARA","Repsol","RFI","Roma Express","ROQUETTE","Saint-Gobain","Saipem","Sapio","Saras","Sasol","Sasol4","Sigef (biomasse)","SIOT - Società  Italiana per l'Oleodotto Transalpino S.p.A.","SNAM","Società per azioni Esercizi Aeroportuali - SEA","Sofidel","Sogin","Solvay-Syensqo","Sorgenia","Statkraft","TERNA_gare","TERNA","TIM Gruppo","Tirreno Power","TRILLIUM","Vaticano","Verallia","Verbund","WARTSILA","Wepa","X-Elio","Zucchetti/TRINSEO/ALTUGLAS"];
const INIT_FLAGS = {"ASS-001___A2A":true,"ASS-001___ACS DOBFAR":true,"ASS-002___A2A":true,"ASS-002___A5":true,"ASS-003___A5":true,"ASS-004___A5":true,"BAN-001___A5":true,"BAN-001___Alperia":true,"CERT-001___A2A":true,"CERT-001___A5":true,"CERT-001___Acciona":true,"CERT-001___ACEA":true,"CERT-001___ACHILLES":true,"CERT-001___Acquedotto Pugliese":true,"CERT-001___ACS DOBFAR":true,"CERT-001___ADR (Aeoroporti di Roma)":true,"CERT-001___AGCM":true,"CERT-001___AGSM AIM Power":true,"CERT-001___Alperia":true,"CERT-001___Amazon (AWS)- Avetta":true,"CERT-001___Aquila Capital":true,"CERT-001___AST Terni":true,"CERT-001___Autorità di Sistema Portuale del Mare di Sardegna":true,"CERT-001___Autorità di Sistema Portuale del Mare di Sicilia":true,"CERT-001___Autostrade per l'Italia":true,"CERT-001___Barilla":true,"CERT-001___Bosch":true,"CERT-001___Cabot Italia":true,"CERT-001___Cefla":true,"CERT-001___Cerved - Rating Agency":true,"CERT-001___Contourglobal":true,"CERT-001___Creditsafe":true,"CERT-001___CVA Energie":true,"CERT-001___CVL MANTOVA":true,"CERT-001___DATA4":true,"CERT-001___Deval":true,"CERT-001___Dolomiti Energia":true,"CERT-001___Duferco Energia":true,"CERT-001___Edison":true,"CERT-001___EDP Renewables":true,"CERT-001___EF Solare":true,"CERT-001___ENEL":true,"CERT-001___Engie":true,"CERT-001___ENI":true,"CERT-001___E-ON":true,"CERT-001___EP Produzione (Biomasse Italia)":true,"CERT-001___ERG":true,"CERT-001___FAREVA":true,"CERT-001___FER Servizi":true,"CERT-001___Ferrari":true,"CERT-001___Ferrari2":true,"CERT-001___Ferrero":true,"CERT-001___Fincantieri":true,"CERT-001___Galileo":true,"CERT-001___Google":true,"CERT-001___Gori":true,"CERT-001___Governo di Malta":true,"CERT-001___GreenYellow":true,"CERT-001___Heidelberg/Italcementi":true,"CERT-001___Hera":true,"CERT-001___Hyphen":true,"CERT-001___Iberdrola":true,"CERT-001___INFN - ISTITUTO NAZIONALE DI FISICA NUCLEARE":true,"CERT-001___INGFERRARI":true,"CERT-001___Iren":true,"CERT-001___Italgas":true,"CERT-001___Italgas3":true,"CERT-001___Kimberly Clark":true,"CERT-001___Lucchini RS":true,"CERT-001___Max Mara Fashion Group":true,"CERT-001___Mediaset":true,"CERT-001___METINVEST":true,"CERT-001___Microsoft":true,"CERT-001___MM (Metropolitane Milanesi)":true,"CERT-001___Nadara":true,"CERT-001___Neoen":true,"CERT-001___Open-es  Altenia":true,"CERT-001___ORION CARBONS":true,"CERT-001___Parmalat":true,"CERT-001___Philipps Morris":true,"CERT-001___Piaggio":true,"CERT-001___Pilkington (NSG)":true,"CERT-001___Pizzarotti":true,"CERT-001___Pratiche.it":true,"CERT-001___Proger - Alantra Solar":true,"CERT-001___RAM":true,"CERT-001___Renantis_NADARA":true,"CERT-001___Repsol":true,"CERT-001___RFI":true,"CERT-001___Roma Express":true,"CERT-001___ROQUETTE":true,"CERT-001___Saint-Gobain":true,"CERT-001___Saipem":true,"CERT-001___Sapio":true,"CERT-001___Saras":true,"CERT-001___Sasol":true,"CERT-001___Sasol4":true,"CERT-001___Sigef (biomasse)":true,"CERT-001___SIOT - Società  Italiana per l'Oleodotto Transalpino S.p.A.":true,"CERT-001___SNAM":true,"CERT-001___Società per azioni Esercizi Aeroportuali - SEA":true,"CERT-001___Sofidel":true,"CERT-001___Sogin":true,"CERT-001___Solvay-Syensqo":true,"CERT-001___Sorgenia":true,"CERT-001___Statkraft":true,"CERT-001___TERNA_gare":true,"CERT-001___TERNA":true,"CERT-001___TIM Gruppo":true,"CERT-001___Tirreno Power":true,"CERT-001___TRILLIUM":true,"CERT-001___Vaticano":true,"CERT-001___Verallia":true,"CERT-001___Verbund":true,"CERT-001___WARTSILA":true,"CERT-001___Wepa":true,"CERT-001___X-Elio":true,"CERT-001___Zucchetti/TRINSEO/ALTUGLAS":true,"CERT-002___A2A":true,"CERT-002___A5":true,"CERT-002___Acciona":true,"CERT-002___ACEA":true,"CERT-002___ACHILLES":true,"CERT-002___Acquedotto Pugliese":true,"CERT-002___ACS DOBFAR":true,"CERT-002___ADR (Aeoroporti di Roma)":true,"CERT-002___AGCM":true,"CERT-002___AGSM AIM Power":true,"CERT-002___Alperia":true,"CERT-002___Amazon (AWS)- Avetta":true,"CERT-002___Aquila Capital":true,"CERT-002___AST Terni":true,"CERT-002___Autorità di Sistema Portuale del Mare di Sardegna":true,"CERT-002___Autorità di Sistema Portuale del Mare di Sicilia":true,"CERT-002___Autostrade per l'Italia":true,"CERT-002___Barilla":true,"CERT-002___Bosch":true,"CERT-002___Cabot Italia":true,"CERT-002___Cefla":true,"CERT-002___Cerved - Rating Agency":true,"CERT-002___Contourglobal":true,"CERT-002___Creditsafe":true,"CERT-002___CVA Energie":true,"CERT-002___CVL MANTOVA":true,"CERT-002___DATA4":true,"CERT-002___Deval":true,"CERT-002___Dolomiti Energia":true,"CERT-002___Duferco Energia":true,"CERT-002___Edison":true,"CERT-002___EDP Renewables":true,"CERT-002___EF Solare":true,"CERT-002___ENEL":true,"CERT-002___Engie":true,"CERT-002___ENI":true,"CERT-002___E-ON":true,"CERT-002___EP Produzione (Biomasse Italia)":true,"CERT-002___ERG":true,"CERT-002___FAREVA":true,"CERT-002___FER Servizi":true,"CERT-002___Ferrari":true,"CERT-002___Ferrari2":true,"CERT-002___Ferrero":true,"CERT-002___Fincantieri":true,"CERT-002___Galileo":true,"CERT-002___Google":true,"CERT-002___Gori":true,"CERT-002___Governo di Malta":true,"CERT-002___GreenYellow":true,"CERT-002___Heidelberg/Italcementi":true,"CERT-002___Hera":true,"CERT-002___Hyphen":true,"CERT-002___Iberdrola":true,"CERT-002___INFN - ISTITUTO NAZIONALE DI FISICA NUCLEARE":true,"CERT-002___INGFERRARI":true,"CERT-002___Iren":true,"CERT-002___Italgas":true,"CERT-002___Italgas3":true,"CERT-002___Kimberly Clark":true,"CERT-002___Lucchini RS":true,"CERT-002___Max Mara Fashion Group":true,"CERT-002___Mediaset":true,"CERT-002___METINVEST":true,"CERT-002___Microsoft":true,"CERT-002___MM (Metropolitane Milanesi)":true,"CERT-002___Nadara":true,"CERT-002___Neoen":true,"CERT-002___Open-es  Altenia":true,"CERT-002___ORION CARBONS":true,"CERT-002___Parmalat":true,"CERT-002___Philipps Morris":true,"CERT-002___Piaggio":true,"CERT-002___Pilkington (NSG)":true,"CERT-002___Pizzarotti":true,"CERT-002___Pratiche.it":true,"CERT-002___Proger - Alantra Solar":true,"CERT-002___RAM":true,"CERT-002___Renantis_NADARA":true,"CERT-002___Repsol":true,"CERT-002___RFI":true,"CERT-002___Roma Express":true,"CERT-002___ROQUETTE":true,"CERT-002___Saint-Gobain":true,"CERT-002___Saipem":true,"CERT-002___Sapio":true,"CERT-002___Saras":true,"CERT-002___Sasol":true,"CERT-002___Sasol4":true,"CERT-002___Sigef (biomasse)":true,"CERT-002___SIOT - Società  Italiana per l'Oleodotto Transalpino S.p.A.":true,"CERT-002___SNAM":true,"CERT-002___Società per azioni Esercizi Aeroportuali - SEA":true,"CERT-002___Sofidel":true,"CERT-002___Sogin":true,"CERT-002___Solvay-Syensqo":true,"CERT-002___Sorgenia":true,"CERT-002___Statkraft":true,"CERT-002___TERNA_gare":true,"CERT-002___TERNA":true,"CERT-002___TIM Gruppo":true,"CERT-002___Tirreno Power":true,"CERT-002___TRILLIUM":true,"CERT-002___Vaticano":true,"CERT-002___Verallia":true,"CERT-002___Verbund":true,"CERT-002___WARTSILA":true,"CERT-002___Wepa":true,"CERT-002___X-Elio":true,"CERT-002___Zucchetti/TRINSEO/ALTUGLAS":true,"CERT-003___A2A":true,"CERT-003___A5":true,"CERT-003___Acciona":true,"CERT-003___ACEA":true,"CERT-003___ACHILLES":true,"CERT-003___Acquedotto Pugliese":true,"CERT-003___ACS DOBFAR":true,"CERT-003___ADR (Aeoroporti di Roma)":true,"CERT-003___AGCM":true,"CERT-003___AGSM AIM Power":true,"CERT-003___Alperia":true,"CERT-003___Amazon (AWS)- Avetta":true,"CERT-003___Aquila Capital":true,"CERT-003___AST Terni":true,"CERT-003___Autorità di Sistema Portuale del Mare di Sardegna":true,"CERT-003___Autorità di Sistema Portuale del Mare di Sicilia":true,"CERT-003___Autostrade per l'Italia":true,"CERT-003___Barilla":true,"CERT-003___Bosch":true,"CERT-003___Cabot Italia":true,"CERT-003___Cefla":true,"CERT-003___Cerved - Rating Agency":true,"CERT-003___Contourglobal":true,"CERT-003___Creditsafe":true,"CERT-003___CVA Energie":true,"CERT-003___CVL MANTOVA":true,"CERT-003___DATA4":true,"CERT-003___Deval":true,"CERT-003___Dolomiti Energia":true,"CERT-003___Duferco Energia":true,"CERT-003___Edison":true,"CERT-003___EDP Renewables":true,"CERT-003___EF Solare":true,"CERT-003___ENEL":true,"CERT-003___Engie":true,"CERT-003___ENI":true,"CERT-003___E-ON":true,"CERT-003___EP Produzione (Biomasse Italia)":true,"CERT-003___ERG":true,"CERT-003___FAREVA":true,"CERT-003___FER Servizi":true,"CERT-003___Ferrari":true,"CERT-003___Ferrari2":true,"CERT-003___Ferrero":true,"CERT-003___Fincantieri":true,"CERT-003___Galileo":true,"CERT-003___Google":true,"CERT-003___Gori":true,"CERT-003___Governo di Malta":true,"CERT-003___GreenYellow":true,"CERT-003___Heidelberg/Italcementi":true,"CERT-003___Hera":true,"CERT-003___Hyphen":true,"CERT-003___Iberdrola":true,"CERT-003___INFN - ISTITUTO NAZIONALE DI FISICA NUCLEARE":true,"CERT-003___INGFERRARI":true,"CERT-003___Iren":true,"CERT-003___Italgas":true,"CERT-003___Italgas3":true,"CERT-003___Kimberly Clark":true,"CERT-003___Lucchini RS":true,"CERT-003___Max Mara Fashion Group":true,"CERT-003___Mediaset":true,"CERT-003___METINVEST":true,"CERT-003___Microsoft":true,"CERT-003___MM (Metropolitane Milanesi)":true,"CERT-003___Nadara":true,"CERT-003___Neoen":true,"CERT-003___Open-es  Altenia":true,"CERT-003___ORION CARBONS":true,"CERT-003___Parmalat":true,"CERT-003___Philipps Morris":true,"CERT-003___Piaggio":true,"CERT-003___Pilkington (NSG)":true,"CERT-003___Pizzarotti":true,"CERT-003___Pratiche.it":true,"CERT-003___Proger - Alantra Solar":true,"CERT-003___RAM":true,"CERT-003___Renantis_NADARA":true,"CERT-003___Repsol":true,"CERT-003___RFI":true,"CERT-003___Roma Express":true,"CERT-003___ROQUETTE":true,"CERT-003___Saint-Gobain":true,"CERT-003___Saipem":true,"CERT-003___Sapio":true,"CERT-003___Saras":true,"CERT-003___Sasol":true,"CERT-003___Sasol4":true,"CERT-003___Sigef (biomasse)":true,"CERT-003___SIOT - Società  Italiana per l'Oleodotto Transalpino S.p.A.":true,"CERT-003___SNAM":true,"CERT-003___Società per azioni Esercizi Aeroportuali - SEA":true,"CERT-003___Sofidel":true,"CERT-003___Sogin":true,"CERT-003___Solvay-Syensqo":true,"CERT-003___Sorgenia":true,"CERT-003___Statkraft":true,"CERT-003___TERNA_gare":true,"CERT-003___TERNA":true,"CERT-003___TIM Gruppo":true,"CERT-003___Tirreno Power":true,"CERT-003___TRILLIUM":true,"CERT-003___Vaticano":true,"CERT-003___Verallia":true,"CERT-003___Verbund":true,"CERT-003___WARTSILA":true,"CERT-003___Wepa":true,"CERT-003___X-Elio":true,"CERT-003___Zucchetti/TRINSEO/ALTUGLAS":true,"CERT-004___A2A":true,"CERT-004___Acquedotto Pugliese":true,"CERT-004___ADR (Aeoroporti di Roma)":true,"DOC-002___A2A":true,"DOC-002___A5":true,"DOC-002___Acciona":true,"DOC-002___ACEA":true,"DOC-002___ACHILLES":true,"DOC-002___Acquedotto Pugliese":true,"DOC-002___ACS DOBFAR":true,"DOC-002___ADR (Aeoroporti di Roma)":true,"DOC-002___AGCM":true,"DOC-002___AGSM AIM Power":true,"DOC-002___Alperia":true,"DOC-002___Amazon (AWS)- Avetta":true,"DOC-002___Aquila Capital":true,"DOC-002___AST Terni":true,"DOC-002___Autorità di Sistema Portuale del Mare di Sardegna":true,"DOC-002___Autorità di Sistema Portuale del Mare di Sicilia":true,"DOC-002___Autostrade per l'Italia":true,"DOC-002___Barilla":true,"DOC-002___Bosch":true,"DOC-002___Cabot Italia":true,"DOC-002___Cefla":true,"DOC-002___Cerved - Rating Agency":true,"DOC-002___Contourglobal":true,"DOC-002___Creditsafe":true,"DOC-002___CVA Energie":true,"DOC-002___CVL MANTOVA":true,"DOC-002___DATA4":true,"DOC-002___Deval":true,"DOC-002___Dolomiti Energia":true,"DOC-002___Duferco Energia":true,"DOC-002___Edison":true,"DOC-002___EDP Renewables":true,"DOC-002___EF Solare":true,"DOC-002___ENEL":true,"DOC-002___Engie":true,"DOC-002___ENI":true,"DOC-002___E-ON":true,"DOC-002___EP Produzione (Biomasse Italia)":true,"DOC-002___ERG":true,"DOC-002___FAREVA":true,"DOC-002___FER Servizi":true,"DOC-002___Ferrari":true,"DOC-002___Ferrari2":true,"DOC-002___Ferrero":true,"DOC-002___Fincantieri":true,"DOC-002___Galileo":true,"DOC-002___Google":true,"DOC-002___Gori":true,"DOC-002___Governo di Malta":true,"DOC-002___GreenYellow":true,"DOC-002___Heidelberg/Italcementi":true,"DOC-002___Hera":true,"DOC-002___Hyphen":true,"DOC-002___Iberdrola":true,"DOC-002___INFN - ISTITUTO NAZIONALE DI FISICA NUCLEARE":true,"DOC-002___INGFERRARI":true,"DOC-002___Iren":true,"DOC-002___Italgas":true,"DOC-002___Italgas3":true,"DOC-002___Kimberly Clark":true,"DOC-002___Lucchini RS":true,"DOC-002___Max Mara Fashion Group":true,"DOC-002___Mediaset":true,"DOC-002___METINVEST":true,"DOC-002___Microsoft":true,"DOC-002___MM (Metropolitane Milanesi)":true,"DOC-002___Nadara":true,"DOC-002___Neoen":true,"DOC-002___Open-es  Altenia":true,"DOC-002___ORION CARBONS":true,"DOC-002___Parmalat":true,"DOC-002___Philipps Morris":true,"DOC-002___Piaggio":true,"DOC-002___Pilkington (NSG)":true,"DOC-002___Pizzarotti":true,"DOC-002___Pratiche.it":true,"DOC-002___Proger - Alantra Solar":true,"DOC-002___RAM":true,"DOC-002___Renantis_NADARA":true,"DOC-002___Repsol":true,"DOC-002___RFI":true,"DOC-002___Roma Express":true,"DOC-002___ROQUETTE":true,"DOC-002___Saint-Gobain":true,"DOC-002___Saipem":true,"DOC-002___Sapio":true,"DOC-002___Saras":true,"DOC-002___Sasol":true,"DOC-002___Sasol4":true,"DOC-002___Sigef (biomasse)":true,"DOC-002___SIOT - Società  Italiana per l'Oleodotto Transalpino S.p.A.":true,"DOC-002___SNAM":true,"DOC-002___Società per azioni Esercizi Aeroportuali - SEA":true,"DOC-002___Sofidel":true,"DOC-002___Sogin":true,"DOC-002___Solvay-Syensqo":true,"DOC-002___Sorgenia":true,"DOC-002___Statkraft":true,"DOC-002___TERNA_gare":true,"DOC-002___TERNA":true,"DOC-002___TIM Gruppo":true,"DOC-002___Tirreno Power":true,"DOC-002___TRILLIUM":true,"DOC-002___Vaticano":true,"DOC-002___Verallia":true,"DOC-002___Verbund":true,"DOC-002___WARTSILA":true,"DOC-002___Wepa":true,"DOC-002___X-Elio":true,"DOC-002___Zucchetti/TRINSEO/ALTUGLAS":true,"DOC-004___A2A":true,"DOC-004___A5":true,"DOC-004___Acciona":true,"DOC-004___ACEA":true,"DOC-004___ACHILLES":true,"DOC-004___Acquedotto Pugliese":true,"DOC-004___ACS DOBFAR":true,"DOC-004___ADR (Aeoroporti di Roma)":true,"DOC-004___AGCM":true,"DOC-004___AGSM AIM Power":true,"DOC-004___Alperia":true,"DOC-004___Amazon (AWS)- Avetta":true,"DOC-004___Aquila Capital":true,"DOC-004___AST Terni":true,"DOC-004___Autorità di Sistema Portuale del Mare di Sardegna":true,"DOC-004___Autorità di Sistema Portuale del Mare di Sicilia":true,"DOC-004___Autostrade per l'Italia":true,"DOC-004___Barilla":true,"DOC-004___Bosch":true,"DOC-004___Cabot Italia":true,"DOC-004___Cefla":true,"DOC-004___Cerved - Rating Agency":true,"DOC-004___Contourglobal":true,"DOC-004___Creditsafe":true,"DOC-004___CVA Energie":true,"DOC-004___CVL MANTOVA":true,"DOC-004___DATA4":true,"DOC-004___Deval":true,"DOC-004___Dolomiti Energia":true,"DOC-004___Duferco Energia":true,"DOC-004___Edison":true,"DOC-004___EDP Renewables":true,"DOC-004___EF Solare":true,"DOC-004___ENEL":true,"DOC-004___Engie":true,"DOC-004___ENI":true,"DOC-004___E-ON":true,"DOC-004___EP Produzione (Biomasse Italia)":true,"DOC-004___ERG":true,"DOC-004___FAREVA":true,"DOC-004___FER Servizi":true,"DOC-004___Ferrari":true,"DOC-004___Ferrari2":true,"DOC-004___Ferrero":true,"DOC-004___Fincantieri":true,"DOC-004___Galileo":true,"DOC-004___Google":true,"DOC-004___Gori":true,"DOC-004___Governo di Malta":true,"DOC-004___GreenYellow":true,"DOC-004___Heidelberg/Italcementi":true,"DOC-004___Hera":true,"DOC-004___Hyphen":true,"DOC-004___Iberdrola":true,"DOC-004___INFN - ISTITUTO NAZIONALE DI FISICA NUCLEARE":true,"DOC-004___INGFERRARI":true,"DOC-004___Iren":true,"DOC-004___Italgas":true,"DOC-004___Italgas3":true,"DOC-004___Kimberly Clark":true,"DOC-004___Lucchini RS":true,"DOC-004___Max Mara Fashion Group":true,"DOC-004___Mediaset":true,"DOC-004___METINVEST":true,"DOC-004___Microsoft":true,"DOC-004___MM (Metropolitane Milanesi)":true,"DOC-004___Nadara":true,"DOC-004___Neoen":true,"DOC-004___Open-es  Altenia":true,"DOC-004___ORION CARBONS":true,"DOC-004___Parmalat":true,"DOC-004___Philipps Morris":true,"DOC-004___Piaggio":true,"DOC-004___Pilkington (NSG)":true,"DOC-004___Pizzarotti":true,"DOC-004___Pratiche.it":true,"DOC-004___Proger - Alantra Solar":true,"DOC-004___RAM":true,"DOC-004___Renantis_NADARA":true,"DOC-004___Repsol":true,"DOC-004___RFI":true,"DOC-004___Roma Express":true,"DOC-004___ROQUETTE":true,"DOC-004___Saint-Gobain":true,"DOC-004___Saipem":true,"DOC-004___Sapio":true,"DOC-004___Saras":true,"DOC-004___Sasol":true,"DOC-004___Sasol4":true,"DOC-004___Sigef (biomasse)":true,"DOC-004___SIOT - Società  Italiana per l'Oleodotto Transalpino S.p.A.":true,"DOC-004___SNAM":true,"DOC-004___Società per azioni Esercizi Aeroportuali - SEA":true,"DOC-004___Sofidel":true,"DOC-004___Sogin":true,"DOC-004___Solvay-Syensqo":true,"DOC-004___Sorgenia":true,"DOC-004___Statkraft":true,"DOC-004___TERNA_gare":true,"DOC-004___TERNA":true,"DOC-004___TIM Gruppo":true,"DOC-004___Tirreno Power":true,"DOC-004___TRILLIUM":true,"DOC-004___Vaticano":true,"DOC-004___Verallia":true,"DOC-004___Verbund":true,"DOC-004___WARTSILA":true,"DOC-004___Wepa":true,"DOC-004___X-Elio":true,"DOC-004___Zucchetti/TRINSEO/ALTUGLAS":true,"DOC-006___A2A":true,"DOC-006___A5":true,"DOC-006___Acciona":true,"DOC-006___ACEA":true,"DOC-006___ACHILLES":true,"DOC-006___Acquedotto Pugliese":true,"DOC-006___ACS DOBFAR":true,"DOC-006___ADR (Aeoroporti di Roma)":true,"DOC-006___AGCM":true,"DOC-006___AGSM AIM Power":true,"DOC-006___Alperia":true,"DOC-006___Amazon (AWS)- Avetta":true,"DOC-006___Aquila Capital":true,"DOC-006___AST Terni":true,"DOC-006___Autorità di Sistema Portuale del Mare di Sardegna":true,"DOC-006___Autorità di Sistema Portuale del Mare di Sicilia":true,"DOC-006___Autostrade per l'Italia":true,"DOC-006___Barilla":true,"DOC-006___Bosch":true,"DOC-006___Cabot Italia":true,"DOC-006___Cefla":true,"DOC-006___Cerved - Rating Agency":true,"DOC-006___Contourglobal":true,"DOC-006___Creditsafe":true,"DOC-006___CVA Energie":true,"DOC-006___CVL MANTOVA":true,"DOC-006___DATA4":true,"DOC-006___Deval":true,"DOC-006___Dolomiti Energia":true,"DOC-006___Duferco Energia":true,"DOC-006___Edison":true,"DOC-006___EDP Renewables":true,"DOC-006___EF Solare":true,"DOC-006___ENEL":true,"DOC-006___Engie":true,"DOC-006___ENI":true,"DOC-006___E-ON":true,"DOC-006___EP Produzione (Biomasse Italia)":true,"DOC-006___ERG":true,"DOC-006___FAREVA":true,"DOC-006___FER Servizi":true,"DOC-006___Ferrari":true,"DOC-006___Ferrari2":true,"DOC-006___Ferrero":true,"DOC-006___Fincantieri":true,"DOC-006___Galileo":true,"DOC-006___Google":true,"DOC-006___Gori":true,"DOC-006___Governo di Malta":true,"DOC-006___GreenYellow":true,"DOC-006___Heidelberg/Italcementi":true,"DOC-006___Hera":true,"DOC-006___Hyphen":true,"DOC-006___Iberdrola":true,"DOC-006___INFN - ISTITUTO NAZIONALE DI FISICA NUCLEARE":true,"DOC-006___INGFERRARI":true,"DOC-006___Iren":true,"DOC-006___Italgas":true,"DOC-006___Italgas3":true,"DOC-006___Kimberly Clark":true,"DOC-006___Lucchini RS":true,"DOC-006___Max Mara Fashion Group":true,"DOC-006___Mediaset":true,"DOC-006___METINVEST":true,"DOC-006___Microsoft":true,"DOC-006___MM (Metropolitane Milanesi)":true,"DOC-006___Nadara":true,"DOC-006___Neoen":true,"DOC-006___Open-es  Altenia":true,"DOC-006___ORION CARBONS":true,"DOC-006___Parmalat":true,"DOC-006___Philipps Morris":true,"DOC-006___Piaggio":true,"DOC-006___Pilkington (NSG)":true,"DOC-006___Pizzarotti":true,"DOC-006___Pratiche.it":true,"DOC-006___Proger - Alantra Solar":true,"DOC-006___RAM":true,"DOC-006___Renantis_NADARA":true,"DOC-006___Repsol":true,"DOC-006___RFI":true,"DOC-006___Roma Express":true,"DOC-006___ROQUETTE":true,"DOC-006___Saint-Gobain":true,"DOC-006___Saipem":true,"DOC-006___Sapio":true,"DOC-006___Saras":true,"DOC-006___Sasol":true,"DOC-006___Sasol4":true,"DOC-006___Sigef (biomasse)":true,"DOC-006___SIOT - Società  Italiana per l'Oleodotto Transalpino S.p.A.":true,"DOC-006___SNAM":true,"DOC-006___Società per azioni Esercizi Aeroportuali - SEA":true,"DOC-006___Sofidel":true,"DOC-006___Sogin":true,"DOC-006___Solvay-Syensqo":true,"DOC-006___Sorgenia":true,"DOC-006___Statkraft":true,"DOC-006___TERNA_gare":true,"DOC-006___TERNA":true,"DOC-006___TIM Gruppo":true,"DOC-006___Tirreno Power":true,"DOC-006___TRILLIUM":true,"DOC-006___Vaticano":true,"DOC-006___Verallia":true,"DOC-006___Verbund":true,"DOC-006___WARTSILA":true,"DOC-006___Wepa":true,"DOC-006___X-Elio":true,"DOC-006___Zucchetti/TRINSEO/ALTUGLAS":true,"DOC-007___A2A":true,"DOC-007___A5":true,"DOC-007___Acciona":true,"DOC-007___ACEA":true,"DOC-007___ACHILLES":true,"DOC-007___Acquedotto Pugliese":true,"DOC-007___ACS DOBFAR":true,"DOC-007___ADR (Aeoroporti di Roma)":true,"DOC-007___AGCM":true,"DOC-007___AGSM AIM Power":true,"DOC-007___Alperia":true,"DOC-007___Amazon (AWS)- Avetta":true,"DOC-007___Aquila Capital":true,"DOC-007___AST Terni":true,"DOC-007___Autorità di Sistema Portuale del Mare di Sardegna":true,"DOC-007___Autorità di Sistema Portuale del Mare di Sicilia":true,"DOC-007___Autostrade per l'Italia":true,"DOC-007___Barilla":true,"DOC-007___Bosch":true,"DOC-007___Cabot Italia":true,"DOC-007___Cefla":true,"DOC-007___Cerved - Rating Agency":true,"DOC-007___Contourglobal":true,"DOC-007___Creditsafe":true,"DOC-007___CVA Energie":true,"DOC-007___CVL MANTOVA":true,"DOC-007___DATA4":true,"DOC-007___Deval":true,"DOC-007___Dolomiti Energia":true,"DOC-007___Duferco Energia":true,"DOC-007___Edison":true,"DOC-007___EDP Renewables":true,"DOC-007___EF Solare":true,"DOC-007___ENEL":true,"DOC-007___Engie":true,"DOC-007___ENI":true,"DOC-007___E-ON":true,"DOC-007___EP Produzione (Biomasse Italia)":true,"DOC-007___ERG":true,"DOC-007___FAREVA":true,"DOC-007___FER Servizi":true,"DOC-007___Ferrari":true,"DOC-007___Ferrari2":true,"DOC-007___Ferrero":true,"DOC-007___Fincantieri":true,"DOC-007___Galileo":true,"DOC-007___Google":true,"DOC-007___Gori":true,"DOC-007___Governo di Malta":true,"DOC-007___GreenYellow":true,"DOC-007___Heidelberg/Italcementi":true,"DOC-007___Hera":true,"DOC-007___Hyphen":true,"DOC-007___Iberdrola":true,"DOC-007___INFN - ISTITUTO NAZIONALE DI FISICA NUCLEARE":true,"DOC-007___INGFERRARI":true,"DOC-007___Iren":true,"DOC-007___Italgas":true,"DOC-007___Italgas3":true,"DOC-007___Kimberly Clark":true,"DOC-007___Lucchini RS":true,"DOC-007___Max Mara Fashion Group":true,"DOC-007___Mediaset":true,"DOC-007___METINVEST":true,"DOC-007___Microsoft":true,"DOC-007___MM (Metropolitane Milanesi)":true,"DOC-007___Nadara":true,"DOC-007___Neoen":true,"DOC-007___Open-es  Altenia":true,"DOC-007___ORION CARBONS":true,"DOC-007___Parmalat":true,"DOC-007___Philipps Morris":true,"DOC-007___Piaggio":true,"DOC-007___Pilkington (NSG)":true,"DOC-007___Pizzarotti":true,"DOC-007___Pratiche.it":true,"DOC-007___Proger - Alantra Solar":true,"DOC-007___RAM":true,"DOC-007___Renantis_NADARA":true,"DOC-007___Repsol":true,"DOC-007___RFI":true,"DOC-007___Roma Express":true,"DOC-007___ROQUETTE":true,"DOC-007___Saint-Gobain":true,"DOC-007___Saipem":true,"DOC-007___Sapio":true,"DOC-007___Saras":true,"DOC-007___Sasol":true,"DOC-007___Sasol4":true,"DOC-007___Sigef (biomasse)":true,"DOC-007___SIOT - Società  Italiana per l'Oleodotto Transalpino S.p.A.":true,"DOC-007___SNAM":true,"DOC-007___Società per azioni Esercizi Aeroportuali - SEA":true,"DOC-007___Sofidel":true,"DOC-007___Sogin":true,"DOC-007___Solvay-Syensqo":true,"DOC-007___Sorgenia":true,"DOC-007___Statkraft":true,"DOC-007___TERNA_gare":true,"DOC-007___TERNA":true,"DOC-007___TIM Gruppo":true,"DOC-007___Tirreno Power":true,"DOC-007___TRILLIUM":true,"DOC-007___Vaticano":true,"DOC-007___Verallia":true,"DOC-007___Verbund":true,"DOC-007___WARTSILA":true,"DOC-007___Wepa":true,"DOC-007___X-Elio":true,"DOC-007___Zucchetti/TRINSEO/ALTUGLAS":true,"DOC-008___A2A":true,"DOC-008___A5":true,"DOC-008___Acciona":true,"DOC-008___ACEA":true,"DOC-008___ACHILLES":true,"DOC-008___Acquedotto Pugliese":true,"DOC-008___ACS DOBFAR":true,"DOC-008___ADR (Aeoroporti di Roma)":true,"DOC-008___AGCM":true,"DOC-008___AGSM AIM Power":true,"DOC-008___Alperia":true,"DOC-008___Amazon (AWS)- Avetta":true,"DOC-008___Aquila Capital":true,"DOC-008___AST Terni":true,"DOC-008___Autorità di Sistema Portuale del Mare di Sardegna":true,"DOC-008___Autorità di Sistema Portuale del Mare di Sicilia":true,"DOC-008___Autostrade per l'Italia":true,"DOC-008___Barilla":true,"DOC-008___Bosch":true,"DOC-008___Cabot Italia":true,"DOC-008___Cefla":true,"DOC-008___Cerved - Rating Agency":true,"DOC-008___Contourglobal":true,"DOC-008___Creditsafe":true,"DOC-008___CVA Energie":true,"DOC-008___CVL MANTOVA":true,"DOC-008___DATA4":true,"DOC-008___Deval":true,"DOC-008___Dolomiti Energia":true,"DOC-008___Duferco Energia":true,"DOC-008___Edison":true,"DOC-008___EDP Renewables":true,"DOC-008___EF Solare":true,"DOC-008___ENEL":true,"DOC-008___Engie":true,"DOC-008___ENI":true,"DOC-008___E-ON":true,"DOC-008___EP Produzione (Biomasse Italia)":true,"DOC-008___ERG":true,"DOC-008___FAREVA":true,"DOC-008___FER Servizi":true,"DOC-008___Ferrari":true,"DOC-008___Ferrari2":true,"DOC-008___Ferrero":true,"DOC-008___Fincantieri":true,"DOC-008___Galileo":true,"DOC-008___Google":true,"DOC-008___Gori":true,"DOC-008___Governo di Malta":true,"DOC-008___GreenYellow":true,"DOC-008___Heidelberg/Italcementi":true,"DOC-008___Hera":true,"DOC-008___Hyphen":true,"DOC-008___Iberdrola":true,"DOC-008___INFN - ISTITUTO NAZIONALE DI FISICA NUCLEARE":true,"DOC-008___INGFERRARI":true,"DOC-008___Iren":true,"DOC-008___Italgas":true,"DOC-008___Italgas3":true,"DOC-008___Kimberly Clark":true,"DOC-008___Lucchini RS":true,"DOC-008___Max Mara Fashion Group":true,"DOC-008___Mediaset":true,"DOC-008___METINVEST":true,"DOC-008___Microsoft":true,"DOC-008___MM (Metropolitane Milanesi)":true,"DOC-008___Nadara":true,"DOC-008___Neoen":true,"DOC-008___Open-es  Altenia":true,"DOC-008___ORION CARBONS":true,"DOC-008___Parmalat":true,"DOC-008___Philipps Morris":true,"DOC-008___Piaggio":true,"DOC-008___Pilkington (NSG)":true,"DOC-008___Pizzarotti":true,"DOC-008___Pratiche.it":true,"DOC-008___Proger - Alantra Solar":true,"DOC-008___RAM":true,"DOC-008___Renantis_NADARA":true,"DOC-008___Repsol":true,"DOC-008___RFI":true,"DOC-008___Roma Express":true,"DOC-008___ROQUETTE":true,"DOC-008___Saint-Gobain":true,"DOC-008___Saipem":true,"DOC-008___Sapio":true,"DOC-008___Saras":true,"DOC-008___Sasol":true,"DOC-008___Sasol4":true,"DOC-008___Sigef (biomasse)":true,"DOC-008___SIOT - Società  Italiana per l'Oleodotto Transalpino S.p.A.":true,"DOC-008___SNAM":true,"DOC-008___Società per azioni Esercizi Aeroportuali - SEA":true,"DOC-008___Sofidel":true,"DOC-008___Sogin":true,"DOC-008___Solvay-Syensqo":true,"DOC-008___Sorgenia":true,"DOC-008___Statkraft":true,"DOC-008___TERNA_gare":true,"DOC-008___TERNA":true,"DOC-008___TIM Gruppo":true,"DOC-008___Tirreno Power":true,"DOC-008___TRILLIUM":true,"DOC-008___Vaticano":true,"DOC-008___Verallia":true,"DOC-008___Verbund":true,"DOC-008___WARTSILA":true,"DOC-008___Wepa":true,"DOC-008___X-Elio":true,"DOC-008___Zucchetti/TRINSEO/ALTUGLAS":true,"DOC-009___A2A":true,"DOC-009___ACEA":true,"DOC-009___Acquedotto Pugliese":true,"DOC-010___A2A":true,"DOC-010___ACEA":true,"DOC-010___Acquedotto Pugliese":true,"DOC-013___ACEA":true,"DOC-013___Acquedotto Pugliese":true,"DOC-014___ACEA":true,"DOC-014___Acquedotto Pugliese":true,"DOC-017___A2A":true,"DOC-017___Acquedotto Pugliese":true,"DOC-020___ACEA":true,"DOC-020___Acquedotto Pugliese":true,"PRIV-002___ADR (Aeoroporti di Roma)":true,"SIC-002___A2A":true,"SIC-002___A5":true,"SIC-004___A5":true,"SIC-005___A2A":true,"SIC-005___A5":true,"SIC-005___ADR (Aeoroporti di Roma)":true,"SIC-007___A2A":true,"TEC-001___A2A":true,"TEC-001___ACEA":true,"TEC-001___Acquedotto Pugliese":true,"TEC-004___ADR (Aeoroporti di Roma)":true,"TEC-005___A5":true,"TEC-005___ACEA":true,"TEC-005___Alperia":true,"FIS-001___A2A":true,"FIS-003___A5":true,"FIS-003___Acciona":true,"FIS-003___ACEA":true,"FIS-003___ACHILLES":true,"FIS-003___Acquedotto Pugliese":true,"FIS-003___ACS DOBFAR":true,"FIS-003___ADR (Aeoroporti di Roma)":true,"FIS-003___AGCM":true,"FIS-003___AGSM AIM Power":true,"FIS-003___Alperia":true,"FIS-003___Amazon (AWS)- Avetta":true,"FIS-003___Aquila Capital":true,"FIS-003___AST Terni":true,"FIS-003___Autorità di Sistema Portuale del Mare di Sardegna":true,"FIS-003___Autorità di Sistema Portuale del Mare di Sicilia":true,"FIS-003___Autostrade per l'Italia":true,"FIS-003___Barilla":true,"FIS-003___Bosch":true,"FIS-003___Cabot Italia":true,"FIS-003___Cefla":true,"FIS-003___Cerved - Rating Agency":true,"FIS-003___Contourglobal":true,"FIS-003___Creditsafe":true,"FIS-003___CVA Energie":true,"FIS-003___CVL MANTOVA":true,"FIS-003___DATA4":true,"FIS-003___Deval":true,"FIS-003___Dolomiti Energia":true,"FIS-003___Duferco Energia":true,"FIS-003___Edison":true,"FIS-003___EDP Renewables":true,"FIS-003___EF Solare":true,"FIS-003___ENEL":true,"FIS-003___Engie":true,"FIS-003___ENI":true,"FIS-003___E-ON":true,"FIS-003___EP Produzione (Biomasse Italia)":true,"FIS-003___ERG":true,"FIS-003___FAREVA":true,"FIS-003___FER Servizi":true,"FIS-003___Ferrari":true,"FIS-003___Ferrari2":true,"FIS-003___Ferrero":true,"FIS-003___Fincantieri":true,"FIS-003___Galileo":true,"FIS-003___Google":true,"FIS-003___Gori":true,"FIS-003___Governo di Malta":true,"FIS-003___GreenYellow":true,"FIS-003___Heidelberg/Italcementi":true,"FIS-003___Hera":true,"FIS-003___Hyphen":true,"FIS-003___Iberdrola":true,"FIS-003___INFN - ISTITUTO NAZIONALE DI FISICA NUCLEARE":true,"FIS-003___INGFERRARI":true,"FIS-003___Iren":true,"FIS-003___Italgas":true,"FIS-003___Italgas3":true,"FIS-003___Kimberly Clark":true,"FIS-003___Lucchini RS":true,"FIS-003___Max Mara Fashion Group":true,"FIS-003___Mediaset":true,"FIS-003___METINVEST":true,"FIS-003___Microsoft":true,"FIS-003___MM (Metropolitane Milanesi)":true,"FIS-003___Nadara":true,"FIS-003___Neoen":true,"FIS-003___Open-es  Altenia":true,"FIS-003___ORION CARBONS":true,"FIS-003___Parmalat":true,"FIS-003___Philipps Morris":true,"FIS-003___Piaggio":true,"FIS-003___Pilkington (NSG)":true,"FIS-003___Pizzarotti":true,"FIS-003___Pratiche.it":true,"FIS-003___Proger - Alantra Solar":true,"FIS-003___RAM":true,"FIS-003___Renantis_NADARA":true,"FIS-003___Repsol":true,"FIS-003___RFI":true,"FIS-003___Roma Express":true,"FIS-003___ROQUETTE":true,"FIS-003___Saint-Gobain":true,"FIS-003___Saipem":true,"FIS-003___Sapio":true,"FIS-003___Saras":true,"FIS-003___Sasol":true,"FIS-003___Sasol4":true,"FIS-003___Sigef (biomasse)":true,"FIS-003___SIOT - Società  Italiana per l'Oleodotto Transalpino S.p.A.":true,"FIS-003___SNAM":true,"FIS-003___Società per azioni Esercizi Aeroportuali - SEA":true,"FIS-003___Sofidel":true,"FIS-003___Sogin":true,"FIS-003___Solvay-Syensqo":true,"FIS-003___Sorgenia":true,"FIS-003___Statkraft":true,"FIS-003___TERNA_gare":true,"FIS-003___TERNA":true,"FIS-003___TIM Gruppo":true,"FIS-003___Tirreno Power":true,"FIS-003___TRILLIUM":true,"FIS-003___Vaticano":true,"FIS-003___Verallia":true,"FIS-003___Verbund":true,"FIS-003___WARTSILA":true,"FIS-003___Wepa":true,"FIS-003___X-Elio":true,"FIS-003___Zucchetti/TRINSEO/ALTUGLAS":true,"FIS-004___A5":true,"FIS-004___Acciona":true,"FIS-004___ACEA":true,"FIS-004___ACHILLES":true,"FIS-004___Acquedotto Pugliese":true,"FIS-004___ACS DOBFAR":true,"FIS-004___ADR (Aeoroporti di Roma)":true,"FIS-004___AGCM":true,"FIS-004___AGSM AIM Power":true,"FIS-004___Alperia":true,"FIS-004___Amazon (AWS)- Avetta":true,"FIS-004___Aquila Capital":true,"FIS-004___AST Terni":true,"FIS-004___Autorità di Sistema Portuale del Mare di Sardegna":true,"FIS-004___Autorità di Sistema Portuale del Mare di Sicilia":true,"FIS-004___Autostrade per l'Italia":true,"FIS-004___Barilla":true,"FIS-004___Bosch":true,"FIS-004___Cabot Italia":true,"FIS-004___Cefla":true,"FIS-004___Cerved - Rating Agency":true,"FIS-004___Contourglobal":true,"FIS-004___Creditsafe":true,"FIS-004___CVA Energie":true,"FIS-004___CVL MANTOVA":true,"FIS-004___DATA4":true,"FIS-004___Deval":true,"FIS-004___Dolomiti Energia":true,"FIS-004___Duferco Energia":true,"FIS-004___Edison":true,"FIS-004___EDP Renewables":true,"FIS-004___EF Solare":true,"FIS-004___ENEL":true,"FIS-004___Engie":true,"FIS-004___ENI":true,"FIS-004___E-ON":true,"FIS-004___EP Produzione (Biomasse Italia)":true,"FIS-004___ERG":true,"FIS-004___FAREVA":true,"FIS-004___FER Servizi":true,"FIS-004___Ferrari":true,"FIS-004___Ferrari2":true,"FIS-004___Ferrero":true,"FIS-004___Fincantieri":true,"FIS-004___Galileo":true,"FIS-004___Google":true,"FIS-004___Gori":true,"FIS-004___Governo di Malta":true,"FIS-004___GreenYellow":true,"FIS-004___Heidelberg/Italcementi":true,"FIS-004___Hera":true,"FIS-004___Hyphen":true,"FIS-004___Iberdrola":true,"FIS-004___INFN - ISTITUTO NAZIONALE DI FISICA NUCLEARE":true,"FIS-004___INGFERRARI":true,"FIS-004___Iren":true,"FIS-004___Italgas":true,"FIS-004___Italgas3":true,"FIS-004___Kimberly Clark":true,"FIS-004___Lucchini RS":true,"FIS-004___Max Mara Fashion Group":true,"FIS-004___Mediaset":true,"FIS-004___METINVEST":true,"FIS-004___Microsoft":true,"FIS-004___MM (Metropolitane Milanesi)":true,"FIS-004___Nadara":true,"FIS-004___Neoen":true,"FIS-004___Open-es  Altenia":true,"FIS-004___ORION CARBONS":true,"FIS-004___Parmalat":true,"FIS-004___Philipps Morris":true,"FIS-004___Piaggio":true,"FIS-004___Pilkington (NSG)":true,"FIS-004___Pizzarotti":true,"FIS-004___Pratiche.it":true,"FIS-004___Proger - Alantra Solar":true,"FIS-004___RAM":true,"FIS-004___Renantis_NADARA":true,"FIS-004___Repsol":true,"FIS-004___RFI":true,"FIS-004___Roma Express":true,"FIS-004___ROQUETTE":true,"FIS-004___Saint-Gobain":true,"FIS-004___Saipem":true,"FIS-004___Sapio":true,"FIS-004___Saras":true,"FIS-004___Sasol":true,"FIS-004___Sasol4":true,"FIS-004___Sigef (biomasse)":true,"FIS-004___SIOT - Società  Italiana per l'Oleodotto Transalpino S.p.A.":true,"FIS-004___SNAM":true,"FIS-004___Società per azioni Esercizi Aeroportuali - SEA":true,"FIS-004___Sofidel":true,"FIS-004___Sogin":true,"FIS-004___Solvay-Syensqo":true,"FIS-004___Sorgenia":true,"FIS-004___Statkraft":true,"FIS-004___TERNA_gare":true,"FIS-004___TERNA":true,"FIS-004___TIM Gruppo":true,"FIS-004___Tirreno Power":true,"FIS-004___TRILLIUM":true,"FIS-004___Vaticano":true,"FIS-004___Verallia":true,"FIS-004___Verbund":true,"FIS-004___WARTSILA":true,"FIS-004___Wepa":true,"FIS-004___X-Elio":true,"FIS-004___Zucchetti/TRINSEO/ALTUGLAS":true,"DOC-021___Acquedotto Pugliese":true,"DOC-021___ADR (Aeoroporti di Roma)":true,"DOC-021___Alperia":true,"FIS-004_1___A2A":true,"FIS-004_1___A5":true,"FIS-004_1___ADR (Aeoroporti di Roma)":true,"FIS-004_1___Alperia":true,"CERT-007___A5":true,"CERT-007___ADR (Aeoroporti di Roma)":true,"ASS-007___ACS DOBFAR":true};

// Initial category registry: prefix -> label
const INIT_CATS = {
  "ASS":"Assicurazioni",
  "BAN":"Bancario/Finanziario",
  "CERT":"Certificazioni",
  "DOC":"Legale/Societario",
  "FIS":"Fiscale",
  "PRIV":"Privacy/Compliance",
  "SIC":"Sicurezza",
  "TEC":"Tecnico/Operativo"
};

// ── utils ─────────────────────────────────────────────────────────────────────
function addMonths(s,m){const d=new Date(s);d.setMonth(d.getMonth()+m);return d;}
function diffDays(a,b=new Date()){return Math.round((a-b)/86400000);}
function fmtDate(d){if(!d)return"—";return new Date(d).toLocaleDateString("it-IT");}
function stato(doc){
  if(!doc.dataEmissione||!doc.durata)return null;
  const exp=addMonths(doc.dataEmissione,+doc.durata),days=diffDays(exp);
  if(days<0)return"SCADUTO";if(days<=30)return"URGENTE";if(days<=60)return"ATTENZIONE";if(days<=90)return"IN SCADENZA";return"VALIDO";
}
function statoColor(s){
  if(s==="SCADUTO")    return{bg:"#d32f2f",text:"#fff",flagBg:"#ffcdd2",flagText:"#b71c1c"};
  if(s==="URGENTE")    return{bg:"#e53935",text:"#fff",flagBg:"#ffcdd2",flagText:"#c62828"};
  if(s==="ATTENZIONE") return{bg:"#f9a825",text:"#000",flagBg:"#fff9c4",flagText:"#7a6000"};
  if(s==="IN SCADENZA")return{bg:"#fb8c00",text:"#fff",flagBg:"#ffe0b2",flagText:"#6d3a00"};
  return{bg:"#388e3c",text:"#fff",flagBg:"#c8e6c9",flagText:"#1b5e20"};
}
function flagKey(docId,cli){return`${docId}___${cli}`;}

// Next code for a given prefix
function nextDocCode(docs, prefix){
  const used = docs
    .filter(d => d.id.startsWith(prefix+"-") && /^\d+$/.test(d.id.split("-")[1]))
    .map(d => parseInt(d.id.split("-")[1]));
  const max = used.length ? Math.max(...used) : 0;
  return `${prefix}-${String(max+1).padStart(3,"0")}`;
}

// Sort docs: by category prefix alphabetically, then by numeric part of id
function sortDocs(docs){
  return [...docs].sort((a,b)=>{
    const pa=a.id.split("-")[0], pb=b.id.split("-")[0];
    if(pa!==pb)return pa.localeCompare(pb);
    // extract numeric part (ignore suffixes like _1)
    const na=parseInt(a.id.split("-")[1])||0, nb=parseInt(b.id.split("-")[1])||0;
    if(na!==nb)return na-nb;
    return a.id.localeCompare(b.id);
  });
}

// ── Excel export ──────────────────────────────────────────────────────────────
function exportExcel(docs,clients,flags){
  const wb=XLSX.utils.book_new();
  const sorted=sortDocs(docs);
  const s1=[["Codice","Categoria","Documento","Data Emissione","Durata (mesi)","Data Scadenza","Giorni Rimanenti","Stato"]];
  sorted.forEach(d=>{
    const exp=d.dataEmissione&&d.durata?addMonths(d.dataEmissione,+d.durata):null;
    s1.push([d.id,d.cat,d.nome,d.dataEmissione?fmtDate(d.dataEmissione):"",d.durata||"",exp?fmtDate(exp):"",exp?diffDays(exp):"",stato(d)||""]);
  });
  const ws1=XLSX.utils.aoa_to_sheet(s1);ws1["!cols"]=[10,22,45,16,14,16,14,14].map(w=>({wch:w}));
  XLSX.utils.book_append_sheet(wb,ws1,"Scadenzario");
  const s2=[["Codice","Documento",...clients]];
  sorted.forEach(d=>s2.push([d.id,d.nome,...clients.map(c=>flags[flagKey(d.id,c)]?"✔":"")]));
  s2.push(["","Totale presenti",...clients.map(c=>sorted.filter(d=>flags[flagKey(d.id,c)]).length)]);
  const ws2=XLSX.utils.aoa_to_sheet(s2);ws2["!cols"]=[11,45,...clients.map(()=>({wch:18}))].map((w,i)=>typeof w==="number"?{wch:w}:w);
  XLSX.utils.book_append_sheet(wb,ws2,"Matrice Documenti-Portali");
  XLSX.writeFile(wb,"Monitoraggio_Scadenze.xlsx");
}

// ── Shared styles ─────────────────────────────────────────────────────────────
// Light grey-scale theme
const C = {
  pageBg:   "#f0f0f0",
  headerBg: "#2c2c2c",
  headerText:"#ffffff",
  cardBg:   "#ffffff",
  cardBorder:"#d0d0d0",
  rowEven:  "#f8f8f8",
  rowOdd:   "#ffffff",
  tabActive: "#2c2c2c",
  tabInactive:"#888",
  inputBg:  "#ffffff",
  inputBorder:"#bbb",
  inputText:"#222",
  label:    "#555",
  mutedText:"#888",
  matrixCellBg:"#3a3a3a",  // dark grey for flag area
  matrixHeaderBg:"#2c2c2c",
  stickyBg: "#ffffff",
};

const IS={width:"100%",background:C.inputBg,border:`1px solid ${C.inputBorder}`,borderRadius:7,padding:"9px 13px",color:C.inputText,fontSize:13,outline:"none",fontFamily:"'IBM Plex Mono',monospace",boxSizing:"border-box"};

function Field({label,children,hint}){
  return(<div style={{marginBottom:14}}>
    <label style={{display:"block",fontSize:10,fontWeight:700,letterSpacing:"0.08em",textTransform:"uppercase",color:C.label,marginBottom:5}}>{label}</label>
    {children}
    {hint&&<div style={{fontSize:10,color:C.mutedText,marginTop:3,fontStyle:"italic"}}>{hint}</div>}
  </div>);
}

function Btn({onClick,children,variant="primary",small,full,disabled}){
  const V={
    primary:{background:"#2c2c2c",color:"#fff"},
    success:{background:"#2e7d32",color:"#fff"},
    danger: {background:"#c62828",color:"#fff"},
    ghost:  {background:"transparent",color:"#555",border:"1px solid #ccc"},
    accent: {background:"#455a64",color:"#fff"},
  };
  return(<button onClick={onClick} disabled={disabled} style={{
    ...V[variant],border:V[variant].border||"none",borderRadius:7,cursor:disabled?"not-allowed":"pointer",
    fontWeight:700,fontSize:small?11:13,padding:small?"5px 12px":"10px 20px",
    letterSpacing:"0.03em",transition:"opacity 0.15s",fontFamily:"'IBM Plex Mono',monospace",
    width:full?"100%":"auto",opacity:disabled?0.4:1,
  }} onMouseOver={e=>{if(!disabled)e.currentTarget.style.opacity=0.8;}} onMouseOut={e=>{e.currentTarget.style.opacity=disabled?0.4:1;}}>{children}</button>);
}

function Overlay({children}){
  return(<div style={{position:"fixed",inset:0,background:"rgba(0,0,0,0.55)",display:"flex",alignItems:"center",justifyContent:"center",zIndex:2000,backdropFilter:"blur(4px)"}}>{children}</div>);
}

function ModalBox({title,onClose,children,width=500}){
  return(<Overlay><div style={{background:"#fff",border:"1px solid #ccc",borderRadius:14,padding:28,width,maxWidth:"96%",boxShadow:"0 16px 48px rgba(0,0,0,0.18)",maxHeight:"92vh",overflowY:"auto"}}>
    <div style={{display:"flex",justifyContent:"space-between",alignItems:"center",marginBottom:20}}>
      <span style={{fontFamily:"'DM Serif Display',serif",fontSize:18,color:"#222"}}>{title}</span>
      <button onClick={onClose} style={{background:"none",border:"none",color:"#888",fontSize:22,cursor:"pointer",lineHeight:1}}>✕</button>
    </div>
    {children}
  </div></Overlay>);
}

function ConfirmDialog({msg,onConfirm,onCancel}){
  return(<Overlay><div style={{background:"#fff",border:"1px solid #e57373",borderRadius:13,padding:32,maxWidth:420,width:"92%",boxShadow:"0 12px 40px rgba(0,0,0,0.15)"}}>
    <div style={{fontSize:32,textAlign:"center",marginBottom:10}}>⚠️</div>
    <p style={{textAlign:"center",color:"#333",fontSize:14,lineHeight:1.6,marginBottom:26,whiteSpace:"pre-line"}}>{msg}</p>
    <div style={{display:"flex",gap:12,justifyContent:"center"}}><Btn onClick={onCancel} variant="ghost">Annulla</Btn><Btn onClick={onConfirm} variant="danger">Elimina</Btn></div>
  </div></Overlay>);
}

function EditDocModal({doc,cats,onSave,onClose}){
  const [form,setForm]=useState({...doc,durata:String(doc.durata)});
  const upd=k=>e=>setForm(p=>({...p,[k]:e.target.value}));
  const exp=form.dataEmissione&&+form.durata?addMonths(form.dataEmissione,+form.durata):null;
  const days=exp?diffDays(exp):null;
  const s=stato({...form,durata:+form.durata}),sc=s?statoColor(s):null;
  return(<ModalBox title="✏️  Modifica documento" onClose={onClose}>
    {s&&<div style={{background:sc.flagBg,border:`1px solid ${sc.bg}`,borderRadius:9,padding:"9px 14px",marginBottom:16,display:"flex",gap:10,alignItems:"center",flexWrap:"wrap"}}>
      <span style={{background:sc.bg,color:sc.text,borderRadius:5,padding:"2px 9px",fontSize:11,fontWeight:800}}>{s}</span>
      {exp&&<span style={{fontFamily:"'IBM Plex Mono',monospace",fontSize:11,color:"#555"}}>Scade: {fmtDate(exp)}</span>}
      {days!==null&&<span style={{fontFamily:"'IBM Plex Mono',monospace",fontSize:11,fontWeight:700,color:sc.bg}}>{days} gg</span>}
    </div>}
    <Field label="Codice" hint="Non modificabile"><input style={{...IS,opacity:0.4,cursor:"not-allowed"}} value={form.id} readOnly/></Field>
    <Field label="Categoria"><select style={IS} value={form.cat} onChange={upd("cat")}>{Object.values(cats).sort().map(c=><option key={c} value={c}>{c}</option>)}</select></Field>
    <Field label="Nome documento *"><input style={IS} value={form.nome} onChange={upd("nome")} autoFocus/></Field>
    <Field label="Data emissione"><input type="date" style={IS} value={form.dataEmissione} onChange={upd("dataEmissione")}/></Field>
    <Field label="Durata (mesi) — 0 = nessuna scadenza"><input type="number" style={IS} value={form.durata} onChange={upd("durata")} min={0}/></Field>
    <div style={{display:"flex",gap:10,marginTop:20}}><Btn onClick={onClose} variant="ghost" full>Annulla</Btn><Btn onClick={()=>{if(!form.nome.trim())return;onSave({...form,durata:+form.durata});}} variant="success" full>💾  Salva</Btn></div>
  </ModalBox>);
}

function EditClientModal({name,docs,flags,onSave,onClose}){
  const [nome,setNome]=useState(name);
  const cnt=docs.filter(d=>flags[flagKey(d.id,name)]).length;
  return(<ModalBox title="✏️  Modifica portale" onClose={onClose} width={420}>
    <div style={{background:"#f5f5f5",border:"1px solid #ddd",borderRadius:9,padding:"9px 14px",marginBottom:16,fontSize:12,color:"#555"}}>Documenti associati: <strong>{cnt}</strong> — i flag vengono preservati.</div>
    <Field label="Nome portale / cliente *"><input style={IS} value={nome} onChange={e=>setNome(e.target.value)} autoFocus/></Field>
    <div style={{display:"flex",gap:10,marginTop:20}}><Btn onClick={onClose} variant="ghost" full>Annulla</Btn><Btn onClick={()=>{if(!nome.trim())return;onSave(nome.trim());}} variant="success" full>💾  Salva</Btn></div>
  </ModalBox>);
}

function NewCatModal({cats,onSave,onClose}){
  const [prefix,setPrefix]=useState("");
  const [label,setLabel]=useState("");
  const [err,setErr]=useState("");

  function validate(){
    const p=prefix.trim().toUpperCase();
    const l=label.trim();
    if(!p){setErr("Inserire il codice categoria.");return;}
    if(!l){setErr("Inserire il nome categoria.");return;}
    if(!/^[A-Z]{2,6}$/.test(p)){setErr("Il codice deve essere 2-6 lettere maiuscole (es. MKT, LEGAL).");return;}
    if(cats[p]){setErr(`Il codice "${p}" esiste già → categoria: "${cats[p]}". Scegliere un altro codice.`);return;}
    if(Object.values(cats).some(v=>v.toLowerCase()===l.toLowerCase())){
      setErr(`La categoria "${l}" esiste già. Scegliere un altro nome.`);return;
    }
    setErr("");
    onSave(p,l);
  }

  return(<ModalBox title="➕  Nuova categoria" onClose={onClose} width={440}>
    <p style={{fontSize:12,color:"#666",marginBottom:16,lineHeight:1.5}}>
      Il codice categoria determina il prefisso dei documenti (es. <strong>MKT</strong> → MKT-001, MKT-002…).<br/>
      Usa 2-6 lettere maiuscole senza spazi.
    </p>
    <Field label="Codice categoria * (es. MKT, LEG, AMM)">
      <input style={IS} value={prefix} onChange={e=>setPrefix(e.target.value.toUpperCase())} placeholder="es. MKT" maxLength={6} autoFocus/>
    </Field>
    <Field label="Nome categoria *">
      <input style={IS} value={label} onChange={e=>setLabel(e.target.value)} placeholder="es. Marketing"/>
    </Field>
    {err&&<div style={{background:"#fff3e0",border:"1px solid #ffb74d",borderRadius:8,padding:"9px 13px",fontSize:12,color:"#e65100",marginBottom:12}}>{err}</div>}
    {/* Preview */}
    {prefix.trim()&&label.trim()&&!err&&<div style={{background:"#e8f5e9",border:"1px solid #a5d6a7",borderRadius:8,padding:"9px 13px",fontSize:12,color:"#1b5e20",marginBottom:12}}>
      ✔ Verrà creata la categoria <strong>{label.trim()}</strong> con prefisso <strong>{prefix.trim().toUpperCase()}</strong>.<br/>
      Il primo documento avrà codice <strong>{prefix.trim().toUpperCase()}-001</strong>.
    </div>}
    {/* Existing categories */}
    <details style={{marginBottom:16}}>
      <summary style={{fontSize:11,color:"#888",cursor:"pointer"}}>Categorie esistenti ({Object.keys(cats).length})</summary>
      <div style={{marginTop:8,display:"flex",flexWrap:"wrap",gap:6}}>
        {Object.entries(cats).sort((a,b)=>a[0].localeCompare(b[0])).map(([k,v])=>(
          <span key={k} style={{background:"#f0f0f0",border:"1px solid #ddd",borderRadius:5,padding:"2px 8px",fontSize:11,fontFamily:"'IBM Plex Mono',monospace"}}><strong>{k}</strong> → {v}</span>
        ))}
      </div>
    </details>
    <div style={{display:"flex",gap:10}}><Btn onClick={onClose} variant="ghost" full>Annulla</Btn><Btn onClick={validate} variant="success" full>✚  Crea categoria</Btn></div>
  </ModalBox>);
}

function MascheraCard({num,title,subtitle,color,onSubmit,btnLabel,children}){
  return(<div style={{background:"#fff",borderRadius:13,border:"1px solid #ddd",overflow:"hidden",boxShadow:"0 2px 8px rgba(0,0,0,0.06)"}}>
    <div style={{background:`${color}18`,borderBottom:`2px solid ${color}`,padding:"15px 20px",display:"flex",alignItems:"center",gap:12}}>
      <div style={{width:32,height:32,borderRadius:"50%",background:color,display:"flex",alignItems:"center",justifyContent:"center",fontFamily:"'DM Serif Display',serif",fontSize:16,color:"#fff",flexShrink:0}}>{num}</div>
      <div><div style={{fontFamily:"'DM Serif Display',serif",fontSize:15,color:"#222"}}>{title}</div><div style={{fontSize:10,color:"#888",marginTop:1}}>{subtitle}</div></div>
    </div>
    <div style={{padding:"18px 20px"}}>
      {children}
      <button onClick={onSubmit} style={{width:"100%",background:color,border:"none",borderRadius:8,padding:"10px",color:"#fff",fontWeight:800,fontSize:12,letterSpacing:"0.05em",cursor:"pointer",fontFamily:"'IBM Plex Mono',monospace",transition:"opacity 0.15s",marginTop:14}} onMouseOver={e=>e.currentTarget.style.opacity=0.85} onMouseOut={e=>e.currentTarget.style.opacity=1}>▶  {btnLabel}</button>
    </div>
  </div>);
}

// ══ MAIN ══════════════════════════════════════════════════════════════════════
export default function App(){
  const [docs,   setDocs]   = useState(()=>sortDocs(INIT_DOCS));
  const [clients,setClients]= useState(INIT_CLIENTS);
  const [flags,  setFlags]  = useState(INIT_FLAGS);
  const [cats,   setCats]   = useState(INIT_CATS); // prefix->label registry
  const [tab,    setTab]    = useState("scadenzario");
  const [toast,  setToast]  = useState(null);

  const [editDoc,    setEditDoc]    = useState(null);
  const [editClient, setEditClient] = useState(null);
  const [confirmDel, setConfirmDel] = useState(null);
  const [showNewCat, setShowNewCat] = useState(false);

  const [fDoc,setFDoc]=useState({id:"",cat:"",nome:"",dataEmissione:"",durata:"",cliente:""});
  const [fCli,setFCli]=useState({nome:""});
  const [fAss,setFAss]=useState({cliente:"",doc:""});
  const [filterCat,setFilterCat]=useState("");
  const [filterStato,setFilterStato]=useState("");
  const [filterText,setFilterText]=useState("");
  const [filterClient,setFilterClient]=useState("");

  function toast_(msg,type="ok"){setToast({msg,type});setTimeout(()=>setToast(null),3600);}

  // ── Docs CRUD ──────────────────────────────────────────────────────────────
  function insertDoc(d,cli){
    const prefix = Object.entries(cats).find(([k,v])=>v===d.cat)?.[0] || "DOC";
    const id = d.id.trim() || nextDocCode(docs, prefix);
    const nd={id,cat:d.cat,nome:d.nome.trim(),dataEmissione:d.dataEmissione,durata:+d.durata};
    setDocs(p=>sortDocs([...p,nd]));
    if(cli)setFlags(f=>({...f,[flagKey(id,cli)]:true}));
    toast_(`"${nd.nome}" inserito (${id})${cli?" → "+cli:""}.`);
  }
  function saveDoc(u){setDocs(p=>sortDocs(p.map(d=>d.id===u.id?u:d)));setEditDoc(null);toast_(`"${u.nome}" aggiornato.`);}
  function deleteDoc(id){
    const d=docs.find(x=>x.id===id);
    setDocs(p=>p.filter(x=>x.id!==id));
    setFlags(p=>{const n={...p};Object.keys(n).forEach(k=>{if(k.startsWith(id+"___"))delete n[k];});return n;});
    setConfirmDel(null);toast_(`"${d?.nome}" eliminato.`,"warn");
  }

  // ── Cats CRUD ──────────────────────────────────────────────────────────────
  function addCat(prefix,label){
    setCats(p=>({...p,[prefix]:label}));
    setShowNewCat(false);
    toast_(`Categoria "${label}" (${prefix}) creata. Primo codice disponibile: ${prefix}-001`);
  }

  // ── Clients CRUD ───────────────────────────────────────────────────────────
  function insertClient(nome){
    if(clients.includes(nome)){toast_("Portale già presente.","err");return;}
    setClients(p=>[...p,nome]);toast_(`"${nome}" aggiunto.`);
  }
  function saveClient(old,nw){
    if(nw!==old&&clients.includes(nw)){toast_("Nome già esistente.","err");return;}
    setClients(p=>p.map(c=>c===old?nw:c));
    if(old!==nw)setFlags(p=>{const n={};Object.entries(p).forEach(([k,v])=>{n[k.endsWith("___"+old)?k.replace("___"+old,"___"+nw):k]=v;});return n;});
    setEditClient(null);toast_(`Rinominato in "${nw}".`);
  }
  function deleteClient(name){
    setClients(p=>p.filter(c=>c!==name));
    setFlags(p=>{const n={...p};Object.keys(n).forEach(k=>{if(k.endsWith("___"+name))delete n[k];});return n;});
    setConfirmDel(null);toast_(`"${name}" eliminato.`,"warn");
  }
  function assignDoc(docNome,cli){
    const doc=docs.find(d=>d.nome===docNome);
    if(!doc||!cli){toast_("Seleziona portale e documento.","err");return;}
    setFlags(f=>({...f,[flagKey(doc.id,cli)]:true}));
    toast_(`"${docNome}" → "${cli}".`);
  }
  function toggleFlag(docId,cli){setFlags(f=>({...f,[flagKey(docId,cli)]:!f[flagKey(docId,cli)]}));}

  const filteredDocs=useMemo(()=>docs.filter(d=>{
    if(filterCat&&d.cat!==filterCat)return false;
    if(filterText&&!d.nome.toLowerCase().includes(filterText.toLowerCase())&&!d.id.toLowerCase().includes(filterText.toLowerCase()))return false;
    if(filterStato){const s=stato(d);if(!s||s!==filterStato)return false;}
    return true;
  }),[docs,filterCat,filterText,filterStato]);

  const filteredClients=useMemo(()=>clients.filter(c=>!filterClient||c.toLowerCase().includes(filterClient.toLowerCase())),[clients,filterClient]);
  const STATI=["VALIDO","IN SCADENZA","ATTENZIONE","URGENTE","SCADUTO"];
  const TABS=[{id:"scadenzario",label:"📄  Scadenzario"},{id:"clienti",label:"👥  Matrice Portali"},{id:"maschera",label:"✏️  Inserimento"}];
  const TD={padding:"9px 13px",borderBottom:"1px solid #ececec",verticalAlign:"middle"};

  // Auto-suggest next code when category changes in insert form
  const suggestedCode = useMemo(()=>{
    if(!fDoc.cat)return"";
    const prefix=Object.entries(cats).find(([k,v])=>v===fDoc.cat)?.[0];
    if(!prefix)return"";
    return nextDocCode(docs,prefix);
  },[fDoc.cat,docs,cats]);

  return(<div style={{minHeight:"100vh",background:C.pageBg,fontFamily:"'IBM Plex Sans',sans-serif",color:"#222"}}>
    <style>{`
      @import url('https://fonts.googleapis.com/css2?family=DM+Serif+Display&family=IBM+Plex+Sans:wght@400;600;700&family=IBM+Plex+Mono:wght@400;700&display=swap');
      *{box-sizing:border-box;}
      ::-webkit-scrollbar{width:5px;height:5px;background:#e8e8e8;}
      ::-webkit-scrollbar-thumb{background:#aaa;border-radius:3px;}
      input:focus,select:focus{border-color:#555!important;box-shadow:0 0 0 2px rgba(0,0,0,0.08);}
      .rh:hover{background:#f0f0f0!important;}
      .ab{background:none;border:1px solid transparent;border-radius:6px;cursor:pointer;padding:4px 8px;font-size:13px;transition:all 0.15s;line-height:1;color:#aaa;}
      .ab:hover{border-color:#ccc;}
      .ab.ed:hover{background:#e3f2fd;color:#1565c0;}
      .ab.dl:hover{background:#ffebee;color:#c62828;}
      .flt{background:#fff;border:1px solid #ccc;border-radius:7px;padding:7px 11px;color:#333;font-size:12px;outline:none;font-family:'IBM Plex Mono',monospace;}
      .flt:focus{border-color:#555;}
      .chip-btn{border-radius:20px;padding:3px 12px;font-size:11px;font-weight:700;cursor:pointer;transition:all 0.15s;letter-spacing:0.04em;}
    `}</style>

    {/* Header */}
    <div style={{background:C.headerBg,padding:"0 24px"}}>
      <div style={{maxWidth:1600,margin:"0 auto",display:"flex",alignItems:"center",justifyContent:"space-between",height:58}}>
        <div><span style={{fontFamily:"'DM Serif Display',serif",fontSize:21,color:"#fff"}}>Monitoraggio </span><span style={{fontFamily:"'DM Serif Display',serif",fontSize:21,color:"#bbb"}}>Scadenze Documenti</span></div>
        <div style={{display:"flex",gap:8,alignItems:"center"}}>
          <span style={{fontSize:11,color:"#888",fontFamily:"'IBM Plex Mono',monospace"}}>{docs.length} doc · {clients.length} portali</span>
          <Btn onClick={()=>exportExcel(docs,clients,flags)} variant="ghost" small>⬇ XLSX</Btn>
        </div>
      </div>
      <div style={{maxWidth:1600,margin:"0 auto",display:"flex"}}>
        {TABS.map(t=>(<button key={t.id} onClick={()=>setTab(t.id)} style={{background:"none",border:"none",padding:"11px 20px",cursor:"pointer",fontFamily:"'IBM Plex Sans',sans-serif",fontWeight:700,fontSize:12,letterSpacing:"0.05em",color:tab===t.id?"#fff":"#888",borderBottom:tab===t.id?"2px solid #fff":"2px solid transparent",transition:"all 0.15s"}}>{t.label}</button>))}
      </div>
    </div>

    {/* Toast */}
    {toast&&<div style={{position:"fixed",top:18,right:18,zIndex:4000,background:toast.type==="err"?"#c62828":toast.type==="warn"?"#e65100":"#2e7d32",color:"#fff",borderRadius:9,padding:"10px 18px",fontSize:13,fontWeight:600,boxShadow:"0 4px 20px rgba(0,0,0,0.2)",maxWidth:360}}>{toast.msg}</div>}

    {/* Overlays */}
    {editDoc&&<EditDocModal doc={editDoc} cats={cats} onSave={saveDoc} onClose={()=>setEditDoc(null)}/>}
    {editClient&&<EditClientModal name={editClient} docs={docs} flags={flags} onSave={n=>saveClient(editClient,n)} onClose={()=>setEditClient(null)}/>}
    {showNewCat&&<NewCatModal cats={cats} onSave={addCat} onClose={()=>setShowNewCat(false)}/>}
    {confirmDel&&<ConfirmDialog msg={confirmDel.type==="doc"?`Eliminare "${docs.find(d=>d.id===confirmDel.payload)?.nome}"?\nTutti i flag associati verranno rimossi.`:`Eliminare il portale "${confirmDel.payload}"?\nTutti i flag associati verranno rimossi.`} onConfirm={()=>confirmDel.type==="doc"?deleteDoc(confirmDel.payload):deleteClient(confirmDel.payload)} onCancel={()=>setConfirmDel(null)}/>}

    <div style={{maxWidth:1600,margin:"0 auto",padding:"20px 24px"}}>

      {/* ══ SCADENZARIO ══════════════════════════════════════════════════════ */}
      {tab==="scadenzario"&&(<div>
        <div style={{display:"flex",justifyContent:"space-between",alignItems:"center",marginBottom:12,flexWrap:"wrap",gap:8}}>
          <h2 style={{fontFamily:"'DM Serif Display',serif",fontSize:20,color:"#222",margin:0}}>
            Documenti — <span style={{color:"#555"}}>{filteredDocs.length}</span>{filteredDocs.length!==docs.length&&<span style={{color:"#aaa"}}> / {docs.length}</span>}
          </h2>
          <div style={{display:"flex",gap:8,flexWrap:"wrap",alignItems:"center"}}>
            <input className="flt" placeholder="🔍 Cerca..." value={filterText} onChange={e=>setFilterText(e.target.value)} style={{width:170}}/>
            <select className="flt" value={filterCat} onChange={e=>setFilterCat(e.target.value)}>
              <option value="">Tutte le categorie</option>
              {Object.values(cats).sort().map(c=><option key={c} value={c}>{c}</option>)}
            </select>
            <select className="flt" value={filterStato} onChange={e=>setFilterStato(e.target.value)}>
              <option value="">Tutti gli stati</option>
              {STATI.map(s=><option key={s} value={s}>{s}</option>)}
            </select>
            {(filterText||filterCat||filterStato)&&<button onClick={()=>{setFilterText("");setFilterCat("");setFilterStato("");}} style={{background:"none",border:"none",color:"#888",cursor:"pointer",fontSize:12}}>✕ Reset</button>}
            <Btn onClick={()=>setTab("maschera")} small>+ Nuovo</Btn>
          </div>
        </div>

        {/* Status chips */}
        <div style={{display:"flex",gap:6,marginBottom:12,flexWrap:"wrap"}}>
          {STATI.map(s=>{
            const cnt=docs.filter(d=>stato(d)===s).length;
            if(!cnt)return null;
            const sc=statoColor(s);
            const active=filterStato===s;
            return(<button key={s} className="chip-btn" onClick={()=>setFilterStato(active?"":s)} style={{background:active?sc.bg:sc.flagBg,color:active?sc.text:sc.flagText,border:`1px solid ${sc.bg}`}}>{s} ({cnt})</button>);
          })}
          <span style={{fontSize:11,color:"#aaa",alignSelf:"center",marginLeft:4}}>Senza scadenza: {docs.filter(d=>!d.durata||!d.dataEmissione).length}</span>
        </div>

        <div style={{overflowX:"auto",borderRadius:11,border:"1px solid #ddd",background:"#fff",boxShadow:"0 1px 4px rgba(0,0,0,0.06)"}}>
          <table style={{width:"100%",borderCollapse:"collapse",fontSize:12}}>
            <thead><tr style={{background:"#f5f5f5"}}>
              {["Codice","Categoria","Documento","Data Emissione","Durata","Scadenza","Giorni","Stato","Azioni"].map(h=>(
                <th key={h} style={{padding:"10px 13px",textAlign:["Azioni","Giorni","Durata"].includes(h)?"center":"left",fontWeight:700,fontSize:10,letterSpacing:"0.07em",textTransform:"uppercase",color:"#666",borderBottom:"2px solid #e0e0e0",whiteSpace:"nowrap"}}>{h}</th>
              ))}
            </tr></thead>
            <tbody>
              {filteredDocs.map((d,i)=>{
                const exp=d.dataEmissione&&d.durata?addMonths(d.dataEmissione,+d.durata):null;
                const days=exp?diffDays(exp):null;
                const s=stato(d),sc=s?statoColor(s):null;
                const rowBg=s?sc.flagBg:(i%2===0?C.rowEven:C.rowOdd);
                return(<tr key={d.id+i} className="rh" style={{background:rowBg}}>
                  <td style={{...TD,fontFamily:"'IBM Plex Mono',monospace",fontSize:11,color:"#444",fontWeight:700,whiteSpace:"nowrap"}}>{d.id}</td>
                  <td style={{...TD}}><span style={{background:"#eeeeee",color:"#555",borderRadius:4,padding:"2px 6px",fontSize:10,whiteSpace:"nowrap"}}>{d.cat}</span></td>
                  <td style={{...TD,fontWeight:600,color:"#222",maxWidth:280}}>{d.nome}</td>
                  <td style={{...TD,fontFamily:"'IBM Plex Mono',monospace",fontSize:11,color:"#666",whiteSpace:"nowrap"}}>{fmtDate(d.dataEmissione)}</td>
                  <td style={{...TD,textAlign:"center",color:"#888",fontSize:11}}>{d.durata?d.durata+"m":"—"}</td>
                  <td style={{...TD,fontFamily:"'IBM Plex Mono',monospace",fontSize:11,color:"#666",whiteSpace:"nowrap"}}>{exp?fmtDate(exp):"—"}</td>
                  <td style={{...TD,textAlign:"center",fontFamily:"'IBM Plex Mono',monospace",fontWeight:700,fontSize:11,color:days===null?"#aaa":sc?sc.bg:"#333"}}>{days===null?"—":days+" gg"}</td>
                  <td style={{...TD,whiteSpace:"nowrap"}}>{s&&<span style={{background:sc.bg,color:sc.text,borderRadius:5,padding:"2px 8px",fontSize:10,fontWeight:800,letterSpacing:"0.05em"}}>{s}</span>}</td>
                  <td style={{...TD,textAlign:"center",whiteSpace:"nowrap"}}>
                    <button className="ab ed" title="Modifica" onClick={()=>setEditDoc(d)}>✏️</button>{" "}
                    <button className="ab dl" title="Elimina" onClick={()=>setConfirmDel({type:"doc",payload:d.id})}>🗑️</button>
                  </td>
                </tr>);
              })}
            </tbody>
          </table>
        </div>
      </div>)}

      {/* ══ MATRICE PORTALI ══════════════════════════════════════════════════ */}
      {tab==="clienti"&&(<div>
        <div style={{display:"flex",justifyContent:"space-between",alignItems:"center",marginBottom:12,flexWrap:"wrap",gap:8}}>
          <h2 style={{fontFamily:"'DM Serif Display',serif",fontSize:20,color:"#222",margin:0}}>Matrice Portali — <span style={{color:"#555"}}>{filteredClients.length}</span>{filteredClients.length!==clients.length&&<span style={{color:"#aaa"}}> / {clients.length}</span>} portali</h2>
          <div style={{display:"flex",gap:8,alignItems:"center"}}>
            <input className="flt" placeholder="🔍 Filtra portali..." value={filterClient} onChange={e=>setFilterClient(e.target.value)} style={{width:190}}/>
            {filterClient&&<button onClick={()=>setFilterClient("")} style={{background:"none",border:"none",color:"#888",cursor:"pointer",fontSize:12}}>✕</button>}
            <Btn onClick={()=>setTab("maschera")} small>+ Nuovo portale</Btn>
          </div>
        </div>
        <p style={{fontSize:11,color:"#aaa",marginBottom:10,fontStyle:"italic"}}>Clicca ✔/· per modificare i flag. ✏️/🗑️ nell'intestazione per rinominare o eliminare un portale.</p>

        <div style={{overflowX:"auto",borderRadius:11,border:"1px solid #ddd",boxShadow:"0 1px 4px rgba(0,0,0,0.06)"}}>
          <table style={{borderCollapse:"collapse",fontSize:11}}>
            <thead><tr style={{background:C.matrixHeaderBg}}>
              <th style={{padding:"9px 11px",textAlign:"left",fontWeight:700,fontSize:9,letterSpacing:"0.07em",textTransform:"uppercase",color:"#ccc",borderBottom:"1px solid #444",minWidth:70,position:"sticky",left:0,background:C.matrixHeaderBg,zIndex:10}}>Codice</th>
              <th style={{padding:"9px 11px",textAlign:"left",fontWeight:700,fontSize:9,letterSpacing:"0.07em",textTransform:"uppercase",color:"#ccc",borderBottom:"1px solid #444",minWidth:230,position:"sticky",left:62,background:C.matrixHeaderBg,zIndex:10}}>Documento</th>
              {filteredClients.map(c=>(<th key={c} style={{padding:"7px 5px 4px",textAlign:"center",borderBottom:"1px solid #444",minWidth:96,background:C.matrixHeaderBg}}>
                <div style={{fontWeight:700,fontSize:9,letterSpacing:"0.04em",textTransform:"uppercase",color:"#ccc",marginBottom:4,whiteSpace:"nowrap",overflow:"hidden",textOverflow:"ellipsis",maxWidth:88,margin:"0 auto 4px"}} title={c}>{c}</div>
                <div style={{display:"flex",justifyContent:"center",gap:2}}>
                  <button className="ab ed" style={{fontSize:10,padding:"2px 4px",color:"#999"}} title={"Rinomina "+c} onClick={()=>setEditClient(c)}>✏️</button>
                  <button className="ab dl" style={{fontSize:10,padding:"2px 4px",color:"#999"}} title={"Elimina "+c} onClick={()=>setConfirmDel({type:"client",payload:c})}>🗑️</button>
                </div>
              </th>))}
            </tr></thead>
            <tbody>
              {docs.map((d,i)=>{
                const s=stato(d),sc=s?statoColor(s):null;
                const stickyBg=i%2===0?"#f8f8f8":"#ffffff";
                return(<tr key={d.id+i} className="rh" style={{borderBottom:"1px solid #eee"}}>
                  <td style={{padding:"7px 11px",fontFamily:"'IBM Plex Mono',monospace",fontSize:10,color:"#555",fontWeight:700,position:"sticky",left:0,background:stickyBg,zIndex:5,whiteSpace:"nowrap"}}>{d.id}</td>
                  <td style={{padding:"7px 11px",color:"#333",fontWeight:600,position:"sticky",left:62,background:stickyBg,zIndex:5,maxWidth:230,overflow:"hidden",textOverflow:"ellipsis",whiteSpace:"nowrap"}} title={d.nome}>{d.nome}</td>
                  {filteredClients.map(c=>{
                    const has=!!flags[flagKey(d.id,c)];
                    // flag cell color: if doc has status, use status color; else neutral dark
                    const cellBg = has ? (sc ? sc.flagBg : "#e0e0e0") : C.matrixCellBg;
                    const cellText = has ? (sc ? sc.flagText : "#333") : "#666";
                    const cellBorder = has ? (sc ? `1px solid ${sc.bg}` : "1px solid #aaa") : "1px solid #4a4a4a";
                    return(<td key={c} style={{padding:"5px",textAlign:"center",background:C.matrixCellBg}}>
                      <button onClick={()=>toggleFlag(d.id,c)} title={has?"Rimuovi flag":"Aggiungi flag"} style={{background:cellBg,border:cellBorder,borderRadius:5,width:28,height:24,cursor:"pointer",fontSize:13,color:cellText,transition:"all 0.12s",fontWeight:has?800:400}}>{has?"✔":"·"}</button>
                    </td>);
                  })}
                </tr>);
              })}
              <tr style={{background:"#f0f0f0",borderTop:"2px solid #ccc"}}>
                <td style={{padding:"8px 11px",position:"sticky",left:0,background:"#f0f0f0",zIndex:5}}></td>
                <td style={{padding:"8px 11px",fontWeight:700,fontSize:11,color:"#555",letterSpacing:"0.04em",position:"sticky",left:62,background:"#f0f0f0",zIndex:5}}>TOTALE</td>
                {filteredClients.map(c=>(<td key={c} style={{padding:"8px 5px",textAlign:"center",fontFamily:"'IBM Plex Mono',monospace",fontWeight:800,fontSize:13,color:"#333",background:C.matrixCellBg}}>{docs.filter(d=>flags[flagKey(d.id,c)]).length}</td>))}
              </tr>
            </tbody>
          </table>
        </div>
      </div>)}

      {/* ══ MASCHERA ═════════════════════════════════════════════════════════ */}
      {tab==="maschera"&&(<div>
        {/* Category bar */}
        <div style={{background:"#fff",border:"1px solid #ddd",borderRadius:11,padding:"14px 18px",marginBottom:20,boxShadow:"0 1px 4px rgba(0,0,0,0.05)"}}>
          <div style={{display:"flex",justifyContent:"space-between",alignItems:"center",marginBottom:10}}>
            <span style={{fontWeight:700,fontSize:13,color:"#333",fontFamily:"'DM Serif Display',serif"}}>Categorie registrate</span>
            <Btn onClick={()=>setShowNewCat(true)} small variant="accent">+ Nuova categoria</Btn>
          </div>
          <div style={{display:"flex",flexWrap:"wrap",gap:6}}>
            {Object.entries(cats).sort((a,b)=>a[0].localeCompare(b[0])).map(([k,v])=>(
              <span key={k} style={{background:"#f5f5f5",border:"1px solid #ddd",borderRadius:6,padding:"3px 10px",fontSize:11,fontFamily:"'IBM Plex Mono',monospace",color:"#444"}}><strong style={{color:"#222"}}>{k}</strong> — {v}</span>
            ))}
          </div>
        </div>

        <div style={{display:"grid",gridTemplateColumns:"repeat(auto-fit,minmax(310px,1fr))",gap:18}}>

          <MascheraCard num="1" title="Nuovo documento" subtitle="Aggiunge riga allo Scadenzario · opz. assegna portale" color="#37474f" btnLabel="Inserisci documento" onSubmit={()=>{
            if(!fDoc.nome.trim()||!fDoc.cat){toast_("Inserire categoria e nome documento.","err");return;}
            insertDoc({id:fDoc.id,cat:fDoc.cat,nome:fDoc.nome,dataEmissione:fDoc.dataEmissione,durata:fDoc.durata},fDoc.cliente||null);
            setFDoc(f=>({...f,id:"",nome:"",dataEmissione:"",durata:"",cliente:""}));
          }}>
            <Field label="Categoria *">
              <select style={IS} value={fDoc.cat} onChange={e=>setFDoc(f=>({...f,cat:e.target.value,id:""}))}>
                <option value="">— seleziona —</option>
                {Object.entries(cats).sort((a,b)=>a[1].localeCompare(b[1])).map(([k,v])=><option key={k} value={v}>{v} ({k})</option>)}
              </select>
            </Field>
            <Field label="Codice" hint={fDoc.cat?`Lascia vuoto → verrà assegnato ${suggestedCode}`:"Seleziona prima la categoria"}>
              <input style={IS} value={fDoc.id} onChange={e=>setFDoc(f=>({...f,id:e.target.value}))} placeholder={suggestedCode||"automatico"}/>
            </Field>
            <Field label="Nome documento *"><input style={IS} value={fDoc.nome} onChange={e=>setFDoc(f=>({...f,nome:e.target.value}))} placeholder="Es. Certificato..."/></Field>
            <Field label="Data emissione"><input type="date" style={IS} value={fDoc.dataEmissione} onChange={e=>setFDoc(f=>({...f,dataEmissione:e.target.value}))}/></Field>
            <Field label="Durata (mesi) — 0 o vuoto = nessuna scadenza"><input type="number" style={IS} value={fDoc.durata} onChange={e=>setFDoc(f=>({...f,durata:e.target.value}))} placeholder="6" min={0}/></Field>
            <Field label="Assegna a portale (opz.)"><select style={IS} value={fDoc.cliente} onChange={e=>setFDoc(f=>({...f,cliente:e.target.value}))}><option value="">— nessuno —</option>{clients.map(c=><option key={c} value={c}>{c}</option>)}</select></Field>
          </MascheraCard>

          <MascheraCard num="2" title="Nuovo portale / cliente" subtitle="Aggiunge colonna alla Matrice" color="#4a5568" btnLabel="Inserisci portale" onSubmit={()=>{
            if(!fCli.nome.trim()){toast_("Inserire il nome del portale.","err");return;}
            insertClient(fCli.nome.trim());setFCli({nome:""});
          }}>
            <Field label="Nome portale / cliente *"><input style={IS} value={fCli.nome} onChange={e=>setFCli(f=>({...f,nome:e.target.value}))} placeholder="Es. Terna, Google..."/></Field>
          </MascheraCard>

          <MascheraCard num="3" title="Assegna documento a portale" subtitle="Aggiunge flag ✔ nella Matrice" color="#5d4037" btnLabel="Assegna documento" onSubmit={()=>{
            if(!fAss.cliente||!fAss.doc){toast_("Seleziona portale e documento.","err");return;}
            assignDoc(fAss.doc,fAss.cliente);setFAss({cliente:"",doc:""});
          }}>
            <Field label="Portale *"><select style={IS} value={fAss.cliente} onChange={e=>setFAss(f=>({...f,cliente:e.target.value}))}><option value="">— seleziona —</option>{clients.map(c=><option key={c} value={c}>{c}</option>)}</select></Field>
            <Field label="Documento *"><select style={IS} value={fAss.doc} onChange={e=>setFAss(f=>({...f,doc:e.target.value}))}><option value="">— seleziona —</option>{docs.map(d=><option key={d.id} value={d.nome}>[{d.id}] {d.nome}</option>)}</select></Field>
            {fAss.cliente&&fAss.doc&&(()=>{
              const doc=docs.find(d=>d.nome===fAss.doc);
              const already=doc&&!!flags[flagKey(doc.id,fAss.cliente)];
              return already?<div style={{background:"#fff3e0",border:"1px solid #ffb74d",borderRadius:7,padding:"8px 11px",fontSize:11,color:"#e65100",marginTop:4}}>⚠ Già presente per questo portale.</div>:<div style={{background:"#e8f5e9",border:"1px solid #a5d6a7",borderRadius:7,padding:"8px 11px",fontSize:11,color:"#2e7d32",marginTop:4}}>✔ Verrà aggiunto il flag.</div>;
            })()}
          </MascheraCard>

        </div>
      </div>)}
    </div>
  </div>);
}
