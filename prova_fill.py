from docxtpl import DocxTemplate


doc = DocxTemplate('AA_PREPDOC/AD_infra_40k_RichiestaOfferta.docx')
context = {'numero_CUP': '38900',
           'servizio_fornitura' : 'servizio',
           'prestazione_servizio_fornitura' : 'Prestazione di servizio',
           'nome_cognome' : 'Alfonso Crisci',
           'mail_contatto':"alfonso.crisci@cnr.it",
           'acronimo_progetto':'CLIMANIMAL',
           'oggetto_fornitura_servizio':'Acquisto di 10 PC da tavolo con supporto di assisenza',
           'nome_ditta':'BASSO SRL',
           'indirizzo_ditta':'Via Carlino Pocciante 7',
           'cap_ditta':'50100',
           'pec_ditta':'basso@pec.it'}
doc.render(context)
doc.save("A_preventivo/RichiestaOfferta_Crisci.docx")
