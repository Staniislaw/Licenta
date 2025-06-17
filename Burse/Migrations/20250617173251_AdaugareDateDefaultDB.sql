INSERT INTO [GrupBursa] ([GrupBursa], [Domeniu]) VALUES
('G2', 'AIA'),
('G3', 'IEN'),
('G3', 'ETI'),
('G3', 'ME'),
('G1', 'C'),
('G1', 'C-DUAL'),
('G2', 'IETTI'),
('G2', 'RST'),
('G1', 'SIC'),
('G3', 'SE'),
('G3', 'SE-DUAL'),
('G3', 'TAMAE'),
('G2', 'RCC'),
('G2', 'SC'),
('G2', 'EA'),
('G3', 'SMCPE'),
('G3', 'ESCCA'),
('G2', 'ESM'),
('G2', 'AIA-DUAL');

INSERT INTO [GrupDomeniu] ([Grup], [Domeniu]) VALUES
('IEN/ME/ETI', 'ME(3)'),
('IEN/ME/ETI', 'ETI(4)'),
('IETTI/RST', 'IETTI(1)(2)'),
('IETTI/RST', 'RST(3)(4)'),
('IEN/ME/ETI', 'IEN (1)(2)');

INSERT INTO [GrupPDF] ([Grup], [Valoare]) VALUES
('Grup: Calculatoare, Calculatoare-DUAL', 'C'),
('Gup: Automatica', 'AIA'),
('Gup: Automatica', 'AIA-DUAL'),
('Grup: Calculatoare, Calculatoare-DUAL', 'C-DUAL');

INSERT INTO [GrupProgramStudii] ([Grup], [Domeniu]) VALUES
('IETTI', 'SC'),
('IETTI', 'IETTI'),
('IETTI', 'RCC'),
('IETTI', 'RST'),
('IEN', 'SMCPE'),
('IEN', 'IEN'),
('IEN', 'ME'),
('IEN', 'ETI'),
('IE', 'SE'),
('IE', 'TAMAE'),
('IE', 'SE-DUAL'),
('IS', 'AIA'),
('SIA', 'ESM'),
('CTI', 'SIC'),
('CTI', 'C'),
('CTI', 'C-DUAL'),
('IS', 'AIA-DUAL'),
('IA', 'ESCCA');

INSERT INTO [BurseDB].[dbo].[TemplateEntity] ([Name], [CreatedAt], [ElementsJson])
VALUES (
    'Template Bursa',
    '2025-05-27T15:53:13.5953878',
    '[{"id":1745696433219,"type":"text","content":"UNIVERSITATEA \"ŞTEFAN CEL MARE \" DIN SUCEAVA","style":{"fontSize":7,"textAlign":"left","color":"#000000"}},{"id":1745696441944,"type":"text","content":"FACULTATEA DE INGINERIE ELECTRICĂ ŞI ŞTIINŢA CALCULATOARELOR","style":{"fontSize":7,"textAlign":"left","color":"#000000"}},{"id":1745696449677,"type":"text","content":"Nr. _________/FIESC/ ___________________","style":{"fontSize":7,"textAlign":"left","color":"#000000"}},{"id":1745696459942,"type":"text","content":"LISTA STUDENŢILOR DE LA PROGRAMUL DE STUDII -  ProgramStudiu.Dynamic","style":{"fontSize":7,"textAlign":"left","color":"#000000"}},{"id":1745696464870,"type":"text","content":"STUDII UNIVERSITARE DE LICENȚĂ, PROPUŞI SĂ PRIMEASCĂ BURSĂ DE PERFORMANȚĂ IN ANUL UNIVERSITAT 2024-2025\n","style":{"fontSize":7,"textAlign":"left","color":"#000000"}},{"id":1748359456216,"type":"table","content":"Text","style":{"fontSize":14,"textAlign":"left","color":"#000000"},"domenii":["Dynamic"]},{"id":1745696480091,"type":"text","content":"aprobate în ședința de vot electronic a Consiliului Facultății din dupa contesttii\n","style":{"fontSize":7,"textAlign":"left","color":"#000000"}},{"id":1745696508131,"type":"text","content":"DECAN,                                                                                  PREŞEDINTE CABF, ","style":{"fontSize":7,"textAlign":"left","color":"#000000"}},{"id":1745696518340,"type":"text","content":"Prof.univ.dr.ing. Laurenţiu                                                       Dan MILICI Conf.univ.dr.ing. Pavel ATĂNĂSOAE","style":{"fontSize":7,"textAlign":"left","color":"#000000"}},{"id":1745696534716,"type":"text","content":"                                                                                                  SECRETAR ŞEF FACULTATE,","style":{"fontSize":7,"textAlign":"left","color":"#000000"}},{"id":1745696545718,"type":"text","content":"                                                                                                  ing. Elena CURELARU","style":{"fontSize":7,"textAlign":"left","color":"#000000"}}]'
);

INSERT INTO [BurseDB].[dbo].[GrupAcronim] ([Grup], [Valoare])
VALUES 
(N'Sisteme electrice', N'SE'),
(N'Echipamente și sisteme medicale', N'ESM'),
(N'AUTOMATICA SI INFORMATICA APLICATA', N'AIA'),
(N'Tehnici avansate în mașini și acționări electrice', N'TAMAE'),
(N'Managementul energiei', N'ME'),
(N'CALCULATOARE', N'C'),
(N'Echipamente și sisteme de comandă și control pentru autovehicule', N'ESCCA'),
(N'Sisteme moderne pentru conducerea proceselor energetice', N'SMCPE'),
(N'Inginerie electronică telecomunicații și tehnologii inforimaționale', N'IETTI'),
(N'Energetică și tehnologii informatice', N'ETI'),
(N'Rețele și software de telecomunicații', N'RST'),
(N'Securitate cibernetică', N'SC'),
(N'Inginerie energetică', N'IEN'),
(N'Reţele de comunicaţii şi calculatoare', N'RCC'),
(N'Știința și ingineria calculatoarelor', N'SIC');

