import pandas as pd

df = pd.read_excel("C:/Users/douggomes/Desktop/Codes/Bancos para padronização/Banco de óbitos (SIM) 1996-2026 - Curitiba.xlsx", nrows=0)

colunas_do_df = ['contador', 'TIPOBITO', 'DTOBITO', 'DTNASC', 'IDADE', 'SEXO', 'RACACOR', 'ESTCIV', 'ESC', 'CODBAIRES', 'CODMUNRES', 'LOCOCOR', 'CODMUNOCOR',
                'IDADEMAE', 'ESCMAE', 'QTDFILVIVO', 'QTDFILMORT', 'GRAVIDEZ', 'GESTACAO', 'PARTO', 'OBITOPARTO', 'PESO', 'OBITOGRAV', 'OBITOPUERP', 'ASSISTMED', 
                'EXAME', 'CIRURGIA', 'NECROPSIA', 'CAUSABAS', 'LINHAA', 'LINHAB', 'LINHAC', 'LINHAD', 'LINHAII', 'CIRCOBITO', 'ACIDTRAB', 'FONTE', 'CODESTAB', 'ATESTANTE', 
                'UFINFORM', 'CONTADOR', 'HORAOBITO', 'CODBAIOCOR', 'TPASSINA', 'DTATESTADO', 'TPPOS', 'DTINVESTIG', 'CAUSABAS_O', 'DTCADASTRO', 'FONTEINV', 'DTRECEBIM', 
                'CB_PRE', 'ORIGEM', 'NUMERODN', 'MORTEPARTO', 'DTCADINF', 'TPOBITOCOR', 'DTCADINV', 'COMUNSVOIM', 'DTRECORIG', 'DTRECORIGA', 'CAUSAMAT', 'ESC2010', 
                'ESCMAE2010', 'STDOEPIDEM', 'STDONOVA', 'SEMAGESTAC', 'TPMORTEOCO', 'DIFDATA', 'DTCONCASO', 'NUDIASOBIN', 'SERIESCFAL', 'SERIESCMAE', 'CODMUNCART', 'CODCART', 
                'NUMREGCART', 'DTREGCART', 'DTCONINV', 'CODMUNNATU', 'ESTABDESCR', 'CRM', 'NUMEROLOTE', 'STCODIFICA', 'CODIFICADO', 'VERSAOSIST', 'VERSAOSCB', 'ATESTADO', 
                'ESCMAEAGR1', 'ESCFALAGR1', 'NUDIASOBCO', 'FONTES', 'TPRESGINFO', 'TPNIVELINV', 'NUDIASINF', 'FONTESINF', 'ALTCAUSA', 'NATURAL', 'IDADEminutos', 'IDADEhoras', 
                'IDADEdias', 'IDADEmeses', 'IDADEanos', 'OCUP', 'munResStatus', 'munResTipo', 'munResNome', 'munResUf', 'munResLat', 'munResLon', 'munResAlt', 'munResArea', 
                'OCUPMAE']

select_columns = ['contador', 'TIPOBITO', 'DTOBITO', 'DTNASC', 'IDADE', 'SEXO', 'RACACOR', 'ESTCIV', 'ESC', 'CODBAIRES', 'CODMUNRES', 'LOCOCOR', 'CODMUNOCOR',
                'QTDFILVIVO', 'QTDFILMORT', 'GRAVIDEZ', 'GESTACAO', 'PARTO', 'OBITOPARTO', 'PESO', 'OBITOGRAV', 'OBITOPUERP', 'ASSISTMED', 
                'NECROPSIA', 'CAUSABAS', 'LINHAA', 'LINHAB', 'LINHAC', 'LINHAD', 'LINHAII', 'CIRCOBITO', 'ACIDTRAB', 'FONTE', 'CODESTAB', 'ATESTANTE', 
                'UFINFORM', 'CONTADOR', 'HORAOBITO', 'CODBAIOCOR', 'TPASSINA', 'DTATESTADO', 'TPPOS', 'DTINVESTIG', 'CAUSABAS_O', 'DTCADASTRO', 'FONTEINV', 'DTRECEBIM', 
                'CB_PRE', 'ORIGEM', 'NUMERODN', 'MORTEPARTO', 'DTCADINF', 'TPOBITOCOR', 'DTCADINV', 'COMUNSVOIM', 'DTRECORIG', 'DTRECORIGA', 'CAUSAMAT', 'ESC2010', 
                'ESCMAE2010', 'STDOEPIDEM', 'STDONOVA', 'SEMAGESTAC', 'TPMORTEOCO', 'DIFDATA', 'DTCONCASO', 'NUDIASOBIN', 'SERIESCFAL', 'SERIESCMAE', 'CODMUNCART', 'CODCART', 
                'NUMREGCART', 'DTREGCART', 'DTCONINV', 'CODMUNNATU', 'ESTABDESCR', 'CRM', 'NUMEROLOTE', 'STCODIFICA', 'CODIFICADO', 'VERSAOSIST', 'VERSAOSCB', 'ATESTADO', 
                'ESCMAEAGR1', 'ESCFALAGR1', 'NUDIASOBCO', 'FONTES', 'TPRESGINFO', 'TPNIVELINV', 'NUDIASINF', 'FONTESINF', 'ALTCAUSA', 'NATURAL', 'IDADEminutos', 'IDADEhoras', 
                'IDADEdias', 'IDADEmeses', 'IDADEanos', 'OCUP', 'munResStatus', 'munResTipo', 'munResNome', 'munResUf', 'munResLat', 'munResLon', 'munResAlt', 'munResArea', 
                'OCUPMAE']


rename_columns = {
    "contador" : "ID",
    "TIPOBITO" : "tipo_obito",
    "DTOBITO" : "data_obito",
    "DTNASC" : "data_nascimento",
    "IDADE" : "idade",
    "SEXO" : "sexo",
    "RACACOR" : "raca_cor",
    "ESTCIV" : "estado_civil",
    "ESC" : "escolaridade",
    "CODBAIRES" : "codigo_bairro_residencia",
    "CODMUNRES" : "codigo_municipio_residencia",
    "LOCOCOR" : "local_ocorrencia_obito",
    "CODMUNOCOR" : "codigo_municipio_ocorrencia_obito",
    "QTDFILVIVO" : "quantidade_filho_nascidos_vivos",
    "QTDFILMORT" : "quantidade_filho_abortos",
    "GRAVIDEZ" : "tipo_gravidez",
    "GESTACAO" : "semanas_gestacao",
    "PARTO" : "tipo_parto",
    "OBITOPARTO" : "momento_obito_relacao_parto",
    "PESO" : "peso_ao_nascer",
    "OBITOGRAV" : "obito_na_gravidez",
    "OBITOPUERP" : "obito_no_puerperio",
    "ASSISTMED" : "assistencia_medica_durante_obito",
    "NECROPSIA" : "necropsia",
    "CAUSABAS" : "causa_da_morte",
    "" : "",
    "" : "",
    "" : "",
    "" : "",
    "" : "",

    
}