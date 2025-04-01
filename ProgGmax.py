import pandas as pd
import numpy as np
from datetime import datetime
from modulos import Navegador, Navegador_Correlatos, Interface, TableWithExport
import sys
import selenium.common.exceptions as exc
from PyQt5.QtWidgets import QApplication

janela = Interface()
valores_login = None
tipo_Janela = janela.selecionar_Tipo()

if tipo_Janela == 1:

    app = QApplication([])
    headers = ['NS*', 'NOM_CLIENTE*', 'DAT_COMPROMISSO', 'MERCADO*', 'SERVICO_EXP*', 'COD_LOCAL*', 'TELEFONE',
               'CELULAR', 'ABSCISSA', 'ORDENADA', 'CONTRATO*', 'COD_SERVICO*', 'NUM_MED*', 'BASE*', 'DAT_RECEBIMENTO',
               'PRAZO']
    window = TableWithExport(rows=1, cols=16, headers=headers, arq="cadastro_servico.csv")
    window.show()
    app.exec_()

    cadastro_servico = pd.read_csv("cadastro_servico.csv", sep=',')

    cadastro_servico['DAT_RECEBIMENTO'] = cadastro_servico['DAT_RECEBIMENTO'].replace(np.nan, "01/01/1900")
    cadastro_servico['DAT_COMPROMISSO'] = cadastro_servico['DAT_COMPROMISSO'].replace(np.nan, "01/01/1900")
    cadastro_servico['NOM_CLIENTE*'] = cadastro_servico['NOM_CLIENTE*'].replace(np.nan, "CLIENTE NÃO IDENTIFICADO")
    cadastro_servico['SERVICO_EXP*'] = cadastro_servico['SERVICO_EXP*'].replace(np.nan, "NID")
    cadastro_servico['COD_SERVICO*'] = cadastro_servico['COD_SERVICO*'].apply(lambda x: '{0:0>4}'.format(x))
    cadastro_servico['COD_LOCAL*'] = cadastro_servico['COD_LOCAL*'].replace(np.nan, "EX-3101")
    cadastro_servico['CONTRATO*'] = cadastro_servico['CONTRATO*'].replace(np.nan, 4680006397)
    cadastro_servico['CONTRATO*'] = np.array(cadastro_servico['CONTRATO*'], dtype=np.int64)
    cadastro_servico['TELEFONE'] = cadastro_servico['TELEFONE'].replace(np.nan, "")
    cadastro_servico['CELULAR'] = cadastro_servico['CELULAR'].replace(np.nan, "")
    cadastro_servico['ABSCISSA'] = cadastro_servico['ABSCISSA'].replace(np.nan, "")
    cadastro_servico['ORDENADA'] = cadastro_servico['ORDENADA'].replace(np.nan, "")
    cadastro_servico['PRAZO'] = cadastro_servico['PRAZO'].replace(np.nan, "")

    if janela.tela_login() == True:
        valores_login = janela.valor_login
        navegador = Navegador()
        if navegador.loginGserv(valores_login) == True:
            arq = open("LOG_CADASTRO_SERV.txt", 'a')
            arq.write("\n***ARQUIVO DE LOGS DA EXECUÇÃO DO CADASTRO DE SERVIÇOS {}***\n\n".format(datetime.today()))
            for index, row in cadastro_servico.iterrows():
                teste = navegador.buscarNS(row['NS*'])
                if teste == True:
                    navegador.cadastrarServ(contrato=row['CONTRATO*'], cod_servico=row['COD_SERVICO*'],
                                            num_med=row['NUM_MED*'], dat_recebimento=row['DAT_RECEBIMENTO'],
                                            base=row['BASE*'], prazo=row['PRAZO'])
                    arq.write("CADASTRADO: NS: {}, COD_SERV: {}, NUM_MED: {}.\n".format(row['NS*'], row['COD_SERVICO*'],
                                                                                        row['NUM_MED*']))
                else:
                    navegador.cadastrarNS(nom_cliente=row['NOM_CLIENTE*'], serv_exp=row['SERVICO_EXP*'],
                                          mercado=row['MERCADO*'], cod_local=row['COD_LOCAL*'],
                                          dat_compromisso=str(row['DAT_COMPROMISSO']), telefone=str(row['TELEFONE']),
                                          celular=str(row['CELULAR']), abscissa=str(row['ABSCISSA']),
                                          ordenada=str(row['ORDENADA']))
                    navegador.cadastrarServ(contrato=row['CONTRATO*'], cod_servico=row['COD_SERVICO*'],
                                            num_med=row['NUM_MED*'], dat_recebimento=row['DAT_RECEBIMENTO'],
                                            base=row['BASE*'], prazo=row['PRAZO'])
                    arq.write("CADASTRADO: NS: {}, COD_SERV: {}, NUM_MED: {}.\n".format(row['NS*'], row['COD_SERVICO*'],
                                                                                        row['NUM_MED*']))
            arq.write("\n***RELAÇÃO PROGRAMADA COM SUCESSO***\n")
            arq.close()

elif tipo_Janela == 2:

    app = QApplication([])
    headers = ['NS*', 'SERV*', 'MED_SAP*', 'COD_ACAO*', 'MATRICULA_RESP', 'DAT_CONCLUSAO', 'PROXIMA_ACAO', 'OBS']
    window = TableWithExport(rows=1, cols=8, headers=headers, arq="cadastro_acao.csv")
    window.show()
    app.exec_()

    cadastro_acao = pd.read_csv("cadastro_acao.csv", sep=',')
    cadastro_acao['DAT_CONCLUSAO'] = cadastro_acao['DAT_CONCLUSAO'].replace(np.nan, "")
    cadastro_acao['PROXIMA_ACAO'] = cadastro_acao['PROXIMA_ACAO'].replace(np.nan, 0)
    cadastro_acao['PROXIMA_ACAO'] = np.array(cadastro_acao['PROXIMA_ACAO'], dtype=np.int64)
    cadastro_acao['SERV*'] = cadastro_acao['SERV*'].apply(lambda x: '{0:0>4}'.format(x))
    cadastro_acao['COD_ACAO*'] = cadastro_acao['COD_ACAO*'].apply(lambda x: '{0:0>3}'.format(x))
    cadastro_acao['PROXIMA_ACAO'] = cadastro_acao['PROXIMA_ACAO'].apply(lambda x: '{0:0>3}'.format(x))
    cadastro_acao['MATRICULA_RESP'] = cadastro_acao['MATRICULA_RESP'].replace(np.nan, "")
    cadastro_acao['OBS'] = cadastro_acao['OBS'].replace(np.nan, "")

    if janela.tela_login() == True:
        valores_login = janela.valor_login
        navegador = Navegador()
        if navegador.loginGserv(valores_login) == True:
            arq = open('LOG_CADASTRO_ACAO.txt', 'a')
            arq.write("\n***ARQUIVO DE LOGS DA EXECUÇÃO DO CADASTRO DE AÇÕES {}***\n\n".format(datetime.today()))
            for index, row in cadastro_acao.iterrows():
                navegador.buscarNS(row['NS*'])
                index_serv = navegador.buscarServ(serv=str(row['SERV*']), med=str(row['MED_SAP*']))
                if index_serv != 0:
                    navegador.cadastrarAcoes(serv_index=index_serv, cod_acao=str(row['COD_ACAO*']),
                                             matricula_resp=str(row['MATRICULA_RESP']),
                                             dat_conclusao=str(row['DAT_CONCLUSAO']),
                                             prox_acao=str(row['PROXIMA_ACAO']), obs=str(row['OBS']))
                    arq.write(
                        "CADASTRADO: NS {}, SERV: {}, nº: {}, ACÃO: {}, RESP: {}.\n".format(row['NS*'], row['SERV*'],
                                                                                            row['MED_SAP*'],
                                                                                            row['COD_ACAO*'],
                                                                                            row['MATRICULA_RESP']))
                else:
                    arq.write("\n***AO BUSCAR SERVIÇO***\n\nO serviço {}, nº {}, não existe na NS {}.".format(
                        str(row['SERV*']), str(row['MED_SAP*']), str(row['NS*'])))
                    arq.close()
                    janela.popUp(text="O serviço {}, nº {}, não existe na NS {}. Encerrando programação!".format(
                        str(row['SERV*']), str(row['MED_SAP*']), str(row['NS*'])))
                    sys.exit()
            arq.write("\n***RELAÇÃO CADASTRADA COM SUCESSO***\n")
            arq.close()

elif tipo_Janela == 3:

    app = QApplication([])
    headers = ['NS*', 'SERV*', 'NUM_MED*', 'ACAO*', 'MATRICULA_RESP*']
    window = TableWithExport(rows=1, cols=5, headers=headers, arq="trocar_resp.csv")
    window.show()
    app.exec_()

    trocaResp = pd.read_csv("trocar_resp.csv", sep=',')
    trocaResp['SERV*'] = trocaResp['SERV*'].apply(lambda x: '{0:0>4}'.format(x))
    trocaResp['ACAO*'] = trocaResp['ACAO*'].apply(lambda x: '{0:0>3}'.format(x))

    if janela.tela_login() == True:
        valores_login = janela.valor_login
        navegador = Navegador()
        if navegador.loginGserv(valores_login) == True:
            arq = open('LOG_TROCA_RESP.txt', 'a')
            arq.write("\n***ARQUIVO DE LOGS DA EXECUÇÃO DA TROCA DE RESPONSÁVEIS {}***\n\n".format(datetime.today()))
            for index, row in trocaResp.iterrows():
                navegador.buscarNS(row['NS*'])
                serv_index = navegador.buscarServ(serv=row['SERV*'], med=str(row['NUM_MED*']))
                if serv_index != 0:
                    acao_index = navegador.buscarAcao(serv_index=serv_index, acao=row['ACAO*'])
                    if acao_index != 0:
                        navegador.trocarRespAcao(acao_index, row['MATRICULA_RESP*'])
                        arq.write("CADASTRADA: NS: {}, MATRICULA: {}, SERV: {}, ACAO: {}\n".format(row['NS*'], row[
                            'MATRICULA_RESP*'], row['SERV*'], row['ACAO*']))
                    else:
                        arq.write("\n***ERRO AO BUSCAR AÇÃO***\n\nA ação {}, não existe no serviço {} da NS {}".format(
                            str(row['ACAO*']), str(row['SERV*']), str(row['NS*'])))
                        arq.close()
                        sys.exit()
                else:
                    janela.popUp(
                        "O serviço {}, nº {}, não existe na NS {}. Encerrando programação!".format(str(row['SERV*']),
                                                                                                   str(row['NUM_MED*']),
                                                                                                   str(row['NS*'])))
                    arq.write("\n***ERRO AO BUSCAR SERVIÇO***\n\nO serviço {}, nº {}, não existe na NS {}".format(
                        str(row['SERV*']), str(row['NUM_MED*']), str(row['NS*'])))
                    arq.close()
                    sys.exit()
            arq.write("\n***RELAÇÃO ALTERADA COM SUCESSO***\n")
            arq.close()

elif tipo_Janela == 4:

    app = QApplication([])
    headers = ['NS*', 'SERV*', 'NUM_MED*', 'RECEBIMENTO_REAL*']
    window = TableWithExport(rows=1, cols=4, headers=headers, arq="trocar_data_receb.csv")
    window.show()
    app.exec_()

    trocar_data_receb = pd.read_csv("trocar_data_receb.csv", sep=',')
    trocar_data_receb['SERV*'] = trocar_data_receb['SERV*'].apply(lambda x: '{0:0>4}'.format(x))

    if janela.tela_login() == True:
        valores_login = janela.valor_login
        navegador = Navegador()
        if navegador.loginGserv(valores_login) == True:
            arq = open('LOG_TROCA_DAT_RECEB.txt', 'a')
            arq.write("\n***ARQUIVO DE LOGS DA TROCA DE DATA DE RECEBIMENTO {}***\n\n".format(datetime.today()))
            for index, row in trocar_data_receb.iterrows():
                navegador.buscarNS(row['NS*'])
                index_serv = navegador.buscarServ(serv=str(row['SERV*']), med=str(row['NUM_MED*']))
                if index_serv > 0:
                    navegador.trocarData(dat_recebimento=row['RECEBIMENTO_REAL*'], serv_index=index_serv)
                    arq.write(
                        f"DATA DE RECEBIMENTO ALTERADA, NS: {row['NS*']}, SERV: {row['SERV*']}, NUM_MED: {row['NUM_MED*']}\n")
                else:
                    arq.write(
                        f"ERRO AO CADASTRAR SERVICO, NS: {row['NS*']}, SERV: {row['SERV*']}, NUM_MED: {row['NUM_MED*']}\n")
                    janela.popUp(
                        f"ERRO AO CADASTRAR SERVICO, NS: {row['NS*']}, SERV: {row['SERV*']}, NUM_MED: {row['NUM_MED*']}")
                    arq.close()
                    sys.exit()
            arq.write("\n***RELAÇÃO ALTERADA COM SUCESSO***\n")
            arq.close()

elif tipo_Janela == 5:

    app = QApplication([])
    headers = ['NS*', 'NOME_CLIENTE*', 'COD_LOCAL*', 'CONTRATO*', 'COD_SERV*', 'COD_ACAO*', 'BASE*', 'MATRICULA_RESP*',
               'DAT_EXEC*', 'OBS', 'REQ*', 'ITEM*', 'QTD_US*']
    window = TableWithExport(rows=1, cols=len(headers), arq="cadastro_correlatos.csv", headers=headers)
    window.show()
    app.exec_()
    cadastro_correlatos = pd.read_csv("cadastro_correlatos.csv", sep=',')
    cadastro_correlatos['OBS'] = cadastro_correlatos['OBS'].replace(np.nan, "")
    cadastro_correlatos['COD_SERV*'] = cadastro_correlatos['COD_SERV*'].apply(lambda x: '{0:0>4}'.format(x))
    cadastro_correlatos['COD_ACAO*'] = cadastro_correlatos['COD_ACAO*'].apply(lambda x: '{0:0>3}'.format(x))

    if janela.tela_login() == True:
        valores_login = janela.valor_login
        navegador = Navegador_Correlatos()
        if navegador.loginGserv(valores_login) == True:
            arq = open('LOG_CADASTRO_CORRELATOS.txt', 'a')
            arq.write(f"\n*** ARQUIVO DE LOGS DO CADASTRO DE CORRELATOS {datetime.today()} ***\n\n")
            for index, row in cadastro_correlatos.iterrows():
                teste = navegador.buscarNS(row['NS*'])
                if teste == True:
                    navegador.cadastrarCorrelato(contrato=row['CONTRATO*'], cod_serv=row['COD_SERV*'],
                                                 cod_acao=row['COD_ACAO*'], base=row['BASE*'],
                                                 matricula_resp=row['MATRICULA_RESP*'], dat_exec=row['DAT_EXEC*'],
                                                 obs=row['OBS'])
                    navegador.cadastroReq(req=str(row['REQ*']), qtd_us=str(row['QTD_US*']), item=str(row['ITEM*']),
                                          cod_acao=row['COD_ACAO*'], ns=row['NS*'])
                    arq.write(
                        f"CADASTRADA, NS: {row['NS*']}, RESP: {row['MATRICULA_RESP*']}, DAT_EXEC {row['DAT_EXEC*']}, ACAO {row['COD_ACAO*']}\n")
                else:
                    navegador.cadastrarNS(nome=row['NOME_CLIENTE*'], cod_local=row['COD_LOCAL*'])
                    navegador.cadastrarCorrelato(contrato=row['CONTRATO*'], cod_serv=row['COD_SERV*'],
                                                 cod_acao=row['COD_ACAO*'], base=row['BASE*'],
                                                 matricula_resp=row['MATRICULA_RESP*'], dat_exec=row['DAT_EXEC*'],
                                                 obs=row['OBS'])
                    navegador.cadastroReq(req=str(row['REQ*']), qtd_us=str(row['QTD_US*']), item=str(row['ITEM*']),
                                          cod_acao=row['COD_ACAO*'], ns=row['NS*'])
                    arq.write(
                        f"CADASTRADA, NS: {row['NS*']}, RESP: {row['MATRICULA_RESP*']}, DAT_EXEC {row['DAT_EXEC*']}, ACAO {row['COD_ACAO*']}\n")
            arq.write("\n*** RELAÇÃO CADASTRADA COM SUCESSO ***")

elif tipo_Janela == 6:

    app = QApplication([])
    headers = ['NS*', 'COD_SERV*', 'NUM_MED*', 'COD_ACAO*', 'MATRICULA_RESP*', 'DAT_INI', 'DAT_CONC*', 'US_PRJ',
               'US_TOP', 'US_GEO', 'US_INT', 'PROX_ACAO', 'AC_OBS', 'REQ_NUM', 'ITEM', 'QTD_US']
    window = TableWithExport(rows=1, cols=len(headers), arq="concluir_acao.csv", headers=headers)
    window.show()
    app.exec_()
    concluir_acao = pd.read_csv("concluir_acao.csv", sep=',')
    concluir_acao['COD_SERV*'] = concluir_acao['COD_SERV*'].apply(lambda x: '{0:0>4}'.format(x))
    concluir_acao['COD_ACAO*'] = concluir_acao['COD_ACAO*'].apply(lambda x: '{0:0>3}'.format(x))
    concluir_acao['PROX_ACAO'] = concluir_acao['PROX_ACAO'].apply(lambda x: '{0:0>3}'.format(x))
    concluir_acao['PROX_ACAO'] = concluir_acao['PROX_ACAO'].replace('nan', '')
    concluir_acao['DAT_INI'] = concluir_acao['DAT_INI'].replace(np.nan, '')
    concluir_acao['US_PRJ'] = concluir_acao['US_PRJ'].replace(np.nan, '')
    concluir_acao['US_TOP'] = concluir_acao['US_TOP'].replace(np.nan, '')
    concluir_acao['US_GEO'] = concluir_acao['US_GEO'].replace(np.nan, '')
    concluir_acao['US_INT'] = concluir_acao['US_INT'].replace(np.nan, '')
    concluir_acao['AC_OBS'] = concluir_acao['AC_OBS'].replace(np.nan, '')
    concluir_acao['REQ_NUM'] = concluir_acao['REQ_NUM'].replace(np.nan, '')
    for index, row in concluir_acao.iterrows():
        if row['REQ_NUM'] != '':
            concluir_acao['REQ_NUM'] = concluir_acao['REQ_NUM'].astype('int64')
            break
    concluir_acao['ITEM'] = concluir_acao['ITEM'].replace(np.nan, '')
    concluir_acao['QTD_US'] = concluir_acao['QTD_US'].replace(np.nan, '')

    if janela.tela_login() == True:
        valores_login = janela.valor_login
        navegador = Navegador()
        if navegador.loginGserv(valores_login) == True:
            arq = open('LOG_FECHAMENTO_ACOES.txt', 'a')
            arq.write(f"*** LOG DE FECHAMENTO DE ACOES {datetime.today()} ***\n\n")
            for index, row in concluir_acao.iterrows():
                navegador.buscarNS(row['NS*'])
                serv_index = navegador.buscarServ(row['COD_SERV*'], str(row['NUM_MED*']))
                if serv_index != 0:
                    acao_index = navegador.buscarAcao(serv_index, str(row['COD_ACAO*']))
                    if acao_index != 0:
                        navegador.concluir_acao(str(acao_index), row['MATRICULA_RESP*'], row['DAT_INI'],
                                                row['DAT_CONC*'], row['US_PRJ'], row['US_TOP'], row['US_GEO'],
                                                row['US_INT'], row['PROX_ACAO'], row['AC_OBS'])
                        if row['REQ_NUM'] != '' and row['ITEM'] != '' and row['QTD_US'] != '':
                            acao_index = navegador.buscarAcao(serv_index, str(row['COD_ACAO*']))
                            navegador.lancarRequisicao(acao_index, str(row['REQ_NUM']), str(row['ITEM']),
                                                       str(row['QTD_US']))
                        arq.write(
                            f"ACAO FECHADA: NS {row['NS*']}, SERV: {row['COD_SERV*']} Nº {row['NUM_MED*']}, ACAO: {row['COD_ACAO*']}, RESP: {row['MATRICULA_RESP*']}, REQUISICAO: {row['REQ_NUM']}, ITEM: {row['ITEM']}, QTD_US_REQ: {row['QTD_US']}\n")
            arq.write("\n*** RELACAO CADASTRADA COM SUCESSO ***")
            arq.close()

elif tipo_Janela == 7:
    app = QApplication([])
    headers = ['NS*', 'COD_SERV*', 'NUM_MED*', 'COD_ACAO*', 'TEXT_OBS']
    window = TableWithExport(rows=1, cols=len(headers), arq="trocar_obs.csv", headers=headers)
    window.show()
    app.exec_()
    troca_obs = pd.read_csv("trocar_obs.csv", sep=',')

    troca_obs['TEXT_OBS'] = troca_obs['TEXT_OBS'].replace(np.nan, '')
    troca_obs['COD_SERV*'] = troca_obs['COD_SERV*'].apply(lambda x: '{0:0>4}'.format(x))
    troca_obs['COD_ACAO*'] = troca_obs['COD_ACAO*'].apply(lambda x: '{0:0>3}'.format(x))

    if janela.tela_login() == True:
        valores_login = janela.valor_login
        navegador = Navegador()
        if navegador.loginGserv(valores_login) == True:
            arq = open('LOG_TROCA_OBS.txt', 'a')
            arq.write(f"*** LOG DE TROCA DE OBSERVAÇÃO DA AÇÃO {datetime.today()}***\n\n")
            for index, row in troca_obs.iterrows():
                navegador.buscarNS(row['NS*'])
                serv_index = navegador.buscarServ(row['COD_SERV*'], str(row['NUM_MED*']))
                if serv_index != 0:
                    acao_index = navegador.buscarAcao(serv_index, str(row['COD_ACAO*']))
                    if acao_index != 0:
                        navegador.trocarObs(acao_index, str(row['TEXT_OBS']))
                        arq.write(
                            f"OBSERVAÇÃO ALTERADA: NS {row['NS*']}, SERV: {row['COD_SERV*']}, AÇÃO: {row['COD_ACAO*']}\n")
            arq.write("\n*** RELACAO CADASTRADA COM SUCESSO ***")
            arq.close()

else:
    janela.popUp(text="Saindo da programação.")