import datetime
import locale
import xlsxwriter as excel
locale.setlocale(locale.LC_ALL, 'pt_BR.UTF-8')


# def toExcel(lista):
#     """IMPLEMENTAR EXCEL DEPOIS DA COLETA DE DADOS"""
#     workbook  = excel.Workbook('primeirotestelesgo.xlsx')
#     workbook.formats[0].set_font_size(14)
#     worksheet = workbook.add_worksheet()
#     # lista = ['NOMEDALOJA',
#     #         [{'30/11/2002':{5:1, 10:1, 20:1, 50:1, 100:90000, 200:41}, 
#     #         '01/12/2009':{5:2, 10:4, 20:9, 50:1, 100:90, 200:4},
#     #         '02/12/2001':{5:2, 10:4, 20:9, 50:1, 100:90, 200:4}, 
#     #         '04/11/1989':{5:2, 10:4, 20:9, 50:1, 100:90, 200:4},
#     #         '04/11/1988':{5:2, 10:4, 20:9, 50:1, 100:90, 200:4},
#     #         '05/12/2002':{5:2, 10:4, 20:9, 50:1, 100:90, 200:4}}]
#     #         ]
#     datas = list(lista[1][0].keys())
#     print(datas)
#     datas.sort(key=lambda date: datetime.strptime(date, "%d/%m/%Y"))
#     print(datas)

#     #################################################################

#     nomeloja = workbook.add_format()
#     nomeloja.set_align('center')
#     nomeloja.set_align('vcenter')
#     worksheet.set_row(0, 35)
#     worksheet.set_column('A:A', 15)
#     worksheet.write(0, 0, 'NOME DA LOJA', nomeloja)

#     #################################################################
#     inicio = 3
#     notas = [5,10,20,50,100,200]
#     somatotal = 0
#     num = 5
#     width= 25
#     worksheet.set_column('A:E',width)
#     titulo_formatacao = workbook.add_format({'bold': True, 'font_color': 'red', 'align': 'center', 'font_size':16})
#     dados_formatacao = workbook.add_format({'align': 'center', 'font_size':14})
#     reais = workbook.add_format({'align': 'center', 'font_size':14, 'num_format':'R$ #,##0.00'})
#     worksheet.write('A1','NOME DA LOJA',dados_formatacao)

#     for i in datas:

#         worksheet.write_row(f'B{num}',[i,'VALOR(R$)', 'QTD', 'VALOR x QTD(R$)'], titulo_formatacao)
#         worksheet.write(f'B{num+1}', datetime.strptime(i, '%d/%m/%Y').strftime("%A").replace(' ','-').upper(), dados_formatacao) ## variavel dia da semana
#         worksheet.write_column(f'C{num+1}', notas, dados_formatacao) 
#         worksheet.write_column(f'D{num+1}', [lista[1][0][i][k] for k in notas], dados_formatacao)
#         worksheet.write_column(f'E{num+1}', [f'=PRODUCT(C{num+1+i}:D{num+1+i})' for i in range(len(notas))], dados_formatacao) ### formatacao de moeda
#         worksheet.write_row(f'C{num+7}', ['SOMA',f'=SUM(D{num+1}:D{num+6})',f'=SUM(E{num+1}:E{num+6})'], dados_formatacao)
#         num += 10
#     worksheet.write(f'A{inicio}', 'SOMA TOTAL',dados_formatacao)
#     worksheet.write(f'A{inicio+1}',f'=SUMIF(C{inicio}:C{num},"SOMA",E{inicio}:E{num})', reais)
#     workbook.close()
#     print('ravioli e bao')
#     return None
    
    
 
    




loja = input('\n\n\nInsira a loja: ')
workbook  = excel.Workbook(f'{loja}.xlsx')


"""CONVERTENDO DADOS DE EN => PT-BR..."""

lojas = ['São Bento - Matriz', 'Misericórdia', 'São Bento - 2']

mainstop = True


v_total = 0
biglistademalotes = [loja,[{}]] 

while mainstop:

    valida_data = False
    while not valida_data:
        try:
            dia, mes, ano = map(int, input('\n\nInsira a data: ').split())
            x = datetime.datetime.strptime(f'{dia}/{mes}/{ano}', '%d/%m/%Y')
            valida_data = True
        except:
            print('ERRO: INSIRA UMA DATA CORRETA!\n')


    date = x.strftime('%d/%m/%Y') #usar a data!!
    # print(x)
    # print(date)
    weekday = x.strftime("%A").replace(' ','-')  #usar dia da semana!
    dic = {5:0, 10:0, 20:0, 50:0, 100:0, 200:0}
    stop = False
    valormalote = 0







    while not stop:

        try:    
            a = int(input('\n\nQuantidade de notas de R$ 5.00:  '))
            dic[5] += a
        except:
            a = 0
            dic[5] += a
            
        try:
            b = int(input('Quantidade de notas de R$ 10.00:  '))
            dic[10] += b
        except:
            b = 0
            dic[10] += b
       
        try:
            c = int(input('Quantidade de notas de R$ 20.00: '))
            dic[20] += c
        except:
            c = 0
            dic[20] += c

        try: 
            d = int(input('Quantidade de notas de R$ 50.00: '))
            dic[50] += d
        except:
            d = 0
            dic[50] += d

        try:
            e = int(input('Quantidade de notas de R$ 100.00: '))
            dic[100] += e
        except:
            e = 0
            dic[100] += e

        try:
            f = int(input('Quantidade de notas de R$ 200.00: '))
            dic[200] += f
        except:
            f = 0
            dic[200] += f 


        suum = a*5+b*10+c*20+d*50+e*100+f*200
        print(f'--> Valor de caixa: R${suum:.2f}')
        v_total += suum
        valormalote += suum
        answ = input('\nDeseja continuar? (SIM -> 1 / NAO -> 2): ')
        print(dic)
        if answ != '2':
            print('\n')
            pass
        else:


            biglistademalotes[1][0][date] = dic
            # biglistademalotes.append({date:dic}) ##################



            print('\n')
            print('***'*25)
            print(f'Notas do malote: {dic}')
            stop = True
    print(f'*MALOTE DO DIA {date}* = R${valormalote:.2f}\n')
    print(f'*VALOR ACUM. DE TODAS AS DATAS (ATÉ AGORA)* = R${v_total:.2f}')
    print('***'*25)
    print('\n')

    excel = input('ENVIAR PARA EXCEL? (S/N):').lower()
    print(biglistademalotes)
    if excel == 's':
        lista = biglistademalotes







        ###############################################################################





        
        # workbook  = excel.Workbook('primeirotestelesgo.xlsx')
        workbook.formats[0].set_font_size(14)
        worksheet = workbook.add_worksheet()
        # lista = ['NOMEDALOJA',
        #         [{'30/11/2002':{5:1, 10:1, 20:1, 50:1, 100:90000, 200:41}, 
        #         '01/12/2009':{5:2, 10:4, 20:9, 50:1, 100:90, 200:4},
        #         '02/12/2001':{5:2, 10:4, 20:9, 50:1, 100:90, 200:4}, 
        #         '04/11/1989':{5:2, 10:4, 20:9, 50:1, 100:90, 200:4},
        #         '04/11/1988':{5:2, 10:4, 20:9, 50:1, 100:90, 200:4},
        #         '05/12/2002':{5:2, 10:4, 20:9, 50:1, 100:90, 200:4},
        #         '05/12/2003':{5:2, 10:4, 20:9, 50:1, 100:90, 200:4},
        #         '05/12/2004':{5:2, 10:4, 20:9, 50:1, 100:90, 200:4},
        #         '05/12/2005':{5:2, 10:4, 20:9, 50:1, 100:90, 200:4},
        #         '05/12/2006':{5:2, 10:4, 20:9, 50:1, 100:90, 200:4},
        #         '05/12/2007':{5:2, 10:4, 20:9, 50:1, 100:90, 200:4},
        #         '05/12/2008':{5:2, 10:4, 20:9, 50:1, 100:90, 200:4},
        #         '05/12/2009':{5:2, 10:4, 20:9, 50:1, 100:90, 200:4},
        #         '05/12/2010':{5:2, 10:4, 20:9, 50:1, 100:90, 200:4}}]
        #         ]
        datas = list(lista[1][0].keys())
        print(datas)
        datas.sort(key=lambda date: datetime.datetime.strptime(date, "%d/%m/%Y"))
        print(datas)

        #################################################################

         
        # worksheet.set_row(0, 35)
        # worksheet.set_column('A:A', 15)
        

        #################################################################
        inicio = 1
        notas = [5,10,20,50,100,200]
        somatotal = 0
        num = 3
        width= 15
        worksheet.set_column('A:E',width)
        titulo_formatacao = workbook.add_format({'border':2,'shrink': True,'bold': True,'valign':'vcenter', 'align': 'center', 'font_size':14})
        dados_formatacao = workbook.add_format({'border':1,'align': 'center','valign':'vcenter', 'font_size':12, 'shrink': True, 'num_format':'#,###0'})
        dados_formatacao_SOMA = workbook.add_format({'border':2,'align': 'center','valign':'vcenter','font_size':12, 'bold': True, 'shrink': True, 'num_format':'#,###0'})
        dados_formatacao_WEEKDAY = workbook.add_format({'border':2,'shrink': True,'align': 'center','valign':'vcenter','font_size':12, 'italic':True})
        dados_formatacao_REAIS = workbook.add_format({'border':2,'align': 'center','valign':'vcenter' ,'font_size':12, 'num_format':'R$ #,##0.00','bold': True, 'italic': True, 'shrink': True})
        nomedaloja_formatacao = workbook.add_format({'border':4,'align': 'center','valign':'vcenter', 'font_size':14, 'bold': True, 'italic': True, 'shrink':True})
        dados_formatacao_REAIS_SOMATOTAL = workbook.add_format({'border':6,'border_color':'#000000','align': 'center','valign':'vcenter' ,'font_size':12, 'num_format':'R$ #,##0.00','bold': True, 'italic': True, 'shrink': True})
        dados_formatacao_SOMATOTAL = workbook.add_format({'border':6,'border_color':'#000000','align': 'center','valign':'vcenter', 'font_size':12, 'shrink': True, 'num_format':'#,###0','bold': True, 'italic': True})
        worksheet.write('A1',lista[0],nomedaloja_formatacao)

        for i in datas:

            worksheet.write_row(f'A{num}',[i,'VALOR(R$)', 'QTD', 'VALOR x QTD(R$)'], titulo_formatacao)
            worksheet.write(f'A{num+1}', datetime.datetime.strptime(i, '%d/%m/%Y').strftime("%A").replace(' ','-').upper(), dados_formatacao_WEEKDAY) ## variavel dia da semana
            worksheet.write_column(f'B{num+1}', notas, dados_formatacao) 
            worksheet.write_column(f'C{num+1}', [lista[1][0][i][k] for k in notas], dados_formatacao)
            worksheet.write_column(f'D{num+1}', [f'=PRODUCT(B{num+1+i}:C{num+1+i})' for i in range(len(notas))], dados_formatacao) ### formatacao de moeda
            worksheet.write_row(f'B{num+7}', ['SOMA',f'=SUM(C{num+1}:C{num+6})',f'=SUM(D{num+1}:D{num+6})'], dados_formatacao_SOMA)
            worksheet.write(f'D{num+7}', f'=SUM(D{num+1}:D{num+6})', dados_formatacao_REAIS)
            num += 9
        

        
         
        worksheet.write(f'C{inicio}', 'SOMA TOTAL',dados_formatacao_SOMATOTAL)
        worksheet.write(f'D{inicio}',f'=SUMIF(B{inicio}:B{num},"SOMA",D{inicio}:D{num})', dados_formatacao_REAIS_SOMATOTAL)
        
        workbook.close()
        print('ravioli e bao')

        
        












        print('EXCEL FINALIZADO COM SUCESSO!')
        mainstop = False
    else:
        pass


    




