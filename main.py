from datetime import datetime

import email_assistente as ea

try:
    import os
    import traceback
    from time import sleep

    import cv2
    import googlencry.demo as gg
    import keyboard as kb
    import pyautogui as pa
    import speech_recognition as sr
    from wmi import WMI

    import getaudiostatus as gas
    import voz
    from alarme_lim import lim_cl
    from close import closealarm


    def brilho(bright):
        c = WMI(namespace='wmi')
        methods = c.WmiMonitorBrightnessMethods()[0]    
        methods.WmiSetBrightness(bright, 0)

    def alarme(hora, minuto, falar=True):
        # VERSAO ANTIGA
        # os.startfile(R'C:\Users\pedro\AppData\Local\Programs\Python\Python310\Lib\site-packages\calarme\alarmevisual.exe')
        # sleep(15)
        # pa.write(f'{hora}')
        # pa.press('tab')
        # pa.write(f'{min}')
        # pa.press('tab')
        # pa.press('space')
        # # kb.press_and_release('alt + esc')

        # NOVA VERSAO 09/01/2023
        import writecode as wc
        wc.arquivo(hora, minuto)
        minuto=int(minuto)
        sleep(0.5)
        if int(hora) < 10:
            hora=f'0{hora}'
        if int(minuto) < 10:
            minuto=f'0{minuto}'
        os.popen(f'C:/Users/pedro/Desktop/nicory/alarmes/alarm{hora}{minuto}.pyw')

        if falar:
            voz.falar(f'Alarme definido para {hora}:{minuto}.') 

    def timer(tempo):
        os.startfile(R'C:\Users\pedro\AppData\Local\Programs\Python\Python310\Lib\site-packages\calarme\timer.py')
        sleep(1) 
        pa.write(f'{tempo}')
        pa.press('enter')
        kb.press_and_release('alt + esc')
        voz.falar('Timer definido.') 


    def reiniciar():
        os.startfile(R'C:\Users\pedro\Desktop\nicory\lua\restartassis.pyw')
        quit()

    dias ={
        'Monday':'Segunda-Feira',
        'Tuesday':'Terça-Feira',
        'Wednesday':'Quarta-Feira',
        'Thursday':'Quinta-Feira',
        'Friday':'Sexta-Feira',
        'Saturday':'Sábado',
        'Sunday':'Domingo'
        }
    dias_key = {
        'Segunda-Feira':0,
        'Terça-Feira':1,
        'Quarta-Feira':2,
        'Quinta-Feira':3,
        'Sexta-Feira':4,
        'Sábado':5,
        'Domingo':6
    }
    key_dias = dict((v, k) for k, v in dias_key.items())

    contador_reiniciar_ = 0
    rec = sr.Recognizer()

    ERRO = 0
    
    

    #-------------------------- PROGRAMA PRINCIPAL --------------------------
    while True:

        try: 
            if gas.mediaName() == 'Advertisement':
                # print(gas.mediaName())
                os.popen(R'C:\Users\pedro\Desktop\nicory\lua\adjump.pyw')
        except:
            pass

# Abrir microfone para detecção de voz
        timeee = datetime.now().time()

        if contador_reiniciar_ == 100:
            reiniciar()

        with sr.Microphone(1) as mic:

            cap = cv2.VideoCapture(0)

            rec.adjust_for_ambient_noise(mic)

            print(f'fale ({timeee})')

            _, frame = cap.read()

            audio = rec.listen(mic,phrase_time_limit=7)

            cap.release()

            try:
                texto = rec.recognize_google(audio, language='pt-BR')
                print(texto)    
                contador_reiniciar_ = 0

            except:

                ########
                # testar comandos
                # texto = 'lua ativar alarme 12:00'
                ########

                # print('nao entendi')
                contador_reiniciar_ += 1
                print(contador_reiniciar_)
                continue
        
# Comando para desligar lua
        if 'desligar' in texto.lower():
            voz.falar('Desligando...')
            quit()

        # comando fniciar
        elif 'reiniciar' in texto.lower():
            voz.falar('Reiniciando.')
            reiniciar()
# COMANDo para desligar alarme/timer
        elif 'acordei' in texto.lower(): # ===============
            kb.press_and_release
# Condição para iniciar comandos:

        # varios nomes para caso haja confusao no MicRecognizer - nome escolhidos apos testes com erros
        elif 'lua' in texto.lower().split() or 'lu' in texto.lower().split() or 'luan' in texto.lower().split():

            comand = texto.lower().split().copy()
            if 'lua' in comand:
                comand.remove('lua')
            elif 'lu' in comand:
                comand.remove('lu')
            elif 'luan' in comand:
                comand.remove('luan')

            # bom dia
            if 'bom' in comand and 'dia' in comand:
                brilho(80)
                comand.remove('bom')
                comand.remove('dia')
                voz.falar('Bom dia pedro.')
                sts = gas.mediaIs('PAUSED')
                if sts:
                    kb.press('play/pause media')


            # comando vazio
            if len(comand) == 0:
                continue
            
            # comandoo para cancelar comando lua
            elif 'cancela' in comand:
                voz.falar('ok')
                continue

            # mozao
            elif 'liara' in comand or 'mozão' in comand:
                voz.falar('Liara, meu senhor tem uma mensagem para você. Meu bem, amo-te demais, sabias disso ?')

            # comandoo que horas sao
            elif 'horas' and 'são' in comand:
                voz.horas()

            # comandos midia
            elif 'música' in comand:
                for n in comand:

                    if n == 'muda' or n == 'próxima' or n == 'passa':
                        kb.press('next track')

                    elif n == 'volta':
                        kb.press('previous track')
                    elif n == 'anterior':
                        kb.press('previous track')
                        sleep(0.5)
                        kb.press('previous track')

                    elif n == 'para' or n == 'pausa' or n == 'play' or n == 'toca':
                        kb.press('play/pause media')
                    
                    else:
                        pass

            # comandoo ativar alarme 
            elif 'ativar' in comand or 'criar' in comand or 'definir' in comand: 
                # print(comand)
                if 'alarme' in comand:

                    lim = lim_cl()
                    # print(lim)
                    if lim >= 6:
                        voz.falar('O limite de alarmes foi atingido!')
                    else:
                        try:
                            if 'para' in comand:
                                comand.remove('para')
                            if 'de' in comand:
                                comand.remove('de')
                            if 'definir' in comand:
                                comand.remove('definir')
                            if 'ativar' in comand:
                                comand.remove('ativar')
                            if 'alarme' in comand:
                                comand.remove('alarme')
                            if 'as' in comand:
                                comand.remove('as')

                            cfrase = comand
                            if not 'e' in cfrase:
                                try:
                                    hora, min = cfrase[0].split(':')
                                except:
                                    horas = comand.index('horas')
                                    hora = cfrase[0]
                                    min = 0
                                alarme(hora, min) 
                                print(comand)
                            else:
                                try:
                                    e = cfrase.index('e')
                                    hora, min = cfrase[e - 1], cfrase[e + 1]
                                    min = int(min)
                                    hora = int(hora)
                                    if min == 6:
                                        min = 30
                                    alarme(hora, min)
                                except Exception as e:
                                    print(e)
                                    voz.falar('Erro 2')
                                    print('foi aqui')
                                    print(cfrase)
                        except Exception as e:
                            raise e
                
                #comando timer
                elif 'timer' in comand:
                    if 'ativar' in comand:
                        comand.remove('ativar')
                    if 'de' in comand:
                        comand.remove('de')
                    if 'para' in comand:
                        comand.remove('para')
                    
                    tempo = int(comand[1])
                    
                    if 'segundos' in comand[2]:
                        pass
                    elif 'minutos' in comand[2]:
                        tempo *= 60
                    elif 'horas' in comand[2]:
                        tempo *= 3600

                    timer(tempo)
                
                else:
                    pass


            elif 'desativar' in comand:
                if 'alarme' in comand:
                    comand.remove('desativar')
                    comand.remove('alarme')
                    if 'às' in comand:
                        comand = comand.replace('às', 'de')
                    comand.remove('de')
                    
                    if 'horas' in comand:
                        h = comand[0]
                        m = 0
                    else:
                        h, m = comand[0].split(':')

                    try:
                        if int(h) < 10:
                            h=f'0{h}'
                        if int(m) < 10:
                            m=f'0{m}'
                        if closealarm(f'alarm{h}{m}') == 1:
                            voz.falar(f'Alarme de {h}:{m} desativado')
                        else:
                            voz.falar(f'Alarme de {h}:{m} não encontrado')

                    except:
                        voz.falar('Erro na desativação')
                elif 'todos' in comand and 'alarmes' in comand:
                    import listalrm as lis
                    lis.desativarall()

            # comando listar eventos calendario

            elif 'agenda' in comand or 'calendário' in comand: 
                
                print(comand)
                if 'mostrar' in comand or 'mostra' in comand:
                    try:
                        calendar = [[],[],[],[],[],[],[]]
                        for agenda in gg.listarEventosSemanaGERAL():
                            
                            data = agenda['data'].split('/')
                            dt = datetime.date(datetime(int(data[2]),int(data[1]),int(data[0])))

                            
                            dia_semana = dias[f'{dt.strftime("%A")}']

                            if len(agenda) == 4:
                                # print(agenda['nome'], agenda['horario'], agenda['data'])
                                ag = agenda['nome'], agenda['horario'], agenda['data']
                            else:
                                # print(agenda['nome'], agenda['data'])
                                ag = agenda['nome'], agenda['data']
                                
                            calendar[dias_key[f'{dia_semana}']].append(ag)
                            
                            # print('day Name:', dia_semana)


                        # print(len(calendar))
                        d = -1
                        for i in calendar:
                            d += 1
                            if len(i) == 0:
                                continue 
                            else:
                                voz.falar(f'Na {key_dias[d]} você tem')
                                # print(len(i))
                                for e in range(len(i)):
                                    print(i)
                                    voz.falar(f'{i[e][0]} as {i[e][1]}')

                    
                    except Exception as e:
                        print(f'An error has ocorred: {e}')
                    calendar.clear()
                else:
                    pass                        


            elif 'fala' in comand:
                comand.remove('fala')
                num = ''
                for n in comand:
                    num = num + f'{n}'
                    num = num + ' '
                voz.falar(f'{num}')
                # print(comand    

            # Rotina Dormir
            elif 'dormir' in comand and 'vou' in comand or 'dormir' in comand and 'indo' in comand or 'boa' in comand and 'noite' in comand:
                brilho(0)
                sts = gas.mediaIs('PLAYING')
                if sts:
                    # kb.press_and_release('play/pause media')
                    os.popen(R'C:\Users\pedro\Desktop\nicory\lua\timer_musica.py')
                voz.falar('Boa noite chefin!')
                kb.press_and_release('win + d')
                alarme(7, 0, False)
                alarme(6, 30, False)
                voz.falar('Os alarmes foram definidos!')


            elif 'listar' in comand or 'lista' in comand:
                if 'alarmes' in comand:
                    import listalrm as lis
                    lis.listar()


            elif 'erro' in comand:
                ERRO = 1
                raise Exception('TESTE ERRO')

            # comandos pra aumentar ou diminuir brilho
            
            elif ('diminui' in comand or 'baixe' in comand or 'abaixa' in comand) and 'brilho' in comand:
                brilho(20) 
                # voz.falar('ok')

            elif ('aumenta' in comand or 'aumentar' in comand or 'aumente' in comand) and 'brilho' in comand:
                brilho(100)
                # voz.falar('ok')

            elif 'clica' in comand or 'clicar' in comand or 'clique' in comand:
                pa.click()

            elif 'pausa' in comand or 'pausar' in comand:
                kb.press_and_release('space')

            elif ('baixo' in comand or 'embaixo' in comand or 'cima' in comand or 'direita' in comand or 'esquerda' in comand):
                xpos, ypos = pa.position()

                for word in comand: 
                    if 'em' in comand:
                        comand.remove('em')
                    if 'para' in comand:
                        comand.remove('para')
                    if 'à' in comand:
                        comand.remove('à')
                    if 'a' in comand:
                        comand.remove('a')   
                    if 'sem' in comand:
                        comand[comand.index('sem')].replace('sem', '100')
                    if 'na' in comand:
                        comand.remove('na')

                print(comand)
                esquerda = False
                direita = False
                cima = False
                baixo = False

                if 'esquerda' in comand:
                    esquerda = True
                if 'direita' in comand:
                    direita = True
                if 'cima' in comand:
                    cima = True
                if 'baixo' in comand or 'embaixo' in comand:
                    baixo = True
                
                try:
                    if 'e' in comand:
                        comand.remove('e')

                        if esquerda:
                            xnew = xpos - int(comand[comand.index('esquerda') - 1])
                        elif direita:
                            xnew = xpos + int(comand[comand.index('direita') - 1])

                        if cima:
                            ynew = ypos - int(comand[comand.index('cima') - 1])
                        elif baixo:
                            ynew = ypos + int(comand[comand.index('baixo') - 1])

                        pa.moveTo(xnew, ynew)
                    elif cima or baixo:
                        if baixo:
                            ynew = ypos + int(comand[comand.index('baixo') - 1])
                        else:
                            ynew = ypos - int(comand[comand.index('cima') - 1])

                        pa.moveTo(xpos, ynew)
                    elif esquerda or direita:
                        if esquerda:
                            xnew = xpos - int(comand[comand.index('esquerda') - 1])
                        if direita:
                            xnew = xpos + int(comand[comand.index('direita') - 1])
                        pa.moveTo(xnew, ypos)

                    else:
                        voz.falar('comando não processado')
                        continue
                except:
                    voz.falar('comando errado')
                    continue
            elif 'abrir' in comand or 'abre' in comand:
                if 'code' in comand:
                    os.popen(R'C:\Users\pedro\Desktop\nicory\lua\lua.code-workspace')
                    voz.falar('Abrindo studio code')
                elif 'estudio' in comand or 'studio' in comand:
                    os.popen(R'C:\Users\pedro\Desktop\nicory\cursoJS\cursoJS.code-workspace')
                    voz.falar('Abrindo estúdio javascript')
                elif 'spotify' in comand:
                    kb.press_and_release('Alt + F1')
                    voz.falar('Abrindo spotify')    
                elif 'samsung' in comand:
                    os.popen(R'C:\Users\pedro\Desktop\nicory\scrcpy\samsung.pyw')
                    voz.falar('abrindo samsung')
                elif 'motorola' in comand:
                    voz.falar('ok')
                    os.popen(R'C:\Users\pedro\Desktop\nicory\scrcpy\motorola.pyw')
                else:
                    voz.falar('Fale novamente por gentileza')
            # nao reconhece comandoo
            else:
                voz.falar('Comando não reconhecido.')
                print(comand)
                continue

        else:
            # print(comand)
            continue
# enviar email de desligamento e erro 
except Exception:
    e = traceback.format_exc()
    erro = f""" 
    <p>Às {datetime.now().ctime()}, lua foi encerrada.</p>
    <p>Erro: {e} </p>
    """

    ea.enviar('Lua desligada!', erro)
    print(e)
    if ERRO:
        quit()
    else:
        voz.falar('Relatório de erro enviado. Reiniciando')
        reiniciar()
