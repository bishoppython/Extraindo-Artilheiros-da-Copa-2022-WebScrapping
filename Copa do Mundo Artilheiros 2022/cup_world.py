import time
import requests, openpyxl
from bs4 import BeautifulSoup as bs
from winotify import Notification, audio

# Notificação no Desktop do Windows
toast = Notification(app_id="Processo de WebScrapping",
                     title="Finalizado o Processo de WebScrapping",
                     msg="Raspagem completada com sucesso!!",
                     duration="long",
                     icon=r"C:\Users\anderson.bispo.HMVCR\Downloads\icons8-bitcoin-stickers\4.png")

toast.set_audio(audio.LoopingCall, loop=True)
toast.add_actions(label="Clique Aqui", launch="C:\\Users\\anderson.bispo.HMVCR\\PycharmProjects\\Crawler")


# Habilitando o módulo Excel
excel = openpyxl.Workbook()
print(excel.sheetnames)
sheet = excel.active
sheet.title = "Artilheiros"
print(excel.sheetnames)
sheet.append(['Rank','Seleção', 'Nome', 'Posição', 'Quant. Gols', 'Foto Jogador', 'Escudo Seleção'])


def artilheirosCopa():
    print("\n Raspagem de Dados Iniciada... ")
    headers = {
        'User-Agent': 'Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36(KHTML, like Gecko)'
                      'Chrome/107.0.0.0 Safari/537.36',
    }

    pageUrl = "https://ge.globo.com/futebol/copa-do-mundo/2022/"#+str(page)
    response = requests.get(pageUrl, headers=headers)
    #print(response.status_code) --Mostra o status do site/link

    #print(response.text)

    #print("\n Filtrando Conteúdo por página", page, "... ")
    soup = bs(response.content, 'html.parser')

    allContent = soup.find_all('div', class_='ranking-item-wrapper')

    for artilheiro in allContent:
        rank = artilheiro.find('div', class_='ranking-item').text
        jogador_foto = artilheiro.find('div', attrs={'class':'jogador-foto'}).find('img').get('src')
        jogador_escudo = artilheiro.find('div', attrs={'class':'jogador-escudo'}).find('img').get('src')
        jogador_selecao = artilheiro.find('div', attrs={'class': 'jogador-escudo'}).find('img').get('alt')
        jogador_info = artilheiro.find('div', attrs={'class':'jogador-nome'}).text
        jogador_pose = artilheiro.find('div', attrs={'class':'jogador-posicao'}).text
        gols = artilheiro.find('div', class_='jogador-gols').text

        sheet.append([rank, jogador_selecao, jogador_info, jogador_pose, gols, jogador_foto, jogador_escudo])

    print("\n Salvando Conteúdo em arquivo excel... ")




    # Salvando em excel
    excel.save('Artilheiros da Copa 2022.xlsx')
    toast.show() #Chama a notificação
    print("\n Conteúdo salvo em arquivo Excel... COMPLETO! ")
    print("\n Raspagem completada com sucesso!! ")

if __name__=='__main__':
    artilheirosCopa()