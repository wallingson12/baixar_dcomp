import os
from tkinter import Tk, Label, Button, filedialog, messagebox
from PIL import Image, ImageTk
import pyautogui
import pandas as pd
from pynput import mouse

def capture_mouse_click():
    """
    Captura a posição de um único clique do mouse e retorna a posição capturada.
    """
    print("Clique em qualquer lugar na tela para capturar a posição...")

    # Variável para armazenar a posição do clique
    click_position = [None, None]  # Inicialmente nenhum clique capturado

    # Função chamada quando ocorre um clique do mouse
    def on_click(x, y, button, pressed):
        nonlocal click_position
        if pressed:
            click_position = [x, y]  # Armazena a posição do clique
            return False  # Encerra o listener após capturar o primeiro clique

    # Cria e inicia um listener de eventos de clique do mouse
    with mouse.Listener(on_click=on_click) as listener:
        listener.join()  # Aguarda até que o listener capture o primeiro clique

    return click_position  # Retorna a posição capturada do clique

def processar_dados():
    try:
        filename = filedialog.askopenfilename(initialdir="/", title="Selecione o arquivo Excel",
                                              filetypes=(("Arquivos Excel", "*.xlsx"), ("Todos os arquivos", "*.*")))
        if filename:
            # Capturar a posição do clique após selecionar o arquivo
            position = capture_mouse_click()
            if position != [None, None]:
                x, y = position
                print(f'Posição do clique capturada - X: {x}, Y: {y}')
                # Aqui você pode utilizar a posição do clique como necessário
            else:
                print("Nenhuma posição de clique foi capturada.")

            # Continuar com o processamento dos dados usando o arquivo Excel selecionado
            df = pd.read_excel(filename)

            # Filtrar o DataFrame para incluir apenas as linhas que não estão marcadas como "Sim" na coluna "Processado"
            df_nao_processado = df[df["Processado"] != "Sim"]

            # Verificar se há linhas a serem processadas
            if df_nao_processado.empty:
                messagebox.showinfo("Concluído", "Todas as linhas já foram processadas.")
                return

            #pyautogui.hotkey('alt', 'tab')
            pyautogui.sleep(2)
            for index, row in df_nao_processado.iterrows():
                numero = row["Número do PER/DCOMP"]
                pyautogui.click(position)
                pyautogui.sleep(1)
                pyautogui.write(str(numero))
                pyautogui.sleep(1)
                pyautogui.press('tab')
                pyautogui.sleep(1)
                pyautogui.click('salvar_dcomp.png')
                pyautogui.sleep(1)
                pyautogui.click(position)
                pyautogui.hotkey('ctrl', 'a')
                pyautogui.sleep(1)
                pyautogui.press('del')
                pyautogui.sleep(1)
                pyautogui.press('tab')

            # Atualizar a planilha para registrar que o processamento foi concluído
            df.loc[df.index.isin(df_nao_processado.index), "Processado"] = "Sim"
            df.to_excel(filename, index=False)
            messagebox.showinfo("Concluído", "Processamento concluído com sucesso!")
    except Exception as e:
        messagebox.showerror("Erro", f"Ocorreu um erro durante o processamento:\n{str(e)}")

# Criar a janela principal
janela = Tk()
janela.title("Processamento de Dados")

# Carregar e exibir a imagem
icone_path = "Toad.jpg"
if os.path.exists(icone_path):
    imagem = Image.open(icone_path)
    imagem = imagem.resize((200, 150))  # Redimensionar a imagem para o tamanho desejado
    imagem = ImageTk.PhotoImage(imagem)

    label_imagem = Label(janela, image=imagem)
    label_imagem.pack()

# Definir tamanho fixo (largura x altura)
largura = 300
altura = 300
janela.geometry(f"{largura}x{altura}")

# Definir cor de fundo azul
janela.configure(bg="#00426b")

# Obter as dimensões da tela
largura_tela = janela.winfo_screenwidth()
altura_tela = janela.winfo_screenheight()

# Calcular as coordenadas para o centro da tela
pos_x = (largura_tela - largura) // 2
pos_y = (altura_tela - altura) // 2

# Definir a posição da janela no centro da tela
janela.geometry(f"+{pos_x}+{pos_y}")

# Adicionar um botão para iniciar o processamento
botao_processar = Button(janela, text="Processar Dados", command=processar_dados)
botao_processar.pack(pady=20)

# Executar a janela
janela.mainloop()
