""" Força Bruta - Excel """
import itertools
from string import digits, punctuation, ascii_letters
import win32com.client as client
from datetime import datetime
import time
import os
from pyfiglet import Figlet
from tkinter import Tk
from tkinter.filedialog import askopenfilename

def clear():
    """Limpa o terminal."""
    os.system('cls' if os.name == 'nt' else 'clear')

def menu():
    """Exibe o menu com um título formatado."""
    preview_text = Figlet(font='slant')
    print(preview_text.renderText("SENHA PLANILHA"))

def escolher_arquivo():
    """Abre uma janela para selecionar o arquivo Excel e retorna o caminho do arquivo selecionado."""
    root = Tk()
    root.withdraw()
    arquivo = askopenfilename(
        title="Escolha o arquivo Excel",
        filetypes=[("Arquivos Excel", "*.xlsx")]
    )
    return arquivo

def obter_tipo_caracteres():
    """Solicita ao usuário o tipo de caracteres para a senha e retorna o conjunto correspondente."""
    print("Escolha o tipo de caracteres da senha:")
    print("1. Apenas números")
    print("2. Apenas letras")
    print("3. Números e letras")
    print("4. Números, letras e símbolos especiais")

    choice = input(": ").strip()
    tipo_caracteres = {
        '1': digits,
        '2': ascii_letters,
        '3': digits + ascii_letters,
        '4': digits + ascii_letters + punctuation
    }
    return tipo_caracteres.get(choice, digits + ascii_letters)  # Padrão: números e letras

def brute_excel_doc():
    """Executa o ataque de força bruta para encontrar a senha do arquivo Excel."""
    clear()
    menu()
    print("Hello friend!".center(30, '*'))

    # Solicita a faixa de comprimento da senha
    password_length = input("Digite o comprimento da senha (Ex: 3-7): ")
    try:
        min_len, max_len = map(int, password_length.split('-'))
    except ValueError:
        print("Formato inválido. Use o formato correto (ex: 3-7).")
        return

    possible_symbols = obter_tipo_caracteres()
    arquivo_excel = escolher_arquivo()

    if not arquivo_excel:
        print("Nenhum arquivo selecionado.")
        return

    start_timestamp = time.time()
    print(f"Início em - {datetime.now().strftime('%H:%M:%S')}")

    # Abre uma instância do Excel
    excel_app = client.Dispatch("Excel.Application")
    excel_app.DisplayAlerts = False  # Desativa alertas

    count = 0
    try:
        for pass_length in range(min_len, max_len + 1):
            for password_tuple in itertools.product(possible_symbols, repeat=pass_length):
                password = ''.join(password_tuple)
                count += 1

                # Tenta abrir o arquivo com a senha
                try:
                    workbook = excel_app.Workbooks.Open(arquivo_excel, False, True, None, password)
                    print(f"\nSenha encontrada: {password} na tentativa #{count}")
                    print(f"Finalizado em - {datetime.now().strftime('%H:%M:%S')}")
                    print(f"Tempo total: {time.time() - start_timestamp:.2f} segundos")
                    workbook.Close(SaveChanges=False)
                    excel_app.Quit()
                    return f"Senha correta: {password}"
                except client.pywintypes.com_error:
                    # Senha incorreta, continua com a próxima tentativa
                    if count % 100 == 0:
                        print(f"Tentativa #{count}, senha incorreta: {password}", end='\r')
    finally:
        # Garante que o aplicativo do Excel seja fechado
        if excel_app:
            excel_app.Quit()

    print("\nNenhuma senha encontrada dentro do intervalo fornecido.")
    print(f"Tempo total: {time.time() - start_timestamp:.2f} segundos")

def main():
    """Função principal que executa o programa."""
    brute_excel_doc()

if __name__ == '__main__':
    main()

