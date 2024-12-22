# import itertools
# from string import digits, punctuation, ascii_letters
# import win32com.client as client
# from datetime import datetime
# import time
# import os
# from pyfiglet import Figlet


# def clear():
#     os.system('cls' if os.name == 'nt' else 'clear')

# def menu():
#     preview_text = Figlet(font='slant')
#     print(preview_text.renderText("Senha Planilha"))


# def brute_excel_doc():
#     clear()
#     menu()
#     print("Hello friend!".center(30, '*'))

#     try:
#         # Solicita a faixa de comprimento da senha que o usuário deseja testar
#         password_length = input("Digite o comprimento da senha, de quantos - até quantos caracteres, por exemplo 3 - 7: ")
#         password_length = [int(item) for item in password_length.split("-")]
#     except:
#         # Caso o usuário insira dados incorretos
#         print("Verifique os dados inseridos")

#     # Solicita ao usuário o tipo de caracteres que a senha contém
#     print("Se a senha contiver apenas números, digite: 1\nSe a senha contiver apenas letras, digite: 2\n"
#           "Se a senha contiver números e letras, digite: 3\nSe a senha contiver números, letras e símbolos especiais, digite: 4")

#     try:
#         # Captura a escolha do usuário sobre o tipo de senha
#         choice = int(input(": "))
#         if choice == 1:
#             possible_symbols = digits  # Apenas números
#         elif choice == 2:
#             possible_symbols = ascii_letters  # Apenas letras
#         elif choice == 3:
#             possible_symbols = digits + ascii_letters  # Números e letras
#         elif choice == 4:
#             possible_symbols = digits + ascii_letters + punctuation  # Números, letras e símbolos
#         else:
#             possible_symbols = "O.o o que você quer, filho?"
#     except:
#         print("O.o o que você quer, filho?")

#     # Inicia o ataque de força bruta no documento Excel
#     start_timestamp = time.time()
#     print(f"Início em - {datetime.utcfromtimestamp(time.time()).strftime('%H:%M:%S')}")

#     count = 0
#     # Gera combinações de senhas com base na faixa de comprimento fornecida
#     for pass_length in range(password_length[0], password_length[1] + 1):
#         for password in itertools.product(possible_symbols, repeat=pass_length):
#             password = "".join(password)

#             # Abre o Excel via automação COM
#             opened_doc = client.Dispatch("Excel.Application")
#             count += 1

#             try:
#                 # Tenta abrir o arquivo Excel com a senha gerada
#                 opened_doc.Workbooks.Open(
#                     # r"T:\\Nan\\Doc\\MinervaFoods\\Banco de Dados\\pass.xlsx",
#                     # r"C:\Users\User\PycharmProjects\brute_excel\fsociety.xlsx",
#                     r"T:\\Users\\Verboten\\PycharmProjects\\pythonProjectII\\pass.xlsx",
#                     False,
#                     True,
#                     None,
#                     password
#                 )

#                 time.sleep(0.1)  # Pequena pausa para garantir que a abertura foi processada
#                 print(f"Finalizado em - {datetime.utcfromtimestamp(time.time()).strftime('%H:%M:%S')}")
#                 print(f"Tempo de quebra da senha - {time.time() - start_timestamp}")

#                 # Retorna a senha correta e o número de tentativas
#                 return f"Tentativa #{count} Senha correta: {password}"
#             except:
#                 # Exibe mensagem para cada senha incorreta
#                 print(f"Tentativa #{count} Senha incorreta: {password}")
#                 pass


# def main():
#     print(brute_excel_doc())


# if __name__ == '__main__':
#     main()


"""COD 02"""
# Senha: a123
# import itertools
# from string import digits, punctuation, ascii_letters
# import win32com.client as client
# from datetime import datetime
# import time
# import os
# from pyfiglet import Figlet
# from tkinter import Tk, filedialog

# def clear():
#     """Limpa o terminal."""
#     os.system('cls' if os.name == 'nt' else 'clear')

# def menu():
#     """Exibe o menu com um título formatado."""
#     preview_text = Figlet(font='slant')
#     print(preview_text.renderText("Senha Planilha"))

# def escolher_arquivo():
#     """Abre uma janela para selecionar o arquivo Excel e retorna o caminho do arquivo selecionado."""
#     root = Tk()
#     root.withdraw()
#     arquivo = filedialog.askopenfilename(
#         title="Escolha o arquivo Excel",
#         filetypes=[("Arquivos Excel", "*.xlsx")]
#     )
#     return arquivo

# def obter_tipo_caracteres():
#     """Solicita ao usuário o tipo de caracteres para a senha e retorna o conjunto correspondente."""
#     print("Se a senha contiver apenas números, digite: 1\n"
#           "Se a senha contiver apenas letras, digite: 2\n"
#           "Se a senha contiver números e letras, digite: 3\n"
#           "Se a senha contiver números, letras e símbolos especiais, digite: 4")

#     try:
#         choice = int(input(": "))
#         if choice == 1:
#             return digits
#         elif choice == 2:
#             return ascii_letters
#         elif choice == 3:
#             return digits + ascii_letters
#         elif choice == 4:
#             return digits + ascii_letters + punctuation
#         else:
#             print("Opção inválida. Usando apenas letras e números.")
#             return digits + ascii_letters
#     except ValueError:
#         print("Entrada inválida. Usando apenas letras e números.")
#         return digits + ascii_letters

# def brute_excel_doc():
#     """Executa o ataque de força bruta para encontrar a senha do arquivo Excel."""
#     clear()
#     menu()
#     print("Hello friend!".center(30, '*'))

#     try:
#         # Solicita a faixa de comprimento da senha
#         password_length = input("Digite o comprimento da senha, de quantos - até quantos caracteres, por exemplo 3 - 7: ")
#         password_length = [int(item) for item in password_length.split("-")]
#     except ValueError:
#         print("Dados inseridos incorretos. Verifique os dados.")
#         return

#     possible_symbols = obter_tipo_caracteres()
#     arquivo_excel = escolher_arquivo()

#     if not arquivo_excel:
#         print("Nenhum arquivo selecionado.")
#         return

#     start_timestamp = time.time()
#     print(f"Início em - {datetime.utcfromtimestamp(time.time()).strftime('%H:%M:%S')}")

#     count = 0
#     for pass_length in range(password_length[0], password_length[1] + 1):
#         for password in itertools.product(possible_symbols, repeat=pass_length):
#             password = "".join(password)
#             opened_doc = client.Dispatch("Excel.Application")
#             count += 1

#             try:
#                 opened_doc.Workbooks.Open(arquivo_excel, False, True, None, password)
#                 time.sleep(0.1)
#                 print(f"Finalizado em - {datetime.utcfromtimestamp(time.time()).strftime('%H:%M:%S')}")
#                 print(f"Tempo de quebra da senha - {time.time() - start_timestamp}")
#                 return f"Tentativa #{count} Senha correta: {password}"
#             except Exception as e:
#                 print(f"Tentativa #{count} Senha incorreta: {password}")
#                 pass

# def main():
#     """Função principal que executa o programa."""
#     resultado = brute_excel_doc()
#     if resultado:
#         print(resultado)

# if __name__ == '__main__':
#     main()


# import itertools
# from string import digits, punctuation, ascii_letters
# import win32com.client as client
# from datetime import datetime
# import time
# import os
# from pyfiglet import Figlet
# from tkinter import Tk, filedialog


# def clear():
#     """Limpa o terminal."""
#     os.system('cls' if os.name == 'nt' else 'clear')


# def menu():
#     """Exibe o menu com um título formatado."""
#     preview_text = Figlet(font='slant')
#     print(preview_text.renderText("Senha Planilha".upper()))


# def escolher_arquivo():
#     """Abre uma janela para selecionar o arquivo Excel e retorna o caminho do arquivo selecionado."""
#     root = Tk()
#     root.withdraw()
#     arquivo = filedialog.askopenfilename(
#         title="Escolha o arquivo Excel",
#         filetypes=[("Arquivos Excel", "*.xlsx")]
#     )
#     return arquivo


# def obter_tipo_caracteres():
#     """Solicita ao usuário o tipo de caracteres para a senha e retorna o conjunto correspondente."""
#     print("Escolha o tipo de caracteres da senha:")
#     print("1. Apenas números")
#     print("2. Apenas letras")
#     print("3. Números e letras")
#     print("4. Números, letras e símbolos especiais")

#     choice = input(": ").strip()
#     tipo_caracteres = {
#         '1': digits,
#         '2': ascii_letters,
#         '3': digits + ascii_letters,
#         '4': digits + ascii_letters + punctuation
#     }

#     return tipo_caracteres.get(choice, digits + ascii_letters)  # Padrão: números e letras


# def brute_excel_doc():
#     """Executa o ataque de força bruta para encontrar a senha do arquivo Excel."""
#     clear()
#     menu()
#     print("Hello friend!".center(30, '*'))

#     # Solicita a faixa de comprimento da senha
#     try:
#         password_length = input("Digite o comprimento da senha (Ex: 3-7): ")
#         min_len, max_len = map(int, password_length.split('-'))
#     except (ValueError, IndexError):
#         print("Dados inseridos incorretos. Use o formato correto (ex: 3-7).")
#         return

#     possible_symbols = obter_tipo_caracteres()
#     arquivo_excel = escolher_arquivo()

#     if not arquivo_excel:
#         print("Nenhum arquivo selecionado.")
#         return

#     start_timestamp = time.time()
#     print(f"Início em - {datetime.now().strftime('%H:%M:%S')}")

#     # Abre uma instância do Excel uma vez
#     opened_doc = client.Dispatch("Excel.Application")
#     opened_doc.DisplayAlerts = False

#     count = 0
#     for pass_length in range(min_len, max_len + 1):
#         for password_tuple in itertools.product(possible_symbols, repeat=pass_length):
#             password = "".join(password_tuple)
#             count += 1

#             try:
#                 # Tenta abrir o arquivo com a senha
#                 opened_doc.Workbooks.Open(arquivo_excel, False, True, None, password)
#                 print(f"Senha encontrada: {password} na tentativa #{count}")
#                 print(f"Finalizado em - {datetime.now().strftime('%H:%M:%S')}")
#                 print(f"Tempo total: {time.time() - start_timestamp:.2f} segundos")
#                 return f"Senha correta: {password}"
#             except Exception:
#                 if count % 100 == 0:  # Exibe status a cada 100 tentativas
#                     print(f"Tentativa #{count}, senha incorreta: {password}")
#                 continue  # Passa para a próxima senha

#     print("Nenhuma senha encontrada dentro do intervalo fornecido.")
#     opened_doc.Quit()  # Garante que o Excel seja fechado após a execução


# def main():
#     """Função principal que executa o programa."""
#     resultado = brute_excel_doc()
#     if resultado:
#         print(resultado)


# if __name__ == '__main__':
#     main()


"""Deepseek"""

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

