# SENHA PLANILHA

Este projeto é uma ferramenta em Python para realizar um ataque de força bruta em arquivos do Microsoft Excel protegidos por senha. Ele permite ao usuário escolher um arquivo Excel e definir os parâmetros do ataque, como comprimento da senha e tipos de caracteres a serem utilizados. Este programa é destinado exclusivamente a fins educacionais e éticos. ⚠️ **Certifique-se de ter permissão para testar arquivos EXCEL antes de utilizá-lo.**


---

## ⚙️ Funcionalidades

- **Seleção de Arquivo**: Permite escolher o arquivo Excel a ser processado por meio de uma janela gráfica.
- **Configuração de Ataque**:
  - Escolha do comprimento da senha (intervalo).
  - Seleção do tipo de caracteres: números, letras, combinação ou inclusão de símbolos especiais.
- **Execução de Força Bruta**:
  - Geração de combinações de senha com base nos parâmetros fornecidos.
  - Tentativas automáticas de abrir o arquivo até encontrar a senha correta ou finalizar todas as combinações possíveis.
- **Exibição de Progresso**:
  - Número da tentativa atual.
  - Tempo total de execução.
  - Resultado final com a senha encontrada ou mensagem de falha.

---

## 📋 Pré-requisitos

Certifique-se de ter as seguintes dependências instaladas:

- Python 3.8 ou superior
- Bibliotecas Python:
  - `pyfiglet`
  - `tkinter`
  - `pywin32`
- Microsoft Excel instalado no sistema operacional.

---

## 🚀 Como usar

1. **Clone o repositório ou baixe os arquivos**:
   ```bash
   git clone https://github.com/Alan-0718-sj/forca-bruta-excel
   cd forca-bruta-excel
