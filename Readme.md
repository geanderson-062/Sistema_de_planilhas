# Sistema de Planilhas

## Descrição

Este projeto é um sistema de planilhas desenvolvido em Python utilizando as bibliotecas Tkinter e CustomTkinter. Ele permite ao usuário adicionar dados, visualizá-los em uma tabela e salvá-los como um arquivo Excel.

## Funcionalidades

- **Adicionar Dados**: Insere dados nos campos e os armazena em uma lista.
- **Visualizar Tabela**: Mostra os dados adicionados em uma tabela numa nova janela.
- **Salvar em Excel**: Exporta os dados para um arquivo Excel.

## Como Executar

Para executar este programa, você precisa ter o Python instalado em seu sistema, além das bibliotecas `tkinter`, `customtkinter`, `openpyxl` e `ttk`. Após instalar as dependências, execute o script Python.

## Personalização e Alterações

### Alterar Tema e Cores

Você pode alterar o tema e as cores padrão do aplicativo ajustando as configurações de `customtkinter.set_appearance_mode` e `customtkinter.set_default_color_theme`.

### Adicionar Novos Campos

Para adicionar novos campos de entrada, siga estes passos:

1. Crie um novo `CTkEntry` para o campo.
2. Posicione-o na janela usando o método `grid`.
3. Adicione a lógica para coletar e armazenar os dados desse campo nas funções `adicionar_dados` e `salvar_excel`.

### Alterar Layout

O layout é gerenciado pelo método `grid`. Você pode ajustar a posição dos widgets alterando os parâmetros `row` e `column` na chamada do método `grid`.

### Modificar Funções

- **adicionar_dados**: Adiciona os dados dos campos de entrada à lista `dados_salvos`.
- **mostrar_tabela**: Abre uma nova janela com uma tabela contendo os dados inseridos.
- **salvar_excel**: Exporta os dados para um arquivo Excel.

### Estilizar Widgets

Você pode personalizar a aparência dos widgets (botões, campos de entrada, labels, etc.) alterando seus parâmetros na criação. Por exemplo, mudar cores, fontes, tamanhos, etc.

## Dependências

- Python
- Tkinter
- CustomTkinter
- openpyxl
- ttk (parte do Tkinter)

# Guia de Instalação

## Pré-requisitos

Certifique-se de ter o Python instalado em seu sistema. Este projeto foi desenvolvido utilizando Python, e é essencial tê-lo para executar o script. Se você ainda não tem o Python, pode baixá-lo do site oficial: [Download Python](https://www.python.org/downloads/).

## Instalação de Dependências

### Tkinter

Tkinter é uma biblioteca padrão do Python para desenvolvimento de interfaces gráficas. Normalmente já vem instalada com o Python.

### CustomTkinter

CustomTkinter é uma versão personalizada do Tkinter com widgets modernos. Instale-a utilizando o pip:

```bash
pip install customtkinter
```

### openpyxl

openpyxl é uma biblioteca para ler e escrever arquivos Excel (xlsx/xlsm/xltx/xltm). Instale utilizando o pip:

```bash
pip install openpyxl
```

## Verificação da Instalação

Após a instalação das dependências, verifique se tudo foi instalado corretamente:

```python
python -c "import tkinter; import customtkinter; import openpyxl"
```

## Executando o Aplicativo

Com todas as dependências instaladas, execute o script do sistema de planilhas:

```bash
python nome_do_script.py
```

---

Se encontrar algum erro, verifique se todas as dependências estão corretamente instaladas e se a versão do Python é compatível com as bibliotecas.
