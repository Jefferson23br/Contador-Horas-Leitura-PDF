#  Extrator de Dados de Viagens (PDF → Excel)

Este projeto tem como objetivo **extrair informações de viagens** de um arquivo PDF (gerado pelo sistema Expense Mobi) e preenchê-las automaticamente em uma planilha Excel modelo (`Horarios.xlsx`), preservando fórmulas e formatação.



##  Funcionalidade

O script lê os seguintes dados do PDF:

- **Data de Partida** (formato `DD/MM/AAAA`)
- **Hora de Partida** (formato `HH:MM:SS`)
- **Data de Chegada**
- **Hora de Chegada**
- **Distância** (somente parte inteira antes do "km")

Essas informações são inseridas automaticamente a partir da **linha 3** do arquivo Excel `Horarios.xlsx`, nas colunas:

| Coluna | Dados Inseridos        |
|--------|-------------------------|
| A      | Data de Partida         |
| B      | Hora de Partida         |
| C      | Data de Chegada         |
| D      | Hora de Chegada         |
| E      | Distância (em km)       |

---

##  Estrutura Esperada

Coloque os seguintes arquivos na mesma pasta do script:


>  O script sempre salva uma nova planilha chamada `Horarios_Preenchido.xlsx`.



## 🖥️ Como Executar o Script (CMD)

### 1. Instale as dependências (apenas uma vez):

pip install pymupdf openpyxl

2. Execute o script:

python seu_script.py

Ao executar, será aberta uma janela para você selecionar o PDF com os dados da viagem.

 Como Criar um Executável .exe
Você pode transformar o script Python em um arquivo .exe para facilitar o uso:

1. Instale o PyInstaller:

pip install pyinstaller

2. Gere o .exe:

pyinstaller --onefile seu_script.py
Após a finalização, o executável estará em:


dist/seu_script.exe
 Agora basta dar dois cliques no .exe para usar o programa sem precisar abrir o Python.

 Exemplo de Saída
Ao final da execução, será criado um novo arquivo:


Horarios_Preenchido.xlsx
Com todas as informações organizadas e preenchidas automaticamente.

Observações
O PDF deve seguir o modelo usado pelo Expense Mobi.

A planilha Horarios.xlsx precisa estar no mesmo diretório que o script.

Os dados são inseridos automaticamente a partir da linha 3, respeitando fórmulas e formatações.

Suporte
Em caso de dúvidas ou melhorias, sinta-se à vontade para entrar em contato.
