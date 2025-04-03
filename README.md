# Inserir-dados-automatizado
Script criado pensando em automatizar um pouco o processo de inserção de dados em planilhas na nuvem (Google Sheets)

# Só funciona com o arquivo .json ("""crendecial""") gerado devidamente no Google Cloud e também necessita do ID da planilha
Só substituir no código: "credenciais.json" e com o ID da planilha em "ID_SHEETS"

# você pode criar um arquivo executavel após o devido funcionamento do código (O executavel deve estar na mesma pasta em que as credenciais(credenciais.json) para funcionar)
cmd >pip install pyinstaller 
  >"""Nague até o diretorio do seu script.py"""
>    >seu_script.py > pyinstaller --onefile seu_script.py
