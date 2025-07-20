# AutoRewards

## Venv e instalação:
    python -m venv .venv
    .venv\Scripts\Activate

## Bibliotecas necessárias:
    pip install selenium
    pip install pyautogui
    pip install time
    pip install random
    pip install os
    pip install datetime

## Instruções:
    Executar isso para fazer o sistema funcionar!

    pyinstaller --onefile --noconsole --add-data "C:\Users\gusou\AppData\Roaming\nltk_data;nltk_data" MacroRewards.py

## Configurar pasta Chrome:

    É necessário executar o código abaixo para configurar a pasta do chrome e fazer login no microsoft store, salvar senha, colocar mecanismo de pesquisa como bing e concordar com todos os cookies para não ter erro futuramente!

    start chrome --user-data-dir="%USERPROFILE%\AppData\Local\Google\Chrome\User Data\SeleniumProfile"