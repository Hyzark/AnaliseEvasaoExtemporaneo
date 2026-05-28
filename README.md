Crie um arquivo executável
```powershell
python -m PyInstaller --noconfirm --onefile --collect-all streamlit --collect-all pandas --add-data "app_evasao.py;." --add-data "appextemporaneo.py;." run_app.py```

Criar venv
```powershell
python -m venv .venv
```

Ativar venv
```powershell
.\.venv\Scripts\activate   
```

Instalar dependencias
```powershell
pip install -r requirements.txt  
```

Rodar aplicação
```powershell
streamlit run app_evasao.py
```
