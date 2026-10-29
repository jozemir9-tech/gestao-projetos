# -*- coding: utf-8 -*-
"""Entry point do Streamlit Community Cloud.

O Community Cloud procura `streamlit_app.py` na raiz do repositório por padrão;
o app de verdade é o `app.py` (todas as páginas, helpers e dados ficam lá).
Este arquivo só importa e chama o `main()` — não duplica lógica.

Rodar localmente (equivalente a `streamlit run app.py`):
    streamlit run streamlit_app.py
"""
from app import main

if __name__ == "__main__":
    main()
