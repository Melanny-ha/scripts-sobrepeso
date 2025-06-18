import streamlit as st
from informe_EnvSol import mostrar_informe_env_sol
from informe_Mol import mostrar_informe_molido

st.sidebar.markdown('<h2 style="font-size: 22px; color: #eb4f37; font-weight: bold;">Menú de Informes</h2>', unsafe_allow_html=True)
opcion_seleccionada = st.sidebar.radio(
    "Ir a:", ("Envase Soluble", "Molido")
)

if opcion_seleccionada == "Envase Soluble":
    mostrar_informe_env_sol()
elif opcion_seleccionada == "Molido":
    mostrar_informe_molido()
