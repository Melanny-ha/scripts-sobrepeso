import streamlit as st
from io import BytesIO
import pandas as pd
from openpyxl import load_workbook

#Funcion para convertir DataFrame a Excel(Descargar tabla)
def convertir_a_excel(df):
  output = BytesIO()
  with pd.ExcelWriter(output, engine='xlsxwriter') as writer:
      df.to_excel(writer, index=False, sheet_name='Datos')
  return output.getvalue()

#Funcion para reducir decimales visualmente
def formatear_visual(df, columnas_0f=[], columnas_2f=[]):
  df_mostrar = df.copy()
  for col in columnas_0f:
    df_mostrar[col] = df_mostrar[col].apply(lambda x: f"{x:.0f}")
  for col in columnas_2f:
    df_mostrar[col] = df_mostrar[col].apply(lambda x: f"{x:.2f}")
  return df_mostrar

#Estilos streamlit
def estilos_css():
  st.markdown("""
      <style>
          /*widht->ancho  heigth->alto*/
          /*Estilos generales*/
          .main {background-color: #f0f0f0;}
          .sidebar header{color: red;}
          .image { padding: 200px;}
          
          /*Estilos para las tablas*/
          .custom-table {
              max-height: 300px;
              overflow-y: auto; /* Solo scroll vertical */
              display: block;
              width: 100%;
              margin: 10px 0 12px 0;
              border-radius: none;
          }
          .custom-table table {
              table-layout: auto; /* Permite que las columnas se ajusten automáticamente */
              width: 100%;
              border-collapse: collapse;
              font-size: 12px;
              min-width: 20px;
          }
          .dataframe{font-size: 12px;}
          .dataframe tbody tr td {
              font-size: 10px;
              padding: 5px;
              white-space: nowrap;
          }
          .custom-table th, .custom-table td {
              border: 1px solid #ddd;  /*grosor bordes tabla*/
              padding: 8px;  /*ancho del cuadro del titulo*/  
              text-align: center;
              word-wrap: break-word;
              white-space: normal;
          }
          .custom-table th {  /*color fondo y letra encabezado*/
              background-color: #ff4630;  
              color: white;
          }
          .custom-table tr:nth-child(even){  /*filas gris oscuras en la tabla*/
              background-color: #dedede;
          }
          .custom-table tr:hover {  /*cambiar color de fila al pasar por ella*/
              background-color: #f35856;
              color: white;
          }
          .dataframe tbody tr td {
              font-size: 14px !important;  /* Tamaño del texto */
              padding: 0px 4px !important;  /* Ancho dentro de la celda */
          }
          .dataframe thead tr th {
              font-size: 12px !important;  /* Tamaño del texto del encabezado */
              padding: 5px !important;  /* Espaciado dentro de las celdas del encabezado */
          }
              
          /*Estilos para el boton descargar*/
          .stDownloadButton {
              display: flex;
              justify-content: flex-end; /* Alinea el contenedor a la derecha */
              margin-top: 5px;
          }
          .stDownloadButton>button {
              background-color: #fe5237;
              color: white;
              padding: 10px 20px;
              border: none;
              border-radius: 10px;
              font-size: 10px;
              font-weight: bold;
              cursor: pointer;
              transition: background-color 0.2s ease;
          }
          .stDownloadButton>button:hover {
              background-color: #c3270f;
              color: white;
          }

          /*Estilos para las tarjetas*/
          .metric-card {
              background-color: white;
              margin: 6px;
              padding: 12px;
              border-radius: 10px;
              box-shadow: 2px 2px 10px rgba(0,0,0,0.1);
              font-size: 13px;
              text-align: center;
          }
      </style>
  """, unsafe_allow_html=True)
