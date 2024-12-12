import win32com
import win32com.client as win32
import datetime
import pandas as pd

def ExtraeReporteSAP():
    
        # Inicializar SAP GUI
        SapGuiAuto  = win32com.client.GetObject("SAPGUI")
        application = SapGuiAuto.GetScriptingEngine
        connection = application.Children(0)
        session    = connection.Children(0)

        # Ejecutar el proceso en SAP
        session.findById("wnd[0]").maximize
        session.findById("wnd[0]/tbar[0]/okcd").Text = "/n/CASWW/SE16D"
        session.findById("wnd[0]").sendVKey(0)
        session.findById("wnd[0]/usr/ctxtGD-TAB").Text = "LIPS"
        session.findById("wnd[0]/usr/txtGD-MAX_LINES").Text = "50"
        session.findById("wnd[0]/usr/ctxtGD-VARIANT").Text = "REPORTEPALLT"
        #Aqui si necesita posicionarse en el boton para despues precionar
        session.findById("wnd[0]/usr/tblSAPLSE16NSELFIELDS_TC/btnPUSH[4,1]").setFocus()
        session.findById("wnd[0]/usr/tblSAPLSE16NSELFIELDS_TC/btnPUSH[4,1]").press()
        
        session.findById("wnd[1]/tbar[0]/btn[7]").press() #New Entries
        session.findById("wnd[1]/tbar[0]/btn[24]").press()#pegar
        session.findById("wnd[1]/tbar[0]/btn[8]").press()#Reloj pantalla secundaria
        session.findById("wnd[0]/tbar[1]/btn[8]").press()#Reloj pantalla principal
      
        session.findById("wnd[0]/usr/cntlRESULT_LIST/shellcont/shell").pressToolbarContextButton("&MB_EXPORT")
        session.findById("wnd[0]/usr/cntlRESULT_LIST/shellcont/shell").selectContextMenuItem("&PC")
        session.findById("wnd[1]/usr/subSUBSCREEN_STEPLOOP:SAPLSPO5:0150/sub:SAPLSPO5:0150/radSPOPLI-SELFLAG[1,0]").select()
        session.findById("wnd[1]/usr/subSUBSCREEN_STEPLOOP:SAPLSPO5:0150/sub:SAPLSPO5:0150/radSPOPLI-SELFLAG[1,0]").setFocus()
        session.findById("wnd[1]/tbar[0]/btn[0]").press()
        session.findById("wnd[1]/usr/ctxtDY_PATH").text = "C:\\Users\\uidp4308\\Documents\\Reporte Embarques\\SAP\\"
        session.findById("wnd[1]/usr/ctxtDY_FILENAME").text = "ejemplo.xls"
        #session.findById("wnd[1]/usr/ctxtDY_FILENAME").caretPosition = 12
        session.findById("wnd[1]/tbar[0]/btn[11]").press()


ExtraeReporteSAP()
RutaArchivoSAP = "C:\\Users\\uidp4308\\Documents\\Reporte Embarques\\SAP\\"
NombreLibro ="ejemplo.xls"
# Abrir el archivo en Excel y copiar los datos
excel = win32.Dispatch("Excel.Application")
wb = excel.Workbooks.Open(RutaArchivoSAP + NombreLibro)
ws = wb.Sheets(1)  # Cambia al número de hoja que deseas leer

# Lee los datos de la hoja activa
data = ws.UsedRange.Value
wb.Close()
excel.Quit()

# Convertir a DataFrame
df = pd.DataFrame(data[1:], columns=data[0])  # Omite el encabezado si corresponde
df = df.drop([0,1,3,5])
#Reiniciar los índices si lo deseas
df.reset_index(drop=True, inplace=True)

df = df.drop(df.columns[0], axis=1)
df.columns = df.iloc[0]  # Establece la primera fila como los nombres de las columnas
df = df.drop(0).reset_index(drop=True)  # Elimina la primera fila (ahora ya es encabezado)
# Cambiar el nombre de la columna 'A' a 'Nuevo_A'
df.rename(columns={'Dlv.qty': 'Cant'}, inplace=True)
df['Delivery'] = df["Delivery"].astype(int)

# Guardar el DataFrame en un archivo Excel
df.to_excel(RutaArchivoSAP + "Ejemplo.xlsx", index=False)  # `index=False` para no guardar el índice #'ejemplo.xlsx'

