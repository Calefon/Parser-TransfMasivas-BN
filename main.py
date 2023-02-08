import pandas as pd
import sys
#Parser para la generación de TXT de transferencias masivas Banco Nación
#Formato de columnas excel de entrada:
#Nombre/Apellido/CBU/Monto/Concepto(1 celda)/Referencia(1 celda)

#Manejo de excepción para que no se cierre la terminal
def show_exception_and_exit(exc_type, exc_value, tb):
    import traceback
    traceback.print_exception(exc_type, exc_value, tb)
    input("Press key to exit.")
    sys.exit(-1)

sys.excepthook = show_exception_and_exit

#Constantes
CBU_ETEC = "0110599520000054713868"
LARGO_DE_LINEA = 218

EXCEL = input("Ingrese nombre o ruta del archivo (con extension):")

#Carga de archivo excel en Pandas Dataframe
excelDF = pd.read_excel(EXCEL)

#Extracción de Concepto y Referencia
conc = str(excelDF.at[0,'Concepto'])
if not isinstance(conc, str):
    conc = str(int(conc))

ref = excelDF.at[0,'Referencia']
if not isinstance(ref, str):
    ref = str(int(ref))

if len(conc)>50:
    raise Exception('\n\nError: Concepto es demasiado largo, solo se admiten hasta 50 caracteres')
if len(ref)>12:
    raise Exception('\n\nError: Referencia es demasiado largo, solo se admiten hasta 50 caracteres')

#Total a transferir
TOTAL = excelDF["Importe"].sum()
TOTAL = int(TOTAL*100)/100

#Clase en donde se parsea cada campo, representando una línea.
class Campos:
    def __init__(self):
        self.cbu_debito = CBU_ETEC
        self.cbu_credito = ""
        self.alias_cbu_debito = " "*22
        self.alias_cbu_credito = " "*22
        self.importe = ""
        self.concepto = conc+" "*(50-len(conc))
        self.motivo = "VAR"
        self.referencia = ref+" "*(12-len(ref))
        self.email = " "*50
        self.titulares = " "
        self.salto_Linea = chr(13)+chr(10)
        self.cant_registros = 0
        self.total_importes = 0
        self.relleno = " "*194
    
    def getCamposFromRow(self,rowDF):
        
        self.cbu_debito = CBU_ETEC
        
        self.cbu_credito = str(rowDF.at[rowDF.index[0],"CBU"])
        

        self.importe = str(int(rowDF.at[rowDF.index[0],"Importe"]*100))
        
        self.importe = "0"*(12-len(self.importe)) + self.importe

        self.cant_registros+=1
        self.total_importes+= int(rowDF.at[rowDF.index[0],"Importe"]*100)

    def genLine(self,lastLine):
        if not lastLine:
            line = ""
            line+=self.cbu_debito + self.cbu_credito + self.alias_cbu_debito + self.alias_cbu_credito + self.importe + self.concepto + self.motivo + self.referencia + self.email + self.titulares + self.salto_Linea
        else:
            line = ""
            cant = str(self.cant_registros+1)
            total = str(self.total_importes)
            line+= "0"*(5-len(cant))+cant+"0"*(17-len(total))+total+self.relleno+self.salto_Linea
        return line

camposLinea = Campos()
TXTstream = ""

#Lectura del DF línea por línea y escritura al TXT
with open(conc[:20]+".txt", "wb") as transfTXT:
    lineaTXT = ""
    for i in excelDF.index:
        camposLinea.getCamposFromRow(excelDF.iloc[[i]])
        lineaTXT = camposLinea.genLine(lastLine=False)
        print(len(lineaTXT))
        if(len(lineaTXT) != LARGO_DE_LINEA ):
            raise Exception('\n\nError: una linea no tiene el formato correcto (longitud de caracteres distinta de 218)')
        transfTXT.write(bytes(lineaTXT, "UTF-8"))

    lineaTXT = camposLinea.genLine(lastLine=True)
    
    transfTXT.write(bytes(lineaTXT, "UTF-8"))
    if (camposLinea.total_importes/100-TOTAL) != 0:
        print(f'{camposLinea.total_importes/100}; {TOTAL}')
        raise Exception("Error: El total en archivo no coincide con total en Excel")


print(f'Terminado!\nCantidad de lineas procesadas: {camposLinea.cant_registros+1}\nImporte total: {camposLinea.total_importes/100}')

input("Toque Enter para cerrar")




    
















