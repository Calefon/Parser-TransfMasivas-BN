import pandas as pd
import sys

def show_exception_and_exit(exc_type, exc_value, tb):
    import traceback
    traceback.print_exception(exc_type, exc_value, tb)
    input("Press key to exit.")
    sys.exit(-1)

sys.excepthook = show_exception_and_exit

CBU_ETEC = "0"*22

LARGO_DE_LINEA = 218
EXCEL = input("Ingrese nombre o ruta del archivo (con extension):")
excelDF = pd.read_excel(EXCEL)
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
        self.salto_Linea = chr(10)+chr(13)
        self.cant_registros = 0
        self.total_importes = 0
        self.relleno = " "*194
    
    def getCamposFromRow(self,rowDF):
        
        self.cbu_debito = CBU_ETEC
        
        self.cbu_credito = str(rowDF.at[rowDF.index[0],"CBU"])
        

        self.importe = str(rowDF.at[rowDF.index[0],"Monto"]*100)
        
        self.importe = "0"*(12-len(self.importe)) + self.importe

        self.cant_registros+=1
        self.total_importes+= rowDF.at[rowDF.index[0],"Monto"]*100

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

with open(conc[:10]+".txt", "w") as transfTXT:
    lineaTXT = ""
    for i in excelDF.index:
        camposLinea.getCamposFromRow(excelDF.iloc[[i]])
        lineaTXT = camposLinea.genLine(lastLine=False)
        print(len(lineaTXT))
        if(len(lineaTXT) != LARGO_DE_LINEA ):
            raise Exception('\n\nError: una linea no tiene el formato correcto (longitud de caracteres distinta de 218)')
        transfTXT.write(lineaTXT)
    
    lineaTXT = camposLinea.genLine(lastLine=True)
    transfTXT.write(lineaTXT)

print(f'Terminado!\nCantidad de lineas procesadas: {camposLinea.cant_registros+1}\nImporte total: {camposLinea.total_importes}')

input("Toque Enter para cerrar")




    
















