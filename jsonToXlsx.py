import customtkinter # Módulo para la interfaz gráfica
import os # Módulo para manejo de directorios en windows
from PIL import Image # Manejo de imágenes
import json
import xlsxwriter # Módulo para crear archivos XLSX


class App(customtkinter.CTk):
    def __init__(self):
        super().__init__() # Constructor de la clase para la interfaz gráfica
        self.listaDeDocs = []
        self.title("JSON a XLSX")
        self.geometry("450x240")
        self.resizable(False,False)
        customtkinter.set_appearance_mode("dark")
        customtkinter.set_default_color_theme("green") # Los parámetros de la ventana de interfaz gráfica

        # set grid layout 1x2
        self.grid_rowconfigure(0, weight=1)
        self.grid_columnconfigure(1, weight=1)

        # load images with light and dark mode image
        image_path = os.path.join(os.path.dirname(os.path.realpath(__file__)), "imagenes")
        self.large_test_image = customtkinter.CTkImage(Image.open(os.path.join(image_path, "imagen.png")), size=(350, 162))
        self.image_icon_image = customtkinter.CTkImage(Image.open(os.path.join(image_path, "icono.png")), size=(20, 20))
        

        # create home frame
        self.home_frame = customtkinter.CTkFrame(self, corner_radius=0, fg_color="transparent")
        self.home_frame.grid_columnconfigure(0, weight=1)
        self.home_frame.grid(row=0, column=1, sticky="nsew")

        self.home_frame_large_image_label = customtkinter.CTkLabel(self.home_frame, text="", image=self.large_test_image)
        self.home_frame_large_image_label.grid(row=0, column=0, padx=20, pady=10)

        self.home_frame_button_2 = customtkinter.CTkButton(self.home_frame, text="Convertir", image=self.image_icon_image, compound="right", command=self.convertirJSON)
        self.home_frame_button_2.grid(row=2, column=0, padx=20, pady=10)

    def convertirJSON(self): # Método principal para la conversión
        self.listaDeDocs = []
        self.workbook = xlsxwriter.Workbook('Items.xlsx')
        self.worksheet = self.workbook.add_worksheet()
        list_of_files = self.get_list_of_json_files() # Obteniendo lista de archivos JSON en la carpeta 'archivosJson'
        fila = 1
        docnum = 1
        bold = self.workbook.add_format({'bold' : True})
        money = self.workbook.add_format({'num_format':'"$" #,##0.00'})
        self.primeraFila(bold)
        for file in list_of_files:
            fila = self.create_list_from_json(f'archivosJson/{file}',fila, docnum, money) # Llamada del método para abrir los JSON
            docnum += 1
        self.worksheet.autofit()
        self.workbook.close()
        self.archivoDeDtes()

    
    def get_list_of_json_files(self):
        bandera = True
        while bandera == True:
            try:
                list_of_files = os.listdir('archivosJson')
                for archivo in list_of_files:
                    if archivo[-5:-1] != ".jso":
                        list_of_files.remove(archivo)
                bandera = False
                return list_of_files
            except:
                os.mkdir('archivosJson')

    def primeraFila(self, bold):
        primFila = ["Num_Documento","Num_Item","Cantidad","Descripción","Precio Unitario","ventaGravada"]
        for file in primFila:
            col = 0
            for item in primFila:
                self.worksheet.write(0, col, primFila[col], bold)
                col += 1

    def create_list_from_json(self,jsonfile,fila,docnum,money):
        print(jsonfile)
        data_list = []
        try:
            with open(jsonfile, 'r', encoding='utf-8-sig', errors='replace') as f:
                data = json.load(f)
            item = []
            if len(data) > 0:
                try:
                    data_list.append(docnum)
                    data_list.append(data["identificacion"].get("numeroControl", "NULL"))
                    data_list.append(data["identificacion"].get("codigoGeneracion", "NULL"))
                    data_list.append(data["identificacion"].get("selloRecepcion", "NULL"))
                    data_list.append(data["identificacion"].get("fecEmi", "NULL"))
                    data_list.append(data["emisor"].get("nombre","NULL"))
                    data_list.append(data["emisor"].get("nit", "NULL"))
                    data_list.append(data["emisor"].get("nrc", "NULL"))
                    #try:
                    print(data["cuerpoDocumento"][0]["numItem"])
                    print(len(data["cuerpoDocumento"]))
                    data_list.append(data["resumen"]["totalGravada"])
                    try:
                        for tributo in data["resumen"]["tributos"]:
                            if tributo["codigo"] == "20":
                                data_list.append(tributo.get("valor", "NULL"))
                    except:
                        data_list.append(data["resumen"].get("totalIva","NULL"))
                    data_list.append(jsonfile)
                    self.listaDeDocs.append(data_list)
                    for uno in data["cuerpoDocumento"]:
                        item = []
                        item.append(docnum)
                        item.append(uno["numItem"])
                        item.append(uno["cantidad"])
                        item.append(uno.get("descripcion","NULL"))
                        item.append(uno.get("precioUni","NULL"))
                        item.append(uno.get("ventaGravada","NULL"))
                        col = 0
                        for dato in item:
                            if col > 2:
                                self.worksheet.write(fila, col, dato, money)
                                col += 1
                            else:
                                self.worksheet.write(fila, col, dato)
                                col += 1
                        fila += 1
                    return fila                            
                    """except:
                        print("fallo")
                        for i in range(1,10):
                            data_list.append("")
                        data_list.append(f'Error al leer items del archivo: {jsonfile}')
                        self.listaDeDocs.append(data_list.copy())
                        data_list = []
                        data_list.append(docnum)
                        for x in range(1,6):
                            data_list.append("")
                        data_list.append(f'Error al leer items del archivo: {jsonfile}')
                        col = 0
                        for dato in data_list:
                            self.worksheet.write(fila, col, dato)
                            col += 1
                        fila += 1
                        return fila"""
                except:
                    try:
                        data_list = []
                        data_list.append(docnum)
                        data_list.append(data["identificacion"].get("numeroControl", "NULL"))
                        data_list.append(data["identificacion"].get("codigoGeneracion", "NULL"))
                        data_list.append(data["identificacion"].get("selloRecepcion", "NULL"))
                        data_list.append(data["identificacion"].get("fecEmi", "NULL"))
                        data_list.append(data["emisor"].get("nombre","NULL"))
                        data_list.append(data["emisor"].get("nit", "NULL"))
                        data_list.append(data["emisor"].get("nrc", "NULL"))
                        data_list.append(data["cuerpoDocumento"].get("subTotal", "NULL"))
                        data_list.append(data["cuerpoDocumento"].get("iva", "NULL"))
                        data_list.append(jsonfile)
                        self.listaDeDocs.append(data_list)
                        item = []
                        item.append(docnum)
                        item.append(1)
                        item.append(1)
                        item.append("Documento contable de liquidación")
                        item.append(data["cuerpoDocumento"].get("subTotal","NULL"))
                        item.append(data["cuerpoDocumento"].get("iva", "NULL"))
                        col = 0
                        for dato in item:
                            if col > 2:
                                self.worksheet.write(fila, col, dato, money)
                                col += 1
                            else:
                                self.worksheet.write(fila, col, dato)
                                col += 1
                        fila += 1
                        return fila
                    except:
                        for i in range(1,11):
                            data_list.append("")
                        data_list.append(f'Error al leer doc de liquidacion del archivo: {jsonfile}')
                        self.listaDeDocs.append(data_list.copy())
                        data_list = []
                        data_list.append(docnum)
                        for x in range(1,6):
                            data_list.append("")
                        data_list.append(f'Error al leer doc de liquidacion del archivo: {jsonfile}')
                        col = 0
                        for dato in data_list:
                            self.worksheet.write(fila, col, dato)
                            col += 1
                        fila += 1
                        return fila
            else:
                data_list = []
                data_list.append(docnum)
                for i in range(1,11):
                    data_list.append("")
                data_list.append(f'Archivo vacio: {jsonfile}')
                self.listaDeDocs.append(data_list.copy())
                data_list = []
                data_list.append(docnum)
                for x in range(1,6):
                    data_list.append("")
                data_list.append(f'Archivo vacio: {jsonfile}')
                col = 0
                for dato in data_list:
                    self.worksheet.write(fila, col, dato)
                    col += 1
                fila += 1
                return fila
        except:
            print("falló al abrir")
            data_list = []
            data_list.append(docnum)
            for i in range(1,10):
                data_list.append("")
            data_list.append(f'Error al abrir el archivo: {jsonfile}')
            self.listaDeDocs.append(data_list.copy())
            data_list = []
            data_list.append(docnum)
            for x in range(1,6):
                data_list.append("")
            data_list.append(f'Error al abrir el archivo: {jsonfile}')
            col = 0
            for dato in data_list:
                self.worksheet.write(fila, col, dato)
                col += 1
            fila += 1
            return fila

        
    def archivoDeDtes(self):
        self.workbook = xlsxwriter.Workbook('DTEs.xlsx')
        self.worksheet = self.workbook.add_worksheet()
        bold = self.workbook.add_format({'bold' : True})
        money = self.workbook.add_format({'num_format':'"$" #,##0.00'})
        primFila = ["Num_Documento","Número de Control","Código Generación","Sello de Recepción","Fecha de Emisión","Nombre del Emisor","NIT Emisor","NRC","TotalGravada","IVA","Archivo JSON"]
        col = 0
        for item in primFila:
            self.worksheet.write(0, col, item, bold)
            col += 1
        fila = 1
        for row in self.listaDeDocs:
            col = 0
            for item in row:
                if col > 7:
                    self.worksheet.write(fila, col, item, money)
                    col += 1
                else:
                    self.worksheet.write(fila, col, item)
                    col += 1
            fila += 1
        self.worksheet.autofit()
        self.workbook.close()


    
if __name__ == '__main__':
    app = App()
    app.mainloop()