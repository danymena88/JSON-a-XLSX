import customtkinter # Módulo para la interfaz gráfica
import os # Módulo para manejo de directorios en windows
from PIL import Image # Manejo de imágenes
import json
import xlsxwriter # Módulo para crear archivos XLSX


class App(customtkinter.CTk):
    def __init__(self):
        super().__init__() # Constructor de la clase para la interfaz gráfica

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
        self.workbook = xlsxwriter.Workbook('DTEs.xlsx')
        self.worksheet = self.workbook.add_worksheet()
        list_of_files = self.get_list_of_json_files() # Obteniendo lista de archivos JSON en la carpeta 'archivosJson'
        fila = 1
        try:
            self.primeraFila()
            for file in list_of_files:
                col = 0
                row = self.create_list_from_json(f'archivosJson/{file}') # Llamada del método para abrir los JSON
                for item in row:
                    self.worksheet.write(fila, col, row[col])
                    col += 1
                fila += 1
            self.worksheet.autofit()
            self.workbook.close()
        except:
            print("Archivo abierto")

    
    def get_list_of_json_files(self):
        bandera = True
        while bandera == True:
            try:
                list_of_files = os.listdir('archivosJson')
                for archivo in list_of_files:
                    if archivo[-5:-1] != ".jso":
                        list_of_files.remove(archivo)
                bandera = False
                print(list_of_files)
                return list_of_files
            except:
                os.mkdir('archivosJson')

    def primeraFila(self):
        primFila = ["Nombre del Emisor","NIT Emisor","NRC","Nombre Comercial","SubTotal"]
        for file in primFila:
            col = 0
            for item in primFila:
                self.worksheet.write(0, col, primFila[col])
                col += 1

    def create_list_from_json(self,jsonfile):
        with open(jsonfile, 'r', encoding='utf-8-sig', errors='replace') as f:
            data = json.load(f)
        
        data_list = []

        if len(data) > 0:
            data_list.append(data["emisor"]["nombre"])
            data_list.append(data["emisor"]["nit"])
            data_list.append(data["emisor"]["nrc"])
            data_list.append(data["emisor"]["nombreComercial"])
            try:
                data_list.append(data["cuerpoDocumento"]["subTotal"])
            except:
                data_list.append(data["resumen"]["subTotal"])
            return data_list
        else:
            data_list.append("0000-0")
            return data_list


    
if __name__ == '__main__':
    app = App()
    app.mainloop()