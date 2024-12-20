import tkinter as tk
from tkinter import filedialog, messagebox, Toplevel, PhotoImage
import os
import shutil
import fiona
from functions.auxiliary import (get_list_ncobert, get_list_ncobert_excel, extract_zip_gdb, delete_folder, 
                                 validate_gdb_folder, is_number, truncar_string, list_files)
from functions.process import generate_result
from functions.root import ROOT_GDB, ROOT_IMG, ROOT_RESULT 
from functions.text_help import HELP_UPLODAD, HELP_TESTUDENT, HELP_OBSERVACIÓN1, HELP_OBSERVACIÓN2, HELP_OBSERVACIÓN3


class InterfazAplicacion:
    def __init__(self):
        self.root = None
        self.frame_campos = None
        self.upload_button = None
        self.logo_image = None
        self.grouped = None
        self.usar_mismo_ha = None
        self.usar_mismo_testudent = None
        self.procesar_button = None
        self.gdb_path = None
        self.processing = False
        self.type_origin = None

        self.create_interface()

    def create_interface(self):
        """
        Creates the graphical interface for uploading ZIP files.

        This method initializes a main window using Tkinter, allows the user to upload 
        a .gdb file, and provides a help label. It ensures that the logo image is kept 
        in memory and centers the window on the screen.

        Steps performed in the method:
        1. Deletes the specified GDB folder in the constant `ROOT_GDB`.
        2. Creates the main window with the title "Upload ZIP File".
        3. Loads and displays the logo image.
        4. Creates a button to allow the user to upload a file and links it to the 
        `upload_file2` function.
        5. Displays a help message in the interface.
        6. Centers the window on the monitor.
        7. Starts the main graphical interface loop.

        Attributes:
            self.root (Tk): The main window of the graphical interface.
            self.logo_image (PhotoImage): The logo image to be displayed in the interface.
            self.upload_button (Button): The button that allows the user to upload files.
        
        Related methods:
            - self.upload_file2: Method executed when the upload button is clicked.
            - self.center_window_on_monitor: Method that centers the window on the monitor.
        """

        delete_folder(ROOT_GDB)

        # Crear la ventana principal
        self.root = tk.Tk()
        self.root.title("Subir Archivo ZIP")

        # # Cargar la imagen del logo
        self.logo_image = PhotoImage(file=f"{ROOT_IMG}logo-anla.png")
        # self.logo_image2 = PhotoImage(file=f"{ROOT_IMG}logo_cm.png")

        # Crear un Frame para contener ambas imágenes en una fila
        logos_frame = tk.Frame(self.root)
        logos_frame.pack(pady=10)  # Añadir espacio superior e inferior al Frame

        # Crear el primer Label para el primer logo y agregarlo al Frame
        logo_label = tk.Label(logos_frame, image=self.logo_image)
        logo_label.pack(side=tk.LEFT, padx=5)  # Añade espacio horizontal

        # # Crear el segundo Label para el segundo logo y agregarlo al Frame
        # logo_label2 = tk.Label(logos_frame, image=self.logo_image2)
        # logo_label2.pack(side=tk.LEFT, padx=5)  # Añade espacio horizontal

        # Mantener referencias a las imágenes para que no se eliminen
        self.root.logo_image = self.logo_image
        # self.root.logo_image2 = self.logo_image2
        
        # Crear un botón para subir el archivo
        self.upload_button = tk.Button(self.root, text="Subir carpeta (gdb)", command=self.upload_file2)
        self.upload_button.pack(padx=50, pady=5)
        self.upload_button2 = tk.Button(self.root, text="Subir archivo excel (xlsx)", command=self.upload_file_xlsx)
        self.upload_button2.pack(padx=50, pady=5)


        ayuda_label = tk.Label(self.root, text=HELP_UPLODAD)
        ayuda_label.pack()

        self.root.update()

        # Centramos la ventana
        self.center_window_on_monitor(self.root)

        self.root.update()
        
        # Configurar el protocolo de cierre de ventana
        self.root.protocol("WM_DELETE_WINDOW", self.on_close)

        # Iniciar el bucle de la interfaz
        self.root.mainloop()


    def on_close(self):
        """
        Prevents the user from closing the window with the X if something is being processed.
        """

        if self.processing:
            # Si el proceso está en marcha, muestra un mensaje de advertencia
            messagebox.showwarning("Advertencia", "Por favor, haga click en Guardar o Cancelar.")
        else:
            # Si no hay ningún proceso en curso, permite cerrar la ventana
            self.root.destroy()

    def center_window_on_monitor(self,root):
        """
        Centers the provided window on the monitor screen.

        This method calculates the dimensions of the window and the screen
        to position the window at the center of the monitor. It ensures
        that all window elements are rendered before calculating the dimensions.

        Parameters:
            root (Tk): The Tkinter application window to be centered.

        Steps performed in the method:
        1. Calls `update_idletasks()` to ensure all window elements have been rendered.
        2. Gets the current dimensions of the window.
        3. Retrieves the dimensions of the main screen.
        4. Calculates the (x, y) coordinates needed to center the window.
        5. Sets the window geometry using the calculated coordinates.
        """

        root.update_idletasks()  # Asegurarse de que todos los elementos se han renderizado

        # Obtener el tamaño actual de la ventana después de renderizar los widgets
        ancho_ventana = root.winfo_width()
        alto_ventana = root.winfo_height()

        # Obtener las dimensiones de la pantalla principal
        ancho_pantalla = root.winfo_screenwidth()
        alto_pantalla = root.winfo_screenheight()

        # Calcular las coordenadas para centrar la ventana en la pantalla principal
        x = (ancho_pantalla // 2) - (ancho_ventana // 2)
        y = (alto_pantalla // 2) - (alto_ventana // 2)

        # Establecer la geometría de la ventana para centrarla
        root.geometry(f"{ancho_ventana}x{alto_ventana}+{x}+{y}")


    def upload_file(self):
        """
        Opens a dialog to select a ZIP file (gdb) and processes the selected file.

        This method allows the user to select a ZIP file containing geodatabase data.
        After selecting the file, it is extracted into a specific folder, and input fields 
        are generated based on the unique 'N_COBERT' values extracted from the file. 
        During the process, a popup window is shown to indicate that the file is being processed.

        Steps performed in the method:
        1. Opens a dialog for the user to select a ZIP file.
        2. If a file is selected, a popup window is created indicating that the processing is in progress.
        3. The popup window is configured as modal to prevent interactions with the main window during processing.
        4. The ZIP file is extracted into the 'fixed/gdb' folder.
        5. If the extraction is successful, the unique 'N_COBERT' values are retrieved.
        6. If unique values are found, input fields are generated.
        7. If no unique values are found or if an error occurs during extraction, appropriate error messages are displayed.
        8. If no file is selected, a warning is shown.
        """

        self.type_origin = "gdb"
        # Abrir un cuadro de diálogo para seleccionar un archivo
        file_path = filedialog.askopenfilename(title="Seleccionar archivo ZIP (gdb)", filetypes=[("ZIP files", "*.zip")])

        if file_path:
            # Crear ventana emergente de "Procesando archivo..."
            processing_window = Toplevel()
            processing_window.title("Procesando")
            tk.Label(processing_window, text="Procesando archivos...").pack(padx=20, pady=10)
            
            # Hacer que la ventana sea modal para evitar que se interactúe con la ventana principal
            processing_window.grab_set()

            # Actualizar la ventana para asegurarse de que se muestre completamente
            processing_window.update()

            # Centramos la ventana
            self.center_window_on_monitor(processing_window)

            # Actualizar la ventana para asegurarse de que se muestre completamente
            processing_window.update()

            # Descomprimir el archivo en la carpeta 'fixed/gdb' después de asegurarnos de que la ventana se muestre
            gdb_path = extract_zip_gdb(file_path, ROOT_GDB)

            if gdb_path:
                # Cerrar ventana emergente de "Procesando archivo..."
                processing_window.destroy()

                # Obtener la lista de valores únicos de 'N_COBERT'
                if "PuntoMuestreoFlora" in fiona.listlayers(gdb_path):
                    print(f"Hay capa")
                    lista_unicos_n_cobert = get_list_ncobert(gdb_path)
                    if lista_unicos_n_cobert:
                        # Generar los campos de entrada debajo del botón
                        self.create_entry_fields(lista_unicos_n_cobert)
                    else:
                        processing_window.destroy()
                        messagebox.showerror("Error", "No existen N_COBERT.")
                else:
                    processing_window.destroy()
                    messagebox.showerror("Error", "No existe la capa 'PuntoMuestreoFlora")
                
            else:
                processing_window.destroy()
                messagebox.showerror("Error", "No se pudo descomprimir el archivo o no se encontró la geodatabase.")
        else:
            messagebox.showwarning("Sin selección", "No se ha seleccionado ningún archivo")

    def upload_file2(self):
        """
        Opens a dialog to select a folder containing a geodatabase (GDB)
        and processes the contents of the selected folder.

        This method allows the user to select a folder containing a geodatabase.
        It verifies that the selected folder is valid and, if so, copies its contents 
        to a specific directory. After the copy, the unique 'N_COBERT' values are extracted 
        from the geodatabase to generate input fields in the interface.

        Steps performed in the method:
        1. Opens a dialog for the user to select a folder containing the GDB.
        2. If a folder is selected, its contents are validated.
        3. A popup window is created to indicate that the selected folder is being processed.
        4. The popup window is centered and updated to display correctly.
        5. If the folder is valid, the destination path in 'ROOT_GDB' is defined and the 
           selected folder is copied to that location.
        6. Unique 'N_COBERT' values are extracted from the GDB.
        7. If unique values are found, input fields are generated and the processing popup is destroyed.
        8. If no unique values are found, an error message is displayed.
        9. If the selected folder is not valid, an error message is shown and the interface is recreated.
        10. If no folder is selected, a warning is displayed.
        """

        self.type_origin = "gdb"

        # Abrir un cuadro de diálogo para seleccionar un archivo
        file_path = filedialog.askdirectory(title="Seleccionar carpeta (gdb)")

        if file_path:
            if validate_gdb_folder(file_path):
                file_gdb = os.path.basename(file_path)
                self.gdb_path = f"{ROOT_GDB}{file_gdb}"
                # Crear ventana emergente de "Procesando archivo..."
                processing_window = Toplevel()
                processing_window.title("Procesando")
                tk.Label(processing_window, text="Procesando archivos...").pack(padx=20, pady=10)
                
                # Hacer que la ventana sea modal para evitar que se interactúe con la ventana principal
                processing_window.grab_set()

                # Actualizar la ventana para asegurarse de que se muestre completamente
                processing_window.update()

                # Centramos la ventana
                self.center_window_on_monitor(processing_window)

                # Actualizar la ventana para asegurarse de que se muestre completamente
                processing_window.update()

                        # Verificar si se seleccionó una carpeta
                if file_path:
                    # Definir el destino en ROOT_GDB
                    dest_path = os.path.join(ROOT_GDB, os.path.basename(file_path))
                    
                    # Verificar si la carpeta ya existe en el destino
                    if not os.path.exists(dest_path):
                        # Copiar la carpeta completa al destino
                        shutil.copytree(file_path, dest_path)
                        # print(f"Carpeta copiada exitosamente a {dest_path}")
                    else:
                        print(f"La carpeta ya existe en {dest_path}")

                # Obtener la lista de valores únicos de 'N_COBERT'
                if "PuntoMuestreoFlora" in fiona.listlayers(file_path):
                    # Obtener la lista de valores únicos de 'N_COBERT'
                    lista_unicos_n_cobert = get_list_ncobert(file_path)
                    if lista_unicos_n_cobert:
                        # Generar los campos de entrada debajo del botón
                        self.create_entry_fields(lista_unicos_n_cobert)
                        self.create_frame_buttons()
                        processing_window.destroy()
                    else:
                        messagebox.showerror("Error", "No existen N_COBERT en la GDB.")
                        processing_window.destroy()
                else:
                    processing_window.destroy()
                    messagebox.showerror("Error", "No existe la capa 'PuntoMuestreoFlora")
                
            else:
                # Mostrar mensaje de error
                messagebox.showerror("Error", "Carpeta GDB no válida. Seleccione una carpeta GDB correcta.")
                self.root.destroy()
                # Llamar a la función para crear de nuevo la interfaz
                self.create_interface()

        else:
            messagebox.showwarning("Sin selección", "No se ha seleccionado ninguna carpeta.")

    def upload_file_xlsx(self):
        """
        Prompts the user to select an Excel (.xlsx) file, validates the selection,
        and processes the file by copying it to a specific directory. A modal window
        is displayed during processing, and based on the data extracted from the file,
        appropriate input fields and buttons are created.

        The function performs the following tasks:
        1. Opens a file dialog for the user to select an Excel file (.xlsx).
        2. Verifies that the selected file is of type .xlsx.
        3. Copies the selected file to a specified directory (ROOT_GDB).
        4. Displays a modal "Processing" window to indicate ongoing processing.
        5. Extracts unique coverage data (N_COBERT) from the Excel file.
        6. If valid data is found, generates user input fields and buttons.
        7. Displays an error message if no valid data is found or if the file is invalid.
        
        If the file is successfully processed, input fields and buttons for further interaction
        are created. If no valid data is found in the Excel file, an error message is displayed.
        
        The modal processing window is closed once the operation is complete.

        Returns:
            None
        """
        self.type_origin = "xlsx"
        file_path = filedialog.askopenfilename(
        title="Seleccionar archivo Excel (xlsx)",
        filetypes=[("Archivos Excel", "*.xlsx")],  # Mostrar solo archivos .xlsx
        )
        if file_path and file_path.endswith(".xlsx"):
            file_gdb = os.path.basename(file_path)
            self.gdb_path = f"{ROOT_GDB}{file_gdb}"
            # Copiar el archivo al directorio de resultados
            try:
                shutil.copy(file_path, ROOT_GDB)
            except Exception as e:
                print(f"Error al guardar el archivo: {e}")
            # Crear ventana emergente de "Procesando archivo..."
            processing_window = Toplevel()
            processing_window.title("Procesando")
            tk.Label(processing_window, text="Procesando archivos...").pack(padx=20, pady=10)          
            # Hacer que la ventana sea modal para evitar que se interactúe con la ventana principal
            processing_window.grab_set()
            # Actualizar la ventana para asegurarse de que se muestre completamente
            processing_window.update()
            # Centramos la ventana
            self.center_window_on_monitor(processing_window)
            # Actualizar la ventana para asegurarse de que se muestre completamente
            processing_window.update()
            lista_unicos_n_cobert = get_list_ncobert_excel(file_path)
            if lista_unicos_n_cobert:
                # Generar los campos de entrada debajo del botón
                self.create_entry_fields(lista_unicos_n_cobert)
                self.create_frame_buttons()
                processing_window.destroy()
            else:
                messagebox.showerror("Error", "No existen coberturas en el Excel")
                processing_window.destroy()
            
        else:
            if file_path:  # Si el usuario seleccionó un archivo pero no es .xlsx
                print("El archivo seleccionado no es un archivo Excel válido (.xlsx).")
            else:  # Si no seleccionó nada
                print("No se seleccionó ningún archivo.")

    def create_entry_fields(self, lista_n_cobert):
        """
        Creates input fields in the interface for each value in the 'lista_n_cobert' list.

        This method generates labels and input fields for the elements of 'lista_n_cobert'.
        The fields are organized in a scrollable frame that allows the user 
        to enter data associated with each 'N_COBERT'. 

        The interface allows showing up to 10 items without scrolling, 
        and if there are more, a vertical scrollbar is enabled for easy navigation.

        Steps performed in the method:
        1. Disables the upload button to prevent concurrent actions.
        2. Initializes lists to store the input fields.
        3. Creates a frame for the canvas and a canvas to contain the fields.
        4. Creates a scrollable frame within the canvas.
        5. Iterates over 'lista_n_cobert' to generate labels and input fields.
        6. If there are 10 or more items in 'lista_n_cobert', a scrollbar is added.
        7. Configures the scroll properties of the canvas.

        Parameters:
            lista_n_cobert (list): List of values that will be used to create labels and input fields.
        """


        self.upload_button.config(state=tk.DISABLED)
        self.upload_button2.config(state=tk.DISABLED)
        self.entry_list = []  # Lista para almacenar los campos de entrada
        self.entry2_list = [] 

        WIDTH_CANVAS = 500
        HEIGHT_CANVAS = 280

        # Crear frame Canvas
        frame_canvas = tk.Frame(self.root)  # Cambia 'self' por el marco superior si corresponde
        frame_canvas.pack(pady=20)

        # Crear el Canvas
        canvas = tk.Canvas(frame_canvas, width=WIDTH_CANVAS, height=HEIGHT_CANVAS)
        canvas.grid(row=0, column=0, sticky="nsew")

        # Crear el scrollable_frame dentro del canvas
        self.scrollable_frame = tk.Frame(canvas)
        canvas.create_window((0, 0), window=self.scrollable_frame, anchor="nw")  # Ancla la ventana en la esquina superior izquierda

        # Añadir etiquetas y campos de entrada en dos columnas
        for i, valor in enumerate(lista_n_cobert):
            valor = truncar_string(valor, 40)
            etiqueta = tk.Label(self.scrollable_frame, text=f"{valor} ")
            etiqueta.grid(row=i, column=0, sticky="w", padx=(5, 5), pady=5)  

            etiqueta2 = tk.Label(self.scrollable_frame, text="ha:")
            etiqueta2.grid(row=i, column=1, sticky="w", padx=(1, 1), pady=5)  

            entry = tk.Entry(self.scrollable_frame, width=8)
            entry.grid(row=i, column=2, sticky="w", padx=(5, 5), pady=5)  

            etiqueta3 = tk.Label(self.scrollable_frame, text="Tstudent:")
            etiqueta3.grid(row=i, column=3, sticky="w", padx=(1, 1), pady=5)  

            entry2 = tk.Entry(self.scrollable_frame, width=8)
            entry2.grid(row=i, column=4, sticky="w", padx=(5, 5), pady=5)  

            # Añadir los campos de entrada a las listas
            self.entry_list.append(entry)
            self.entry2_list.append(entry2)

        # print(f"len(lista_n_cobert):{len(lista_n_cobert)}")
        if len(lista_n_cobert) >= 8:

             # Crear el Scrollbar en el mismo contenedor que el Canvas
            scrollbar = tk.Scrollbar(frame_canvas, orient="vertical", command=canvas.yview)
            scrollbar.grid(row=0, column=1, sticky="ns")  # Coloca el scrollbar a la derecha del canvas
            # Configurar el scroll del canvas con el scrollbar
            canvas.configure(yscrollcommand=scrollbar.set)
            self.scrollable_frame.bind("<Configure>", lambda e: canvas.configure(scrollregion=canvas.bbox("all")))
        else:
            self.scrollable_frame.bind("<Configure>", lambda e: canvas.configure(scrollregion=canvas.bbox("all")))

    def create_frame_buttons(self):
        """
        Creates a frame containing checkboxes and a button to process data.

        This method sets up a new frame in the user interface that contains 
        several checkboxes and a "Process" button. The checkboxes allow the user 
        to select different options for data processing, and the button invokes the 
        corresponding method to carry out the processing action.

        Steps performed in the method:
        1. Creates an additional frame (`frame_opciones`) that is placed below the input fields.
        2. Calls specific methods to create several checkboxes that allow the user 
           to select additional options.
        3. Creates a button that, when pressed, executes the `process_data` method.
        4. Adjusts the size of the main window to fit the content and centers it on the monitor.
        """

        
        # Crear un Frame adicional debajo para los checkboxes y el botón
        self.frame_grouped = tk.Frame(self.root)
        self.frame_grouped.pack(side="top", fill="x", pady=10)

        ayuda_label2 = tk.Label(self.frame_grouped, text=HELP_TESTUDENT, justify="left", anchor="w", width=80, font=("Helvetica", 10))
        ayuda_label2.pack(padx=25,pady=4)
        ayuda_label3 = tk.Label(self.frame_grouped, text=HELP_OBSERVACIÓN1, justify="left", anchor="w", width=80, font=("Helvetica", 10))
        ayuda_label3.pack(padx=25,pady=4)
        ayuda_label4 = tk.Label(self.frame_grouped, text=HELP_OBSERVACIÓN2, justify="left", anchor="w", width=80, font=("Helvetica", 10))
        ayuda_label4.pack(padx=25,pady=4)
        ayuda_label5 = tk.Label(self.frame_grouped, text=HELP_OBSERVACIÓN3, justify="left", anchor="w", width=80, font=("Helvetica", 10))
        ayuda_label5.pack(padx=25,pady=4)

        # Crear un checkbox para usar el mismo ha en todos los campos
        self.grouped = tk.BooleanVar()
        self.check_grouped = tk.Checkbutton(self.frame_grouped , text="Agrupado por ecosistema", variable=self.grouped, command=lambda: self.apply_same_ha(self.usar_mismo_ha))
        self.check_grouped.pack(pady=3)

        # Crear un Frame adicional debajo para los checkboxes y el botón
        self.frame_opciones = tk.Frame(self.root)
        self.frame_opciones.pack(side="top", fill="x", pady=3)


        # Añadir los checkboxes y el botón a este Frame independiente
        self.create_checkboxes()
        self.procesar_button = tk.Button(self.frame_opciones, text="Procesar", command=self.process_data)
        self.procesar_button.grid(row=1, column=0, columnspan=4, pady=(10, 5))

        # Ajustar tamaño de ventana y centrar
        self.root.update_idletasks()
        self.root.geometry("")  # Permitir que la ventana ajuste su tamaño
        self.center_window_on_monitor(self.root)

    def create_checkboxes(self):
        """
        Creates multiple checkboxes in the user interface to allow the user to select options that affect data input.

        This method sets up several checkboxes that give the user the ability to apply uniform settings 
        to the input fields in the interface. Each checkbox is linked to a boolean variable and triggers 
        a specific function when checked or unchecked.

        The created checkboxes are:
        1. "mismo ha" (same area): Applies the same area (ha) value to all input fields.
        2. "mismo testudent" (same Tstudent): Applies the same Tstudent value to all input fields.
        3. "ha desde gdb" (area from GDB): Uses area values from a GDB file.
        4. "Tstudent calculado" (calculated Tstudent): Uses calculated Tstudent values from a file.

        Steps performed in the method:
        1. Defines four boolean variables to control the checkbox states.
        2. Creates each checkbox with its corresponding label and links its state to the 
           corresponding variable.
        3. Sets a command for each checkbox to be executed when its state changes, 
           invoking the corresponding method.
        4. Places the checkboxes in the options frame in their respective positions.
        """

        
         # Configurar las columnas de la grilla
        total_columns = 4
        for i in range(total_columns):
            self.frame_opciones.grid_columnconfigure(i, weight=1)  # Igualar pesos para centrar

        # Checkbox 1
        self.usar_mismo_ha = tk.BooleanVar()
        self.check_button_ha = tk.Checkbutton(
            self.frame_opciones,
            text="Misma área de muestreo",
            variable=self.usar_mismo_ha,
            anchor="w",
            justify="left",
            command=lambda: self.apply_same_ha(self.usar_mismo_ha)
        )
        self.check_button_ha.grid(row=0, column=0, padx=10, pady=5, sticky="ew")

        # Checkbox 2
        self.usar_mismo_testudent = tk.BooleanVar()
        self.check_button_testudent = tk.Checkbutton(
            self.frame_opciones,
            text="Misma tstudent tipo\nde cobert",
            variable=self.usar_mismo_testudent,
            anchor="w",
            justify="left",
            command=lambda: self.apply_same_testudent(self.usar_mismo_testudent)
        )
        self.check_button_testudent.grid(row=0, column=1, padx=10, pady=5, sticky="ew")

        # Checkbox 3
        self.usar_desde_archivo = tk.BooleanVar(value=False)
        check_button_archivo = tk.Checkbutton(
            self.frame_opciones,
            text="Usar área de muestreo\nde la gdb",
            variable=self.usar_desde_archivo,
            anchor="w",
            justify="left",
            command=lambda: self.use_same_ha_file(self.usar_desde_archivo)
        )
        check_button_archivo.grid(row=0, column=2, padx=10, pady=5, sticky="ew")
        if self.type_origin == "xlsx":
            check_button_archivo.config(state='disabled')

        # Checkbox 4
        self.usar_desde_archivo2 = tk.BooleanVar()
        check_button_archivo2 = tk.Checkbutton(
            self.frame_opciones,
            text="Calcular Tstudent\n1 cola",
            variable=self.usar_desde_archivo2,
            anchor="w",
            justify="left",
            command=lambda: self.use_same_testudent_file(self.usar_desde_archivo2)
        )
        check_button_archivo2.grid(row=0, column=3, padx=10, pady=5, sticky="ew")

    def use_same_ha_file(self,usar_desde_archivo):
        """
        Manages the logic for using an input file for the area (ha) value based on the state of the corresponding checkbox.

        This method is called when the user interacts with the "ha desde gdb" checkbox. 
        If the checkbox is checked, all input fields are disabled and cleared, 
        and the "mismo ha" checkbox is disabled. If the checkbox is unchecked, 
        the input fields are re-enabled and the "mismo ha" checkbox is re-enabled 
        if it was still selected.

        Parameters:
        ----------
        usar_desde_archivo : BooleanVar
            A variable that represents the state of the "ha desde gdb" checkbox. 
            Returns True if the checkbox is checked, False if unchecked.

        Method Logic:
        - If the "ha desde gdb" checkbox is checked:
            1. Input fields are enabled so the user can interact with them.
            2. All input fields are cleared.
            3. All input fields are disabled to prevent accidental data modification 
            while the file is in use.
            4. The "mismo ha" checkbox is disabled and unchecked.
        - If the "ha desde gdb" checkbox is unchecked:
            1. All input fields are re-enabled.
            2. The "mismo ha" checkbox is re-enabled.
            3. If "mismo ha" is still checked, the logic for setting the same value 
            across all input fields is reapplied.
        """

        
        # Si el checkbox de usar archivo está marcado
        if usar_desde_archivo.get():
            # Vaciar y desactivar todos los campos
            for entry in self.entry_list[1:]:
                    entry.config(state="normal")
                    
            for entry in self.entry_list:
                entry.delete(0, tk.END)  # Borrar cualquier valor previo
                entry.insert(0, "")
                entry.config(state="disabled")  # Desactivar todos los campos
          
            # Desactivar el checkbox de "Usar el mismo ha"
            self.check_button_ha.config(state="disabled")
            self.usar_mismo_ha.set(False)  # Desmarcar el checkbox

        else:
            # Si se desmarca el checkbox de usar archivo, activar los campos nuevamente
            for entry in self.entry_list:
                entry.config(state="normal")
           
            # Reactivar el checkbox de "Usar el mismo ha"
            self.check_button_ha.config(state="normal")
            
            # Si además está marcado el checkbox de usar el mismo ha, volver a aplicar esa lógica
            if self.usar_mismo_ha.get():
                self.apply_same_ha(self.usar_mismo_ha)

    def use_same_testudent_file(self,usar_desde_archivo2):
        """
        Manages the logic for using an input file for the Tstudent value based on the state of the corresponding checkbox.

        This method is called when the user interacts with the "Tstudent calculado" checkbox. 
        If the checkbox is checked, all Tstudent input fields are disabled and cleared, 
        and the "mismo Tstudent" checkbox is disabled. If the checkbox is unchecked, 
        the Tstudent input fields are re-enabled and the "mismo Tstudent" checkbox is re-enabled 
        if it was still selected.

        Parameters:
        ----------
        usar_desde_archivo2 : BooleanVar
            A variable that represents the state of the "Tstudent calculado" checkbox. 
            Returns True if the checkbox is checked, False if unchecked.

        Method Logic:
        - If the "Tstudent calculado" checkbox is checked:
            1. Tstudent input fields are enabled so the user can interact with them.
            2. All Tstudent input fields are cleared.
            3. All Tstudent input fields are disabled to prevent accidental data modification 
            while the file is in use.
            4. The "mismo Tstudent" checkbox is disabled and unchecked.
        - If the "Tstudent calculado" checkbox is unchecked:
            1. All Tstudent input fields are re-enabled.
            2. The "mismo Tstudent" checkbox is re-enabled.
            3. If "mismo Tstudent" is still checked, the logic for setting the same value 
            across all Tstudent input fields is reapplied.
        """

        
        # Si el checkbox de usar archivo está marcado
        if usar_desde_archivo2.get():
            # Vaciar y desactivar todos los campos

            for entry2 in self.entry2_list[1:]:
                    entry2.config(state="normal")
                    
            for entry2 in self.entry2_list:
                entry2.delete(0, tk.END)  # Borrar cualquier valor previo
                entry2.insert(0, "")
                entry2.config(state="disabled")  # Desactivar todos los campos
            
            # Desactivar el checkbox de "Usar el mismo ha"
            self.check_button_testudent.config(state="disabled")
            self.usar_mismo_testudent.set(False)  # Desmarcar el checkbox

        else:
            # Si se desmarca el checkbox de usar archivo, activar los campos nuevamente
            for entry2 in self.entry2_list:
                entry2.config(state="normal")
            
            # Reactivar el checkbox de "Usar el mismo ha"
            self.check_button_testudent.config(state="normal")
            
            # Si además está marcado el checkbox de usar el mismo ha, volver a aplicar esa lógica
            if self.usar_mismo_testudent.get():
                self.apply_same_testudent(self.usar_mismo_testudent)

    def apply_same_ha(self, usar_mismo_ha):
        """
        Applies the same "ha" value to all input fields based on the state of the corresponding checkbox.

        This method is called when the user interacts with the "mismo ha" checkbox. 
        If the checkbox is checked and the "usar archivo" checkbox is not active, 
        the value from the first input field is taken and applied to all the other fields. 
        If the checkbox is unchecked or if the "usar archivo" checkbox is active, 
        the input fields are re-enabled.

        Parameters:
        ----------
        usar_mismo_ha : BooleanVar
            A variable representing the state of the "mismo ha" checkbox. 
            Returns True if the checkbox is checked, False if unchecked.

        Method Logic:
        - If the "mismo ha" checkbox is checked and the "usar archivo" checkbox is not active:
            1. The value from the first input field is retrieved.
            2. All other input fields are iterated over, their values are cleared, 
            and the same value as the first field is set for them.
            3. All input fields are disabled to prevent accidental modification.
        - If the "mismo ha" checkbox is unchecked or the "usar archivo" checkbox is active:
            1. It is checked if the "usar archivo" checkbox is not active.
            2. All input fields, except the first one, are re-enabled.
        """

        
        # Si el checkbox está marcado y no está activo el checkbox de usar archivo
        if usar_mismo_ha.get() and not self.usar_desde_archivo.get():
            primer_valor = self.entry_list[0].get()  # Obtener el valor del primer campo
            for entry in self.entry_list[1:]:
                entry.delete(0, tk.END)  # Borrar cualquier valor previo
                entry.insert(0, primer_valor)  # Poner el valor del primer campo en los demás
                entry.config(state="disabled")  # Desactivar el campo

        else:
            # Si se desmarca el checkbox o está activo el de usar archivo, activar los campos nuevamente
            if not self.usar_desde_archivo.get():
                for entry in self.entry_list[1:]:
                    entry.config(state="normal")


    def apply_same_testudent(self, usar_mismo_student):
        """
        Applies the same "Tstudent" value to all input fields based on the state of the corresponding checkbox.

        This method is called when the user interacts with the "mismo Tstudent" checkbox. 
        If the checkbox is checked and the "usar archivo" checkbox is not active, 
        the value from the first input field is taken and applied to all the other fields. 
        If the checkbox is unchecked or if the "usar archivo" checkbox is active, 
        the input fields are re-enabled.

        Parameters:
        ----------
        usar_mismo_student : BooleanVar
            A variable representing the state of the "mismo Tstudent" checkbox. 
            Returns True if the checkbox is checked, False if unchecked.

        Method Logic:
        - If the "mismo Tstudent" checkbox is checked and the "usar archivo" checkbox is not active:
            1. The value from the first input field is retrieved.
            2. All other input fields are iterated over, their values are cleared, 
            and the same value as the first field is set for them.
            3. All input fields are disabled to prevent accidental modification.
        - If the "mismo Tstudent" checkbox is unchecked or the "usar archivo" checkbox is active:
            1. It is checked if the "usar archivo" checkbox is not active.
            2. All input fields, except the first one, are re-enabled.
        """
        
        # Si el checkbox está marcado y no está activo el checkbox de usar archivo
        if usar_mismo_student.get() and not self.usar_desde_archivo2.get():
            primer_valor2 = self.entry2_list[0].get()  # Obtener el valor del primer campo
            for entry2 in self.entry2_list[1:]:
                entry2.delete(0, tk.END)  # Borrar cualquier valor previo
                entry2.insert(0, primer_valor2)  # Poner el valor del primer campo en los demás
                entry2.config(state="disabled")  # Desactivar el campo
        else:
            # Si se desmarca el checkbox o está activo el de usar archivo, activar los campos nuevamente
            if not self.usar_desde_archivo2.get():
                for entry2 in self.entry2_list[1:]:
                    entry2.config(state="normal")

    def process_data(self):
        """
        Processes the entered data and generates a result file.

        This method validates the input fields, and if they are valid, it displays 
        a pop-up window indicating that the information is being processed. 
        Depending on the states of the checkboxes, it collects the values for HA and 
        Tstudent from the input fields or uses predefined values. It then calls 
        the `generate_result` function to create the result file.

        Method Logic:
        --------------
        1. Validates the entered numbers using the `validate_numbers` method.
        2. Creates a pop-up window to indicate that the information is being processed.
        3. Disables the "Process" button to prevent multiple invocations.
        4. Depending on the checkbox states:
            - If a file is used, assigns -1 to the HA or Tstudent values list.
            - If the same value is used, repeats the first entered value to fill the list.
            - Otherwise, the values are obtained from the input fields.
        5. Calls `generate_result` with the GDB file path and the collected values.
        6. If the generation is successful:
            - Closes the pop-up window.
            - Displays a button to return to the initial process.
            - Adjusts the geometry of the main window.
            - Saves the result file.
        7. If an error occurs during generation, displays an error message and restarts the process.
        """

    
        # Validar los campos antes de procesar
        if self.validate_numbers():
            
            # Crear ventana emergente de "Procesando Información..."
            processing_window = Toplevel()
            processing_window.title("Procesando")
            tk.Label(processing_window, text="Procesando Información...").pack(padx=20, pady=10)
            
            # Hacer que la ventana sea modal para evitar que se interactúe con la ventana principal
            processing_window.grab_set()

            # Actualizar la ventana para asegurarse de que se muestre completamente
            processing_window.update()

            # Centramos la ventana
            self.center_window_on_monitor(processing_window)

            # Actualizar la ventana para asegurarse de que se muestre completamente
            processing_window.update()
            # Desactivar el botón "Procesar"
            self.procesar_button.config(state="disabled")

            if self.grouped.get():
                boo_grouped = True
            else:
                boo_grouped = False

            # Verificar si el checkbox está marcado
            if self.usar_desde_archivo.get():
                ha_values = [-1]
            else:
                if self.usar_mismo_ha.get():
                    # Obtener el valor del primer campo y repetirlo para completar la lista
                    primer_valor = float(self.entry_list[0].get())
                    ha_values = [primer_valor] * len(self.entry_list)

                else:
                    # Obtener los valores ingresados en los campos de entrada
                    ha_values = [float(entry.get()) for entry in self.entry_list]

            # Verificar si el checkbox está marcado
            if self.usar_desde_archivo2.get():
                t_student_values = [-1]
            else:
                if self.usar_mismo_testudent.get():
                    # Obtener el valor del primer campo y repetirlo para completar la lista
                    primer_valor2 = float(self.entry2_list[0].get())
                    t_student_values = [primer_valor2] * len(self.entry2_list)
                else:
                    # Obtener los valores ingresados en los campos de entrada
                    t_student_values = [float(entry2.get()) for entry2 in self.entry2_list]

            # Llamar a la función para generar el archivo resultante
            result, message = generate_result(self.gdb_path, boo_grouped, ha_values, t_student_values, ROOT_RESULT, self.type_origin)

            if result:
                # Cerrar ventana emergente de "Procesando archivo..."
                processing_window.destroy()
                delete_folder(ROOT_GDB)
                # Si la generación fue exitosa, abrir cuadro de diálogo para guardar el archivo
                # Mostrar botón "Volver"
                volver_button = tk.Button(self.frame_opciones, text="Volver", command=self.restart_process)
                volver_button.grid(row=2, column=0, columnspan=4, pady=(10, 5))

                # Asegurarse de que la ventana se actualice
                self.root.update_idletasks()  
                
                # Ajustar la geometría de la ventana
                nuevo_ancho = self.root.winfo_width()
                nuevo_alto = self.root.winfo_height() + 50  # Ajustar según sea necesario
                
                self.root.geometry(f"{nuevo_ancho}x{nuevo_alto}")  # Usar formato de cadena
                self.center_window_on_monitor(self.root)

                self.save_file(ROOT_RESULT)  
                
            else:
                processing_window.destroy()
                messagebox.showerror("Error", message)
                delete_folder(ROOT_GDB)
                self.restart_process()


    def validate_numbers(self):
        """
        Validates the numbers entered in the input fields.

        This method checks that the input fields are not empty and that they contain
        only numeric values. The validation is based on the state of various checkboxes
        that determine whether all fields or just the first field should be validated.

        Method Logic:
        --------------
        1. If the `usar_desde_archivo` checkbox is checked, validation is skipped.
        2. If the `usar_mismo_ha` checkbox is checked:
            - Only the first field (na) is validated.
            - An error message is shown if the field is empty or non-numeric.
        3. If `usar_mismo_ha` is not checked, all fields (na) are validated:
            - Each field is checked to ensure it is not empty and contains only numbers.
            - An error message is shown if any field is empty or non-numeric.
        4. A similar process is repeated for the `Tstudent` fields:
            - If the `usar_desde_archivo2` checkbox is checked, validation is skipped.
            - If the `usar_mismo_testudent` checkbox is checked, only the first field (Tstudent) is validated.
            - If `usar_mismo_testudent` is not checked, all fields (testudent) are validated.
        5. Returns `True` if all validations are successful, otherwise returns `False`.

        Returns:
            bool: True if all fields are valid, False otherwise.
        """

        
        # Validar si el checkbox está marcado o no
        if self.usar_desde_archivo.get():
            pass
        else:
            if self.usar_mismo_ha.get():
                # Si está marcado, validar solo el primer campo
                primer_input = self.entry_list[0].get()
                if primer_input == "":
                    messagebox.showerror("Error", "El primer campo área de muestreo en (ha) no puede estar vacío.")
                    return False
                if not is_number(primer_input):
                    messagebox.showerror("Error", "El primer campo área de muestreo en (ha) debe contener solo números y el decimal es con punto (.).")
                    return False
            else:
                # Si no está marcado, validar todos los campos
                for entry in self.entry_list:
                    valor = entry.get()
                    if valor == "":
                        messagebox.showerror("Error", "Todos los campos área de muestreo en (ha) deben tener un valor.")
                        return False
                    if not is_number(valor):
                        messagebox.showerror("Error", "Todos los campos de área de muestreo en (ha) deben contener solo números y el decimal es con punto (.).")
                        return False
                    
        if self.usar_desde_archivo2.get():
            pass
        else:
            if self.usar_mismo_testudent.get():  
                primer_input2 = self.entry2_list[0].get()
                if primer_input2 == "":
                    messagebox.showerror("Error", "El primer campo (Tstudent) no puede estar vacío.")
                    return False
                if not is_number(primer_input2):
                    messagebox.showerror("Error", "El primer campo (Tstudent) debe contener solo números y el decimal es con punto (.).")
                    return False
            else:
                # Si no está marcado, validar todos los campos                  
                for entry2 in self.entry2_list:
                    valor2 = entry2.get()
                    if valor2 == "":
                        messagebox.showerror("Error", "Todos los campos (testudent) deben tener un valor.")
                        return False
                    if not is_number(valor2):
                        messagebox.showerror("Error", "Todos los campos (testudent) deben contener solo números y el decimal es con punto (.).")
                        return False
                    
        return True
    
    def restart_process(self):
        """
        Resets the application's process by closing the current window and
        creating a new user interface.

        This method is used to restore the application's state to its
        initial configuration. When this method is invoked, the current
        main window is destroyed, and the `create_interface` method is
        called to generate a new user interface from scratch.

        Method Logic:
        --------------
        1. The current window (`self.root`) is destroyed.
        2. The `create_interface` method is called to create a new
           instance of the user interface.

        Returns:
            None
        """

        
        self.root.destroy()
        self.create_interface()

    def upload_zip_button(self):
        """
        Displays the "Upload ZIP file" button in the user interface.

        This method is used to add the button that allows the user to
        select and upload a ZIP file from their local system. When this
        method is invoked, the button is presented in the interface,
        allowing the user to interact with it.

        Method Logic:
        --------------
        1. The `pack` method from Tkinter is used to add the upload button
           to the interface, with a vertical margin of 20 pixels (pady).

        Returns:
            None
        """

        
        # Reagregar el botón "Subir archivo ZIP"
        self.upload_button.pack(pady=20)

    def save_file(self,output_folder):
        """
        Opens a dialog box for the user to select the location 
        to save a generated file and saves the file at the selected 
        location.

        This method allows the user to save an Excel file, which 
        is specified as an argument. A dialog box is used to choose 
        the location and name of the file. If the operation is 
        successful, a confirmation message is shown; otherwise, 
        any errors encountered during the process are reported.

        Parameters:
        -----------
        output_file : str
            The path of the file to be saved.

        Method Logic:
        --------------
        1. A dialog box is opened to select the location and name of the 
           file. The filter is set to Excel files with the ".xlsx" extension.
        2. If the user selects a location and name, an attempt is made 
           to copy the generated file to that location.
        3. If the copy is successful, an informational message is displayed. 
           If an error occurs, an error message is shown. If the 
           user cancels the operation, a cancellation message is shown.

        Returns:
            None
        """

        self.processing = True
        
        lista_archivos_result = list_files(output_folder)

        if isinstance(lista_archivos_result, list) and len(lista_archivos_result) > 0:
            archivo = lista_archivos_result[0]  # Obtener el primer archivo de la lista
            extension = os.path.splitext(archivo)[1]

            ruta_result = archivo  # Ajusta la ruta según la estructura de tu proyecto
            file_path = filedialog.asksaveasfilename(defaultextension=extension, filetypes=[(f"{extension.upper()} files", f"*{extension}")])
        
            if file_path:
                try:
                    # Copiar el archivo result.zip a la nueva ubicación seleccionada por el usuario
                    print(ruta_result)
                    print(file_path)
                    shutil.copyfile(ruta_result, file_path)
                    print(f"Archivo guardado en: {file_path}")
                    # Aquí podrías mostrar un mensaje de éxito o hacer otras operaciones si es necesario
                except Exception as e:
                    print(f"Error al guardar el archivo: {e}")
            else:
                print("Guardado cancelado por el usuario")
            
            self.processing = False
        else:
            print("No se encontraron archivos para guardar")


# Para ejecutar la aplicación
if __name__ == "__main__":
    app = InterfazAplicacion()
