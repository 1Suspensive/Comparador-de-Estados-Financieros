import tkinter as tk
from tkinter import ttk, filedialog, messagebox, scrolledtext
import pandas as pd
import numpy as np
import re
import os
import ctypes  # Para el escalado High DPI en Windows

class ExcelComparator:
    """
    Contiene toda la lógica de negocio para procesar y comparar
    los archivos Excel. Es independiente de la GUI.
    """
    def __init__(self, cliente_path, salida_path, cliente_config):
        self.cliente_path = cliente_path
        self.salida_path = salida_path
        self.cliente_config = cliente_config
        self.inconsistencias = []

        # Configuración fija para el archivo de salida
        self.salida_config = {
            'title_range': 'C:H',
            'actual_col': 'I',
            'anterior_col': 'J'
        }
        
        # Mapeo de hojas (basado en índice 0-based de pandas)
        # Cliente Hoja 3 (idx 2) -> Salida Hoja 2 (idx 1)
        # Cliente Hoja 4 (idx 3) -> Salida Hoja 3 (idx 2)
        # Cliente Hoja 5 (idx 4) -> Salida Hoja 4 (idx 3)
        # self.sheet_map = [(2, 1), (3, 2), (4, 3)]
        self.sheet_map = [(2, 1)]
        

    def _col_to_int(self, col_str):
        """Convierte una letra de columna de Excel (A, B, AA) a un índice 0-based."""
        num = 0
        for c in col_str:
            num = num * 26 + (ord(c.upper()) - ord('A')) + 1
        return num - 1

    def _parse_col_range(self, range_str):
        """Convierte un rango (ej: "C:H") a una lista de índices [2, 3, 4, 5, 6, 7]."""
        try:
            start_col, end_col = range_str.split(':')
            start_idx = self._col_to_int(start_col)
            end_idx = self._col_to_int(end_col)
            return list(range(start_idx, end_idx + 1))
        except Exception:
            raise ValueError(f"Formato de rango de título inválido: '{range_str}'. Use 'A:F'.")

    def _normalize_title(self, title):
        """
        Normaliza un título. Busca *ESTRICTAMENTE* el formato 'Num Texto'.
        Si coincide, devuelve (numero, texto).
        Si no coincide, devuelve None.
        """
        if not isinstance(title, str):
            return None
        
        title = title.strip()
        
        # Busca el patrón "5.11.00.00 Texto..."
        # ^\s* -> Comienzo de línea, opcionalmente espacios
        # ([\d\.]+) -> Captura el grupo de números y puntos (EL NÚMERO)
        # \s+       -> Uno o más espacios (el separador)
        # (.*)      -> Captura todo lo demás (EL TEXTO)
        match = re.search(r'^\s*([\d\.]+)\s+(.*)', title)
        
        if match:
            # Devuelve una tupla: (numero, texto_normalizado)
            # ej: ('5.11.00.00', 'total inversiones financieras')
            return (match.group(1).strip(), match.group(2).strip().lower())
        
        # Si no hay match, NO es un título válido. Devolver None.
        return None

    def _find_start_row(self, df, col_range_str):
        """Encuentra el índice de la primera fila que contiene un título normalizable."""
        col_indices = self._parse_col_range(col_range_str)
        
        num_cols = len(df.columns) 

        for index, row in df.iterrows():
            for col_idx in col_indices:
                if col_idx >= num_cols:
                    continue 

                cell_value = row.iloc[col_idx]
                # self._normalize_title ahora devuelve (num, text) o None.
                # Si no es None, es un título válido y encontramos la fila.
                if self._normalize_title(cell_value) is not None:
                    return index  # Devuelve el índice de la fila
        return None # No se encontró ninguna fila con títulos

    def _process_dataframe(self, df, start_row, config):
        """
        Extrae los datos relevantes (título, actual, anterior) del DataFrame
        y los devuelve en una LISTA de diccionarios.
        """
        data_list = []
        if start_row is None:
            return data_list

        try:
            title_col_indices = self._parse_col_range(config['title_range'])
            actual_col_idx = self._col_to_int(config['actual_col'])
            anterior_col_idx = self._col_to_int(config['anterior_col'])
        except Exception as e:
            raise ValueError(f"Error en configuración de columnas: {e}")

        num_cols = len(df.columns) 
        if actual_col_idx >= num_cols or anterior_col_idx >= num_cols:
             raise IndexError(f"CONFIGURACIÓN INVÁLIDA: Las columnas de período (Actual: {config['actual_col']}, Anterior: {config['anterior_col']}) están fuera de los límites. El archivo solo tiene {num_cols} columnas.")
        
        max_title_col = max(title_col_indices)
        if max_title_col >= num_cols:
            print(f"Advertencia: El Rango de Títulos ('{config['title_range']}') incluye una columna (índice {max_title_col}) que está fuera de los límites. El archivo solo tiene {num_cols} columnas.")

        df_subset = df.iloc[start_row:]
        
        for _, row in df_subset.iterrows():
            title_parts = None
            original_full_title = None
            
            # 1. Encontrar el título en la fila
            for col_idx in title_col_indices:
                if col_idx >= num_cols:
                    continue
                
                cell_value = row.iloc[col_idx]
                title_parts = self._normalize_title(cell_value)
                if title_parts:
                    original_full_title = cell_value.strip() # Guardamos el título original
                    break 
            
            # 2. Si se encontró título, extraer valores
            if title_parts and original_full_title:
                number, text = title_parts
                
                actual_val = pd.to_numeric(row.iloc[actual_col_idx], errors='coerce')
                anterior_val = pd.to_numeric(row.iloc[anterior_col_idx], errors='coerce')
                
                data_list.append({
                    'num': number,
                    'text': text,
                    'full_title': original_full_title,
                    'actual': actual_val,
                    'anterior': anterior_val,
                    'matched': False  # Flag para rastrear coincidencias
                })
                
        return data_list

    def _process_dataframe(self, df, start_row, config):
        """
        Extrae los datos relevantes (título, actual, anterior) del DataFrame
        y los devuelve en un diccionario: {'titulo_normalizado': (actual, anterior)}
        """
        data_map = {}
        if start_row is None:
            return data_map # DataFrame vacío si no se encontró fila de inicio

        try:
            title_col_indices = self._parse_col_range(config['title_range'])
            actual_col_idx = self._col_to_int(config['actual_col'])
            anterior_col_idx = self._col_to_int(config['anterior_col'])
        except Exception as e:
            raise ValueError(f"Error en configuración de columnas: {e}")

        # Itera solo desde la fila de inicio
        df_subset = df.iloc[start_row:]
        
        for _, row in df_subset.iterrows():
            normalized_title = None
            # 1. Encontrar el título en la fila
            for col_idx in title_col_indices:
                if col_idx < len(row):
                    cell_value = row.iloc[col_idx]
                    title = self._normalize_title(cell_value)
                    if title:
                        normalized_title = title
                        break # Encontrado el título para esta fila
            
            # 2. Si se encontró título, extraer valores
            if normalized_title:
                # Usar pd.to_numeric para convertir de forma segura, errores se vuelven NaN
                actual_val = pd.to_numeric(row.iloc[actual_col_idx], errors='coerce')
                anterior_val = pd.to_numeric(row.iloc[anterior_col_idx], errors='coerce')
                
                if normalized_title in data_map:
                    # Advertencia: Título duplicado. Podríamos sumar, pero por ahora sobreescribimos.
                    pass 
                
                data_map[normalized_title] = (actual_val, anterior_val)
                
        return data_map

    def _check_values(self, context, title, cliente_val, salida_val):
        """
        Lógica de comparación específica con reglas de 0 y multiplicador 1000,
        usando comparación exacta de enteros.
        """
        
        # 1. Tratar NaN (vacío) como 0 ANTES de cualquier cálculo
        #    Esto es crucial por los 'nan' que se ven en tu data_salida
        cliente_val = cliente_val if pd.notna(cliente_val) else 0
        salida_val = salida_val if pd.notna(salida_val) else 0 

        # 2. Escalar valor de Salida y REDONDEAR al entero más cercano
        #    Ej: 178054291000 / 1000.0 = 178054291.0
        #    round(178054291.0) = 178054291
        salida_scaled = round(salida_val / 1000.0)
        
        # 3. Redondear valor de Cliente (por si acaso es un float 123.0)
        #    round(178054299) = 178054299
        cliente_val = round(cliente_val)

        # 4. Aplicar reglas de negocio
        
        # Regla: Si Cliente es 0, Salida (escalada) debe ser 0.
        if cliente_val == 0:
            if salida_scaled != 0:
                self.inconsistencias.append(
                    f"[{context}] '{title}': Cliente es 0, pero Salida reporta {salida_val} (escalado: {salida_scaled})."
                )
            # Si ambos son 0, está OK.
            return 

        # Regla: Si Cliente NO es 0, deben coincidir EXACTAMENTE.
        # ¡Ya no usamos np.isclose!
        if cliente_val != salida_scaled:
            self.inconsistencias.append(
                f"[{context}] '{title}': DISCREPANCIA. Cliente: {cliente_val}, Salida: {salida_val} (escalado: {salida_scaled})."
            )

    def compare(self):
        """
        Método principal que ejecuta todo el proceso de comparación.
        """
        self.inconsistencias = [] # Limpiar corridas anteriores

        for cliente_sheet_idx, salida_sheet_idx in self.sheet_map:
            sheet_context = f"Hoja Cliente {cliente_sheet_idx + 1} vs Hoja Salida {salida_sheet_idx + 1}"
            
            try:
                # Cargar hojas (sin cabecera para controlar nosotros mismos)
                df_cliente = pd.read_excel(self.cliente_path, sheet_name=cliente_sheet_idx, header=None)
                df_salida = pd.read_excel(self.salida_path, sheet_name=salida_sheet_idx, header=None)
            except Exception as e:
                self.inconsistencias.append(f"ERROR: No se pudo leer las hojas {sheet_context}. Detalle: {e}")
                continue # Saltar a la siguiente hoja

            # 1. Encontrar filas de inicio
            print("Procesando hoja:", sheet_context)
            start_row_cliente = self._find_start_row(df_cliente, self.cliente_config['title_range'])
            print("start_row_cliente:", start_row_cliente)
            start_row_salida = self._find_start_row(df_salida, self.salida_config['title_range'])
            print("start_row_salida:", start_row_salida)
            
            if start_row_cliente is None:
                self.inconsistencias.append(f"ADVERTENCIA: No se encontraron títulos en {sheet_context} (Cliente).")
                continue
            if start_row_salida is None:
                self.inconsistencias.append(f"ADVERTENCIA: No se encontraron títulos en {sheet_context} (Salida).")
                continue

            # 2. Procesar DFs en diccionarios {titulo: (val1, val2)}
            data_cliente = self._process_dataframe(df_cliente, start_row_cliente, self.cliente_config)
            print("data_cliente:", data_cliente)
            data_salida = self._process_dataframe(df_salida, start_row_salida, self.salida_config)
            print("data_salida:", data_salida)

            # 3. Comparar los diccionarios
            if not data_cliente:
                 self.inconsistencias.append(f"ADVERTENCIA: No se extrajeron datos de Cliente en {sheet_context}.")
                 continue

            for title, (cliente_actual, cliente_anterior) in data_cliente.items():
                
                if title in data_salida:
                    salida_actual, salida_anterior = data_salida[title]
                    
                    # Comparar Periodo Actual vs Col I
                    self._check_values(f"{sheet_context} [Actual]", title, cliente_actual, salida_actual)
                    
                    # Comparar Periodo Anterior vs Col J
                    self._check_values(f"{sheet_context} [Anterior]", title, cliente_anterior, salida_anterior)

                else:
                    # Título está en Cliente pero no en Salida.
                    # Verificar si es válido bajo la regla del 0.
                    cliente_actual = cliente_actual if pd.notna(cliente_actual) else 0
                    cliente_anterior = cliente_anterior if pd.notna(cliente_anterior) else 0
                    
                    if cliente_actual != 0 or cliente_anterior != 0:
                        self.inconsistencias.append(
                            f"[{sheet_context}] '{title}': Título existe en Cliente (Valores: {cliente_actual}, {cliente_anterior}) pero FALTA en Salida."
                        )
        
        if not self.inconsistencias:
            return "--- PROCESO COMPLETADO --- \n\n¡No se encontraron inconsistencias!"
            
        return "--- PROCESO COMPLETADO --- \n\nSe encontraron las siguientes inconsistencias:\n\n" + "\n".join(self.inconsistencias)


class ComparisonApp:
    """
    Construye y maneja la interfaz gráfica (GUI) con Tkinter.
    """
    def __init__(self, root):
        self.root = root
        self.root.title("Comparador de Archivos Excel")
        
        # Escala de DPI
        self.setup_dpi()
        
        # Configuración de estilo
        style = ttk.Style()
        style.theme_use('clam')
        style.configure('TButton', padding=6, relief="flat", background="#0078D4", foreground="white")
        style.map('TButton', background=[('active', '#005A9E')])
        style.configure('TLabel', padding=5)
        style.configure('TEntry', padding=5)
        style.configure('Main.TFrame', background='#F0F0F0')
        style.configure('Header.TLabel', font=('Segoe UI', 12, 'bold'))

        # Frame principal
        main_frame = ttk.Frame(root, padding="10 10 10 10", style='Main.TFrame')
        main_frame.grid(row=0, column=0, sticky=(tk.W, tk.E, tk.N, tk.S))
        root.columnconfigure(0, weight=1)
        root.rowconfigure(0, weight=1)

        # --- Sección de Archivos ---
        file_frame = ttk.Labelframe(main_frame, text="1. Seleccionar Archivos", padding="10")
        file_frame.grid(row=0, column=0, columnspan=2, sticky=(tk.W, tk.E), pady=5)
        file_frame.columnconfigure(1, weight=1)

        # Archivo Cliente
        ttk.Label(file_frame, text="Archivo Cliente:").grid(row=0, column=0, sticky=tk.W)
        self.cliente_path_entry = ttk.Entry(file_frame, width=60)
        self.cliente_path_entry.grid(row=0, column=1, sticky=(tk.W, tk.E), padx=5)
        ttk.Button(file_frame, text="Buscar...", command=lambda: self.browse_file(self.cliente_path_entry)).grid(row=0, column=2)

        # Archivo Salida
        ttk.Label(file_frame, text="Archivo Salida:").grid(row=1, column=0, sticky=tk.W)
        self.salida_path_entry = ttk.Entry(file_frame, width=60)
        self.salida_path_entry.grid(row=1, column=1, sticky=(tk.W, tk.E), padx=5)
        ttk.Button(file_frame, text="Buscar...", command=lambda: self.browse_file(self.salida_path_entry)).grid(row=1, column=2)

        # --- Sección de Configuración Cliente ---
        config_frame = ttk.Labelframe(main_frame, text="2. Configuración Cliente", padding="10")
        config_frame.grid(row=1, column=0, columnspan=2, sticky=(tk.W, tk.E), pady=5)
        config_frame.columnconfigure(1, weight=1)

        ttk.Label(config_frame, text="Rango Títulos (ej: A:F):").grid(row=0, column=0, sticky=tk.W, padx=5, pady=2)
        self.title_range_entry = ttk.Entry(config_frame)
        self.title_range_entry.grid(row=0, column=1, sticky=(tk.W, tk.E), padx=5, pady=2)

        ttk.Label(config_frame, text="Col. Periodo Actual (ej: G):").grid(row=1, column=0, sticky=tk.W, padx=5, pady=2)
        self.actual_col_entry = ttk.Entry(config_frame)
        self.actual_col_entry.grid(row=1, column=1, sticky=(tk.W, tk.E), padx=5, pady=2)

        ttk.Label(config_frame, text="Col. Periodo Anterior (ej: H):").grid(row=2, column=0, sticky=tk.W, padx=5, pady=2)
        self.anterior_col_entry = ttk.Entry(config_frame)
        self.anterior_col_entry.grid(row=2, column=1, sticky=(tk.W, tk.E), padx=5, pady=2)

        # --- Sección de Acción ---
        self.run_button = ttk.Button(main_frame, text="🚀 Ejecutar Comparación", command=self.run_comparison)
        self.run_button.grid(row=2, column=0, columnspan=2, sticky=(tk.W, tk.E), pady=10, ipady=5)

        # --- Sección de Resultados ---
        results_frame = ttk.Labelframe(main_frame, text="3. Resultados", padding="10")
        results_frame.grid(row=3, column=0, columnspan=2, sticky=(tk.W, tk.E, tk.N, tk.S), pady=5)
        main_frame.rowconfigure(3, weight=1) # Permite que el frame de resultados se expanda

        self.results_text = scrolledtext.ScrolledText(results_frame, wrap=tk.WORD, width=80, height=20, font=("Consolas", 9))
        self.results_text.grid(row=0, column=0, sticky=(tk.W, tk.E, tk.N, tk.S))
        results_frame.columnconfigure(0, weight=1)
        results_frame.rowconfigure(0, weight=1)

    def setup_dpi(self):
        """Configura el escalado de DPI para Windows."""
        if os.name == 'nt':
            try:
                ctypes.windll.shcore.SetProcessDpiAwareness(1)
            except Exception as e:
                print(f"Advertencia: No se pudo establecer el DPI awareness. {e}")

    def browse_file(self, entry_widget):
        """Abre un diálogo para seleccionar un archivo Excel."""
        file_path = filedialog.askopenfilename(
            title="Seleccionar archivo Excel",
            filetypes=[("Archivos Excel", "*.xlsx *.xls")]
        )
        if file_path:
            entry_widget.delete(0, tk.END)
            entry_widget.insert(0, file_path)

    def run_comparison(self):
        """
        Valida las entradas y ejecuta el proceso de comparación.
        """
        # 1. Obtener y validar entradas
        cliente_path = self.cliente_path_entry.get()
        salida_path = self.salida_path_entry.get()
        title_range = self.title_range_entry.get()
        actual_col = self.actual_col_entry.get()
        anterior_col = self.anterior_col_entry.get()

        if not all([cliente_path, salida_path, title_range, actual_col, anterior_col]):
            messagebox.showerror("Error de Validación", "Todos los campos son obligatorios.")
            return

        if not os.path.exists(cliente_path):
            messagebox.showerror("Error de Archivo", f"No se encuentra el archivo Cliente:\n{cliente_path}")
            return
        
        if not os.path.exists(salida_path):
            messagebox.showerror("Error de Archivo", f"No se encuentra el archivo Salida:\n{salida_path}")
            return
            
        cliente_config = {
            'title_range': title_range,
            'actual_col': actual_col,
            'anterior_col': anterior_col
        }

        # 2. Deshabilitar botón y mostrar "Procesando"
        self.run_button.config(text="Procesando... ⏳", state=tk.DISABLED)
        self.results_text.delete(1.0, tk.END)
        self.results_text.insert(tk.END, "Iniciando comparación, esto puede tardar unos segundos...")
        self.root.update_idletasks() # Forzar actualización de la GUI

        # 3. Ejecutar la lógica de negocio
        try:
            comparator = ExcelComparator(cliente_path, salida_path, cliente_config)
            results = comparator.compare()
            
            # 4. Mostrar resultados
            self.results_text.delete(1.0, tk.END)
            self.results_text.insert(tk.END, results)
            
        except Exception as e:
            # 5. Manejar errores
            self.results_text.delete(1.0, tk.END)
            self.results_text.insert(tk.END, f"--- ERROR INESPERADO ---\n\n{type(e).__name__}: {e}")
            messagebox.showerror("Error en Procesamiento", f"Ocurrió un error: {e}")
        
        finally:
            # 6. Reactivar el botón
            self.run_button.config(text="🚀 Ejecutar Comparación", state=tk.NORMAL)


if __name__ == "__main__":
    root = tk.Tk()
    app = ComparisonApp(root)
    root.mainloop()