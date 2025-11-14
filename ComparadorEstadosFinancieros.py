import tkinter as tk
from tkinter import ttk, filedialog, messagebox, scrolledtext
import pandas as pd
import numpy as np
import re
import os
import ctypes
from dateutil.parser import parse as date_parse
from dateutil.parser._parser import ParserError

class ExcelComparator:
    """
    Contiene toda la lógica de negocio. 
    Auto-detecta las columnas EN CADA HOJA.
    """
    
    # --- MÉTODO __init__ ACTUALIZADO ---
    # Ya no necesita self.cliente_config ni self.salida_config
    def __init__(self, cliente_path, salida_path, year, month):
        self.cliente_path = cliente_path
        self.salida_path = salida_path
        self.year = year
        self.month = month
        self.inconsistencias = []

        # Mapeo de hojas (basado en índice 0-based de pandas)
        self.sheet_map = [(2, 1), (3, 2), (4, 3)]

    # --- (Las siguientes funciones auxiliares NO CAMBIAN) ---

    def _col_to_int(self, col_str):
        num = 0
        for c in col_str:
            num = num * 26 + (ord(c.upper()) - ord('A')) + 1
        return num - 1

    def _int_to_col(self, n):
        string = ""
        n_copy = n
        while n_copy >= 0:
            string = chr(n_copy % 26 + ord('A')) + string
            n_copy = n_copy // 26 - 1
        return string

    def _parse_col_range(self, range_str):
        try:
            start_col, end_col = range_str.split(':')
            start_idx = self._col_to_int(start_col)
            end_idx = self._col_to_int(end_col)
            return list(range(start_idx, end_idx + 1))
        except Exception:
            raise ValueError(f"Formato de rango de título inválido: '{range_str}'. Use 'A:F'.")

    def _normalize_title(self, title):
        if not isinstance(title, str):
            return None
        title = title.strip()
        match = re.search(r'^\s*([\d\.]+)\s+(.*)', title)
        if match:
            return (match.group(1).strip(), match.group(2).strip().lower())
        return None

    def _find_start_row(self, df, col_range_str):
        col_indices = self._parse_col_range(col_range_str)
        num_cols = len(df.columns) 
        for index, row in df.iterrows():
            for col_idx in col_indices:
                if col_idx >= num_cols:
                    continue 
                cell_value = row.iloc[col_idx]
                if self._normalize_title(cell_value) is not None:
                    return index
        return None

    def _detect_columns(self, file_path, sheet_idx):
        """
        Auto-detecta el rango de títulos y las columnas de período 
        para un archivo y hoja específicos (usa las primeras 50 filas).
        Ahora maneja rangos de fechas (ej: 'YYYY-MM-DD - YYYY-MM-DD').
        """
        try:
            df = pd.read_excel(file_path, sheet_name=sheet_idx, header=None, nrows=120)
        except Exception as e:
            raise ValueError(f"No se pudo leer la hoja {sheet_idx+1} de {os.path.basename(file_path)}. Detalle: {e}")

        # --- 1. Detección de Columna de Período (Columnas D hasta L) ---
        periodo_search_cols = list(range(3, 12)) # D(3) a L(11)
        actual_col = None
        anterior_col = None

        for col_idx in periodo_search_cols:
            if col_idx >= len(df.columns):
                continue
            for row_idx in range(df.shape[0]):
                cell_value = df.iloc[row_idx, col_idx]
                try:
                    # --- ¡AQUÍ ESTÁ LA MODIFICACIÓN! ---
                    
                    # 1. Convertir a string y dividir por " - " (espacio-guion-espacio)
                    #    Tomamos el último elemento (-1) por si es un rango.
                    #    .strip() elimina espacios en blanco al inicio/final.
                    date_string_to_check = str(cell_value).split(' - ')[-1].strip()
                    
                    # 2. Analizar solo la cadena resultante (la fecha de fin)
                    date = date_parse(date_string_to_check, dayfirst=True)
                    
                    # --- FIN DE LA MODIFICACIÓN ---
                    
                    if date.year == self.year and date.month == self.month:
                        actual_col = self._int_to_col(col_idx)
                        anterior_col = self._int_to_col(col_idx + 1)
                        break 
                except (ParserError, TypeError, OverflowError, ValueError):
                    # Ignora celdas que no son fechas
                    continue
            if actual_col:
                break

        if not actual_col:
            raise ValueError(f"AUTO-DETECCIÓN FALLIDA (Período):\nNo se pudo encontrar una columna de fecha con {self.month:02d}/{self.year} en '{os.path.basename(file_path)}' (Hoja {sheet_idx+1}, Cols D-L).")

        # --- 2. Detección de Rango de Títulos (Columnas A hasta F) ---
        # (Esta parte no cambia)
        title_search_cols = list(range(0, 6)) # A(0) a F(5)
        found_title_cols = [] 

        for col_idx in title_search_cols:
            if col_idx >= len(df.columns):
                continue
            for row_idx in range(df.shape[0]):
                cell_value = df.iloc[row_idx, col_idx]
                if self._normalize_title(cell_value) is not None:
                    found_title_cols.append(col_idx)
                    break 

        if not found_title_cols:
            raise ValueError(f"AUTO-DETECCIÓN FALLIDA (Títulos):\nNo se pudo encontrar ninguna columna con Títulos en '{os.path.basename(file_path)}' (Hoja {sheet_idx+1}, Cols A-F).")

        min_col = min(found_title_cols)
        max_col = max(found_title_cols)
        title_range = f"{self._int_to_col(min_col)}:{self._int_to_col(max_col)}"

        # --- 3. Retornar Configuración ---
        return {
            'title_range': title_range,
            'actual_col': actual_col,
            'anterior_col': anterior_col
        }

    def _process_dataframe(self, df, start_row, config):
        """
        Extrae los datos relevantes (título, actual, anterior) del DataFrame
        y los devuelve en una LISTA de diccionarios.
        
        NUEVA REGLA: Deja de procesar si encuentra 4 filas consecutivas 
        sin un título válido.
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
        

        # --- ¡CAMBIO! Se reemplaza '_' por 'index' para capturar la fila ---
        for index, row in df_subset.iterrows():
            title_parts = None
            original_full_title = None
            
            # 1. Encontrar el título en la fila
            for col_idx in title_col_indices:
                if col_idx >= num_cols:
                    continue
                cell_value = row.iloc[col_idx]
                title_parts = self._normalize_title(cell_value)
                if title_parts:
                    original_full_title = cell_value.strip()
                    break 
            

            # 3. Si se encontró título (y no hemos parado), extraer valores
            if title_parts and original_full_title:
                number, text = title_parts
                actual_val = pd.to_numeric(row.iloc[actual_col_idx], errors='coerce')
                anterior_val = pd.to_numeric(row.iloc[anterior_col_idx], errors='coerce')
                
                data_list.append({
                    'num': number, 'text': text, 'full_title': original_full_title,
                    'actual': actual_val, 'anterior': anterior_val, 'matched': False,
                    'row_num': index + 1  # <-- ¡AQUÍ SE GUARDA EL NÚMERO DE FILA (1-based)!
                })
                
        return data_list

    def _check_values(self, context, title, cliente_val, salida_val):
        """
        Lógica de comparación específica con reglas de 0 y multiplicador 1000,
        usando el VALOR ABSOLUTO de enteros (INT).
        """
        
        # 1. Tratar NaN (vacío) como 0 ANTES de cualquier cálculo
        cliente_val = cliente_val if pd.notna(cliente_val) else 0
        salida_val = salida_val if pd.notna(salida_val) else 0 

        # 2. Escalar valor de Salida y REDONDEAR al entero más cercano
        #    Luego, forzar la conversión a int.
        salida_scaled = int(round(salida_val / 1000.0))
        
        # 3. Redondear valor de Cliente y forzar la conversión a int.
        cliente_val = int(round(cliente_val))

        # --- ¡AQUÍ ESTÁ LA MODIFICACIÓN! ---
        # 4. Obtener el valor absoluto de ambos números
        abs_cliente_val = abs(cliente_val)
        abs_salida_scaled = abs(salida_scaled)
        # --- FIN DE LA MODIFICACIÓN ---

        # 5. Aplicar reglas de negocio (usando los valores absolutos)
        
        # Regla: Si Cliente (abs) es 0, Salida (abs) debe ser 0.
        if abs_cliente_val == 0:
            if abs_salida_scaled != 0:
                self.inconsistencias.append(
                    f"[{context}] '{title}': Cliente es 0, pero Salida reporta valor (Escalado: {abs_salida_scaled})."
                )
            return 

        # Regla: Si Cliente (abs) NO es 0, deben coincidir EXACTAMENTE.
        if abs_cliente_val != abs_salida_scaled:
            self.inconsistencias.append(
                f"[{context}] '{title}': DISCREPANCIA . Cliente: {abs_cliente_val}, Salida (Escalado): {abs_salida_scaled}."
            )

    # --- MÉTODO compare ACTUALIZADO ---
    def compare(self):
        """
        Método principal. La auto-detección ahora ocurre DENTRO del bucle.
        """
        self.inconsistencias = []

        try:
            cliente_sheet_names = pd.ExcelFile(self.cliente_path).sheet_names
            salida_sheet_names = pd.ExcelFile(self.salida_path).sheet_names
        except Exception as e:
            self.inconsistencias.append(f"ERROR CRÍTICO: No se pudo abrir uno de los archivos Excel. Detalle: {e}")
            return "\n".join(self.inconsistencias)

        for cliente_sheet_idx, salida_sheet_idx in self.sheet_map:
            
            cliente_config = None
            salida_config = None
            sheet_context = f"(Hoja Cliente idx {cliente_sheet_idx+1} vs Hoja Salida idx {salida_sheet_idx+1})" 
            
            try:
                cliente_sheet_name = cliente_sheet_names[cliente_sheet_idx]
                salida_sheet_name = salida_sheet_names[salida_sheet_idx]
                sheet_context = f"Hoja '{cliente_sheet_name}' vs Hoja '{salida_sheet_name}'"

                self.inconsistencias.append(f"\n")
                cliente_config = self._detect_columns(self.cliente_path, cliente_sheet_idx)
                self.inconsistencias.append(f"  -> Cliente OK: Títulos en '{cliente_config['title_range']}', Actual en '{cliente_config['actual_col']}'")

                salida_config = self._detect_columns(self.salida_path, salida_sheet_idx)
                self.inconsistencias.append(f"  -> Salida OK: Títulos en '{salida_config['title_range']}', Actual en '{salida_config['actual_col']}'")
                
                df_cliente = pd.read_excel(self.cliente_path, sheet_name=cliente_sheet_idx, header=None)
                df_salida = pd.read_excel(self.salida_path, sheet_name=salida_sheet_idx, header=None)
            
            except Exception as e:
                self.inconsistencias.append(f"ERROR: No se pudo procesar/detectar {sheet_context}. Detalle: {e}")
                continue 

            start_row_cliente = self._find_start_row(df_cliente, cliente_config['title_range'])
            start_row_salida = self._find_start_row(df_salida, salida_config['title_range'])
            
            if start_row_cliente is None:
                self.inconsistencias.append(f"ADVERTENCIA: No se encontraron títulos en {sheet_context} (Cliente).")
                continue
            if start_row_salida is None:
                self.inconsistencias.append(f"ADVERTENCIA: No se encontraron títulos en {sheet_context} (Salida).")
                continue

            data_cliente = self._process_dataframe(df_cliente, start_row_cliente, cliente_config)
            data_salida = self._process_dataframe(df_salida, start_row_salida, salida_config)

            if not data_cliente:
                 self.inconsistencias.append(f"ADVERTENCIA: No se extrajeron datos de Cliente en {sheet_context}.")
                 continue
            if not data_salida:
                 self.inconsistencias.append(f"ADVERTENCIA: No se extrajeron datos de Salida en {sheet_context}.")
                 continue
                 
            # 6.A. Lógica de Coincidencia por NÚMERO
            for cliente_item in data_cliente:
                found_match = None
                for salida_item in data_salida:
                    if not salida_item['matched'] and cliente_item['num'] == salida_item['num']:
                        found_match = salida_item
                        break 
                if found_match:
                    cliente_item['matched'] = True
                    found_match['matched'] = True
                    self._check_values(f"{sheet_context} [Actual]", cliente_item['full_title'], cliente_item['actual'], found_match['actual'])
                    self._check_values(f"{sheet_context} [Anterior]", cliente_item['full_title'], cliente_item['anterior'], found_match['anterior'])

            # 6.B. Lógica de Coincidencia por TEXTO
            for cliente_item in data_cliente:
                if cliente_item['matched']: continue
                found_match = None
                for salida_item in data_salida:
                    if not salida_item['matched'] and cliente_item['text'] == salida_item['text']:
                        found_match = salida_item
                        break
                if found_match:
                    cliente_item['matched'] = True
                    found_match['matched'] = True
                    self.inconsistencias.append(
                        f"[{sheet_context}] [AVISO] Coincidencia por TEXTO: Cliente ('{cliente_item['full_title']}') vs Salida ('{found_match['full_title']}')"
                    )
                    self._check_values(f"{sheet_context} [Actual]", cliente_item['full_title'], cliente_item['actual'], found_match['actual'])
                    self._check_values(f"{sheet_context} [Anterior]", cliente_item['full_title'], cliente_item['anterior'], found_match['anterior'])

            # 6.C. Reportar FALTANTES en Salida
            for cliente_item in data_cliente:
                if not cliente_item['matched']:
                    cliente_actual = cliente_item['actual'] if pd.notna(cliente_item['actual']) else 0
                    cliente_anterior = cliente_item['anterior'] if pd.notna(cliente_item['anterior']) else 0
                    if cliente_actual != 0 or cliente_anterior != 0:
                        # --- ¡CAMBIO! Se añade la fila del cliente ---
                        self.inconsistencias.append(
                            f"[{sheet_context}] (Fila Cliente: {cliente_item['row_num']}) '{cliente_item['full_title']}': Título FALTA en Salida (Valores Cliente: {cliente_actual}, {cliente_anterior})."
                        )
            
            # 6.D. Reportar SOBRANTES en Salida
            for salida_item in data_salida:
                if not salida_item['matched']:
                    salida_actual = salida_item['actual'] if pd.notna(salida_item['actual']) else 0
                    salida_anterior = salida_item['anterior'] if pd.notna(salida_item['anterior']) else 0
                    if salida_actual != 0 or salida_anterior != 0:
                        salida_scaled_actual = abs(int(round(salida_actual / 1000.0)))
                        salida_scaled_anterior = abs(int(round(salida_anterior / 1000.0)))
                        if salida_scaled_actual != 0 or salida_scaled_anterior != 0:
                            # --- ¡CAMBIO! Se añade la fila de salida ---
                            self.inconsistencias.append(
                                f"[{sheet_context}] (Fila Salida: {salida_item['row_num']}) '{salida_item['full_title']}': Título SOBRA en Salida, no existe en Cliente (Valores Salida escalados: {salida_scaled_actual}, {salida_scaled_anterior})."
                            )

        if not self.inconsistencias:
            return "--- PROCESO COMPLETADO --- \n\n¡No se encontraron inconsistencias!"
            
        final_messages = [msg for msg in self.inconsistencias 
                          if not msg.startswith("  ->") and not msg.endswith("Detectando...")]
        
        if not final_messages:
            return "--- PROCESO COMPLETADO --- \n\n¡No se encontraron inconsistencias!"

        return "--- PROCESO COMPLETADO --- \n\nSe encontraron las siguientes inconsistencias:\n\n" + "\n".join(final_messages)


# --- CLASE ComparisonApp ACTUALIZADA ---
class ComparisonApp:
    """
    Construye y maneja la interfaz gráfica (GUI) con Tkinter.
    """
    def __init__(self, root):
        self.root = root
        self.root.title("Comparador de Archivos Excel (Auto-Detección)")
        self.setup_dpi()
        
        style = ttk.Style()
        style.theme_use('clam')
        style.configure('TButton', padding=6, relief="flat", background="#0078D4", foreground="white")
        style.map('TButton', background=[('active', '#005A9E')])
        style.configure('TLabel', padding=5)
        style.configure('TEntry', padding=5)
        style.configure('Main.TFrame', background='#F0F0F0')
        style.configure('Header.TLabel', font=('Segoe UI', 12, 'bold'))

        main_frame = ttk.Frame(root, padding="10 10 10 10", style='Main.TFrame')
        main_frame.grid(row=0, column=0, sticky=(tk.W, tk.E, tk.N, tk.S))
        root.columnconfigure(0, weight=1)
        root.rowconfigure(0, weight=1)

        # --- Sección de Archivos (Sin cambios) ---
        file_frame = ttk.Labelframe(main_frame, text="1. Seleccionar Archivos", padding="10")
        file_frame.grid(row=0, column=0, columnspan=2, sticky=(tk.W, tk.E), pady=5)
        file_frame.columnconfigure(1, weight=1)

        ttk.Label(file_frame, text="Archivo Cliente:").grid(row=0, column=0, sticky=tk.W)
        self.cliente_path_entry = ttk.Entry(file_frame, width=60)
        self.cliente_path_entry.grid(row=0, column=1, sticky=(tk.W, tk.E), padx=5)
        ttk.Button(file_frame, text="Buscar...", command=lambda: self.browse_file(self.cliente_path_entry)).grid(row=0, column=2)

        ttk.Label(file_frame, text="Archivo Salida:").grid(row=1, column=0, sticky=tk.W)
        self.salida_path_entry = ttk.Entry(file_frame, width=60)
        self.salida_path_entry.grid(row=1, column=1, sticky=(tk.W, tk.E), padx=5)
        ttk.Button(file_frame, text="Buscar...", command=lambda: self.browse_file(self.salida_path_entry)).grid(row=1, column=2)

        # --- Sección de Configuración (ACTUALIZADA) ---
        config_frame = ttk.Labelframe(main_frame, text="2. Configuración del Período Actual", padding="10")
        config_frame.grid(row=1, column=0, columnspan=2, sticky=(tk.W, tk.E), pady=5)
        config_frame.columnconfigure(1, weight=1)

        ttk.Label(config_frame, text="Año (YYYY):").grid(row=0, column=0, sticky=tk.W, padx=5, pady=2)
        self.year_entry = ttk.Entry(config_frame, width=10)
        self.year_entry.grid(row=0, column=1, sticky=tk.W, padx=5, pady=2)

        ttk.Label(config_frame, text="Mes (1-12):").grid(row=1, column=0, sticky=tk.W, padx=5, pady=2)
        self.month_var = tk.StringVar(value="1") # Valor inicial
        self.month_entry = ttk.Spinbox(config_frame, from_=1, to=12, width=8, textvariable=self.month_var, format="%02.0f")
        self.month_entry.grid(row=1, column=1, sticky=tk.W, padx=5, pady=2)


        # --- Sección de Acción (Sin cambios) ---
        self.run_button = ttk.Button(main_frame, text="🚀 Ejecutar Comparación", command=self.run_comparison)
        self.run_button.grid(row=2, column=0, columnspan=2, sticky=(tk.W, tk.E), pady=10, ipady=5)

        # --- Sección de Resultados (Sin cambios) ---
        results_frame = ttk.Labelframe(main_frame, text="3. Resultados", padding="10")
        results_frame.grid(row=3, column=0, columnspan=2, sticky=(tk.W, tk.E, tk.N, tk.S), pady=5)
        main_frame.rowconfigure(3, weight=1) 

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

    # --- MÉTODO run_comparison ACTUALIZADO ---
    def run_comparison(self):
        """
        Valida las nuevas entradas y ejecuta el proceso de comparación.
        """
        # 1. Obtener y validar entradas
        cliente_path = self.cliente_path_entry.get()
        salida_path = self.salida_path_entry.get()
        
        try:
            year = int(self.year_entry.get())
            month = int(self.month_entry.get())
            if not (1 <= month <= 12):
                raise ValueError("Mes fuera de rango 1-12")
            if not (1900 < year < 2100):
                 raise ValueError("Año fuera de rango 1900-2100")
        except ValueError:
            messagebox.showerror("Error de Validación", "Por favor, ingrese un Año (ej: 2024) y un Mes (1-12) válidos.")
            return

        if not os.path.exists(cliente_path):
            messagebox.showerror("Error de Archivo", f"No se encuentra el archivo Cliente:\n{cliente_path}")
            return
        if not os.path.exists(salida_path):
            messagebox.showerror("Error de Archivo", f"No se encuentra el archivo Salida:\n{salida_path}")
            return

        # 2. Deshabilitar botón y mostrar "Procesando"
        self.run_button.config(text="Procesando... ⏳", state=tk.DISABLED)
        self.results_text.delete(1.0, tk.END)
        self.results_text.insert(tk.END, "Iniciando comparación...\n")
        self.root.update_idletasks() 

        # 3. Ejecutar la lógica de negocio (constructor actualizado)
        try:
            comparator = ExcelComparator(cliente_path, salida_path, year, month)
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