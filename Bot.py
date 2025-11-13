"""
BOT DE AUTOMATIZACIÓN - REGISTRO DE DENUNCIAS SUNAT
Con Interfaz Gráfica (GUI)
Autor: Sistema Automatizado
Fecha: 2025
"""

import pandas as pd
from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait, Select
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.chrome.options import Options
from selenium.common.exceptions import TimeoutException, NoSuchElementException
from datetime import datetime, timedelta
import time
import logging
import os
import threading
import tkinter as tk
from tkinter import ttk, filedialog, messagebox, scrolledtext


# ============================================
# INTERFAZ GRÁFICA
# ============================================
class InterfazBot:
    def __init__(self):
        self.ventana = tk.Tk()
        self.ventana.title("Bot de Registro de Denuncias SUNAT")
        self.ventana.geometry("800x600")
        self.ventana.resizable(False, False)
        
        # Variables
        self.ruta_archivo = tk.StringVar()
        self.usuario = tk.StringVar()
        self.password = tk.StringVar()
        self.bot = None
        self.proceso_activo = False
        self.hilo_proceso = None
        
        self.crear_interfaz()
        
    def crear_interfaz(self):
        """Crea todos los elementos de la interfaz"""
        
        # ============================================
        # TÍTULO
        # ============================================
        frame_titulo = tk.Frame(self.ventana, bg="#0063AE", height=80)
        frame_titulo.pack(fill=tk.X)
        
        label_titulo = tk.Label(
            frame_titulo,
            text="🤖 BOT DE REGISTRO DE DENUNCIAS SUNAT",
            font=("Arial", 16, "bold"),
            bg="#0063AE",
            fg="white"
        )
        label_titulo.pack(pady=25)
        
        # ============================================
        # FRAME PRINCIPAL
        # ============================================
        frame_principal = tk.Frame(self.ventana, padx=20, pady=20)
        frame_principal.pack(fill=tk.BOTH, expand=True)
        
        # ============================================
        # SECCIÓN: CREDENCIALES
        # ============================================
        label_seccion1 = tk.Label(
            frame_principal,
            text="🔐 CREDENCIALES DE ACCESO",
            font=("Arial", 12, "bold"),
            fg="#0063AE"
        )
        label_seccion1.grid(row=0, column=0, columnspan=3, sticky="w", pady=(0, 10))
        
        # Usuario
        tk.Label(frame_principal, text="Usuario:", font=("Arial", 10)).grid(
            row=1, column=0, sticky="e", padx=5, pady=5
        )
        entry_usuario = tk.Entry(
            frame_principal,
            textvariable=self.usuario,
            font=("Arial", 10),
            width=30
        )
        entry_usuario.grid(row=1, column=1, sticky="w", padx=5, pady=5)
        
        # Contraseña
        tk.Label(frame_principal, text="Contraseña:", font=("Arial", 10)).grid(
            row=2, column=0, sticky="e", padx=5, pady=5
        )
        entry_password = tk.Entry(
            frame_principal,
            textvariable=self.password,
            font=("Arial", 10),
            width=30,
            show="●"
        )
        entry_password.grid(row=2, column=1, sticky="w", padx=5, pady=5)
        
        # ============================================
        # SECCIÓN: ARCHIVO EXCEL
        # ============================================
        label_seccion2 = tk.Label(
            frame_principal,
            text="📂 ARCHIVO DE DENUNCIAS",
            font=("Arial", 12, "bold"),
            fg="#0063AE"
        )
        label_seccion2.grid(row=3, column=0, columnspan=3, sticky="w", pady=(20, 10))
        
        # Ruta del archivo
        tk.Label(frame_principal, text="Archivo Excel:", font=("Arial", 10)).grid(
            row=4, column=0, sticky="e", padx=5, pady=5
        )
        entry_archivo = tk.Entry(
            frame_principal,
            textvariable=self.ruta_archivo,
            font=("Arial", 10),
            width=30,
            state="readonly"
        )
        entry_archivo.grid(row=4, column=1, sticky="w", padx=5, pady=5)
        
        # Botón examinar
        btn_examinar = tk.Button(
            frame_principal,
            text="📁 Examinar",
            command=self.seleccionar_archivo,
            font=("Arial", 10),
            bg="#0063AE",
            fg="white",
            cursor="hand2",
            width=12
        )
        btn_examinar.grid(row=4, column=2, padx=5, pady=5)
        
        # ============================================
        # SECCIÓN: CONSOLA DE REGISTRO
        # ============================================
        label_seccion3 = tk.Label(
            frame_principal,
            text="📋 REGISTRO DE ACTIVIDAD",
            font=("Arial", 12, "bold"),
            fg="#0063AE"
        )
        label_seccion3.grid(row=5, column=0, columnspan=3, sticky="w", pady=(20, 10))
        
        # Área de texto con scroll
        self.consola = scrolledtext.ScrolledText(
            frame_principal,
            width=85,
            height=12,
            font=("Courier", 9),
            bg="#f5f5f5",
            fg="#333333",
            state="disabled"
        )
        self.consola.grid(row=6, column=0, columnspan=3, pady=5)
        
        # ============================================
        # SECCIÓN: BOTONES DE CONTROL
        # ============================================
        frame_botones = tk.Frame(frame_principal)
        frame_botones.grid(row=7, column=0, columnspan=3, pady=20)
        
        # Botón INICIAR
        self.btn_iniciar = tk.Button(
            frame_botones,
            text="▶️ INICIAR PROCESO",
            command=self.iniciar_proceso,
            font=("Arial", 11, "bold"),
            bg="#28a745",
            fg="white",
            cursor="hand2",
            width=20,
            height=2
        )
        self.btn_iniciar.pack(side=tk.LEFT, padx=10)
        
        # Botón CANCELAR
        self.btn_cancelar = tk.Button(
            frame_botones,
            text="⏹️ CANCELAR",
            command=self.cancelar_proceso,
            font=("Arial", 11, "bold"),
            bg="#dc3545",
            fg="white",
            cursor="hand2",
            width=20,
            height=2,
            state="disabled"
        )
        self.btn_cancelar.pack(side=tk.LEFT, padx=10)
        
        # ============================================
        # BARRA DE ESTADO
        # ============================================
        self.label_estado = tk.Label(
            self.ventana,
            text="Estado: Esperando...",
            font=("Arial", 9),
            bg="#f0f0f0",
            anchor="w",
            padx=10
        )
        self.label_estado.pack(side=tk.BOTTOM, fill=tk.X)
        
    def seleccionar_archivo(self):
        """Abre diálogo para seleccionar archivo Excel"""
        archivo = filedialog.askopenfilename(
            title="Seleccionar archivo Excel",
            filetypes=[
                ("Archivos Excel", "*.xlsx *.xls"),
                ("Todos los archivos", "*.*")
            ]
        )
        if archivo:
            self.ruta_archivo.set(archivo)
            self.escribir_consola(f"✅ Archivo seleccionado: {os.path.basename(archivo)}\n")
    
    def escribir_consola(self, mensaje):
        """Escribe mensaje en la consola de registro"""
        self.consola.config(state="normal")
        self.consola.insert(tk.END, mensaje)
        self.consola.see(tk.END)
        self.consola.config(state="disabled")
        self.ventana.update()
    
    def limpiar_consola(self):
        """Limpia la consola de registro"""
        self.consola.config(state="normal")
        self.consola.delete(1.0, tk.END)
        self.consola.config(state="disabled")
    
    def validar_campos(self):
        """Valida que todos los campos estén completos"""
        if not self.usuario.get().strip():
            messagebox.showerror("Error", "Por favor ingrese el usuario")
            return False
        
        if not self.password.get().strip():
            messagebox.showerror("Error", "Por favor ingrese la contraseña")
            return False
        
        if not self.ruta_archivo.get().strip():
            messagebox.showerror("Error", "Por favor seleccione un archivo Excel")
            return False
        
        if not os.path.exists(self.ruta_archivo.get()):
            messagebox.showerror("Error", "El archivo seleccionado no existe")
            return False
        
        return True
    
    def iniciar_proceso(self):
        """Inicia el proceso de registro de denuncias"""
        if not self.validar_campos():
            return
        
        # Confirmar inicio
        respuesta = messagebox.askyesno(
            "Confirmar",
            "¿Desea iniciar el proceso de registro de denuncias?\n\n"
            "El proceso puede tardar varios minutos dependiendo\n"
            "de la cantidad de denuncias."
        )
        
        if not respuesta:
            return
        
        # Limpiar consola
        self.limpiar_consola()
        
        # Cambiar estado de botones
        self.btn_iniciar.config(state="disabled")
        self.btn_cancelar.config(state="normal")
        self.proceso_activo = True
        self.label_estado.config(text="Estado: Procesando...", bg="#ffc107")
        
        # Iniciar proceso en hilo separado
        self.hilo_proceso = threading.Thread(target=self.ejecutar_bot, daemon=True)
        self.hilo_proceso.start()
    
    def cancelar_proceso(self):
        """Cancela el proceso en ejecución"""
        respuesta = messagebox.askyesno(
            "Confirmar Cancelación",
            "¿Está seguro que desea cancelar el proceso?\n\n"
            "Las denuncias procesadas hasta el momento\n"
            "se mantendrán registradas."
        )
        
        if respuesta:
            self.proceso_activo = False
            self.escribir_consola("\n⚠️ CANCELANDO PROCESO...\n")
            self.label_estado.config(text="Estado: Cancelando...", bg="#dc3545")
            
            # Cerrar navegador si existe
            if self.bot and self.bot.driver:
                try:
                    self.bot.driver.quit()
                except:
                    pass
            
            self.btn_iniciar.config(state="normal")
            self.btn_cancelar.config(state="disabled")
            self.label_estado.config(text="Estado: Proceso cancelado", bg="#f0f0f0")
    
    def ejecutar_bot(self):
        """Ejecuta el bot en un hilo separado"""
        try:
            # Crear instancia del bot
            self.bot = BotDenunciasSUNAT(
                archivo_excel=self.ruta_archivo.get(),
                usuario=self.usuario.get(),
                password=self.password.get(),
                interfaz=self
            )
            
            # Ejecutar proceso
            self.bot.ejecutar()
            
        except Exception as e:
            self.escribir_consola(f"\n❌ ERROR CRÍTICO: {str(e)}\n")
            messagebox.showerror("Error", f"Error crítico en el proceso:\n{str(e)}")
        
        finally:
            # Restaurar botones
            self.btn_iniciar.config(state="normal")
            self.btn_cancelar.config(state="disabled")
            self.proceso_activo = False
            
            if hasattr(self.bot, 'denuncias_exitosas'):
                self.label_estado.config(
                    text=f"Estado: Completado - {self.bot.denuncias_exitosas} denuncias exitosas",
                    bg="#28a745"
                )
            else:
                self.label_estado.config(text="Estado: Finalizado", bg="#f0f0f0")
    
    def ejecutar(self):
        """Ejecuta el loop principal de la interfaz"""
        self.ventana.mainloop()


# ============================================
# CLASE PRINCIPAL DEL BOT
# ============================================
class BotDenunciasSUNAT:
    
    def __init__(self, archivo_excel, usuario, password, interfaz):
        """Inicializa el bot con la configuración necesaria"""
        self.archivo_excel = archivo_excel
        self.USUARIO = usuario
        self.PASSWORD = password
        self.interfaz = interfaz
        
        self.driver = None
        self.wait = None
        self.denuncias_exitosas = 0
        self.denuncias_fallidas = 0
        
        # URL
        self.URL_LOGIN = "https://intranet.sunat.peru/cl-at-iamenu/"
        
        self.log("Bot inicializado correctamente")
    
    def log(self, mensaje):
        """Escribe en la consola de la interfaz"""
        timestamp = datetime.now().strftime("%H:%M:%S")
        self.interfaz.escribir_consola(f"[{timestamp}] {mensaje}\n")
    
    # ============================================
    # MÉTODOS DE INICIALIZACIÓN
    # ============================================
    
    def iniciar_navegador(self):
        """Inicia el navegador Chrome"""
        try:
            self.log("Iniciando navegador...")
            
            chrome_options = Options()
            chrome_options.add_argument('--start-maximized')
            chrome_options.add_argument('--disable-blink-features=AutomationControlled')
            chrome_options.add_experimental_option("excludeSwitches", ["enable-automation"])
            chrome_options.add_experimental_option('useAutomationExtension', False)
            
            self.driver = webdriver.Chrome(options=chrome_options)
            self.wait = WebDriverWait(self.driver, 20)
            
            self.log("✅ Navegador iniciado correctamente")
            return True
        except Exception as e:
            self.log(f"❌ Error al iniciar navegador: {str(e)}")
            return False
    
    def cerrar_navegador(self):
        """Cierra el navegador"""
        try:
            if self.driver:
                self.driver.quit()
                self.log("Navegador cerrado")
        except:
            pass
    
    # ============================================
    # LOGIN Y NAVEGACIÓN
    # ============================================
    
    def hacer_login(self):
        """Realiza el login en el sistema"""
        try:
            self.log("Realizando login...")
            self.driver.get(self.URL_LOGIN)
            time.sleep(2)

            # Usuario
            campo_usuario = self.wait.until(
                EC.presence_of_element_located((By.NAME, "cuenta"))
            )
            campo_usuario.clear()
            campo_usuario.send_keys(self.USUARIO)

            # Password
            campo_password = self.driver.find_element(By.NAME, "password")
            campo_password.clear()
            campo_password.send_keys(self.PASSWORD)

            # Click Iniciar Sesion
            boton_iniciar = self.driver.find_element(By.XPATH, "//input[@type='button' and @value='Iniciar Sesion']")
            boton_iniciar.click()

            time.sleep(3)
            self.log("✅ Login exitoso")
            return True

        except Exception as e:
            self.log(f"❌ Error en login: {str(e)}")
            return False
    
    def navegar_a_denuncias(self):
        """Navega al módulo de Denuncias"""
        try:
            self.log("Navegando a Denuncias...")
            time.sleep(3)
            
            # Click Tributarios
            link_tributarios = self.wait.until(
                EC.element_to_be_clickable((By.LINK_TEXT, "Tributarios"))
            )
            link_tributarios.click()
            time.sleep(2)
            
            # Click Denuncias
            link_denuncias = self.wait.until(
                EC.element_to_be_clickable((By.LINK_TEXT, "Denuncias"))
            )
            link_denuncias.click()
            time.sleep(2)
            
            self.log("✅ Navegación exitosa")
            return True
            
        except Exception as e:
            self.log(f"❌ Error al navegar: {str(e)}")
            return False
    
    def navegar_a_formulario_registro(self):
        """Navega al formulario de Registro"""
        try:
            # Cambiar a iframe
            self.wait.until(
                EC.frame_to_be_available_and_switch_to_it((By.NAME, "iframeApplication"))
            )
            time.sleep(2)
            
            # Click menú Denuncias
            link_menu = self.wait.until(
                EC.element_to_be_clickable((By.ID, "5.5.2.1"))
            )
            link_menu.click()
            time.sleep(1)
            
            # Click Registro de Denuncias
            link_registro = self.wait.until(
                EC.element_to_be_clickable((By.XPATH, "//a[text()='Registro de Denuncias']"))
            )
            link_registro.click()
            time.sleep(3)
            
            return True
            
        except Exception as e:
            self.log(f"❌ Error al abrir formulario: {str(e)}")
            return False
    
    # ============================================
    # SECCIÓN 1: IDENTIFICACIÓN DEL DENUNCIADO
    # ============================================
    
    def llenar_seccion1_identificacion(self, datos):
        """Llena Sección 1"""
        try:
            self.log("📝 Llenando Sección 1...")
            
            # Tipo Documento
            if 'TIPO' in datos and pd.notna(datos['TIPO']):
                valor = str(datos['TIPO']).strip()
                select_tipo = Select(self.driver.find_element(By.NAME, "tipoDocumento"))
                select_tipo.select_by_visible_text(valor)
                time.sleep(0.5)
            
            # Número
            if 'NRO' in datos and pd.notna(datos['NRO']):
                valor = str(int(datos['NRO'])) if isinstance(datos['NRO'], float) else str(datos['NRO'])
                campo_numero = self.driver.find_element(By.NAME, "numero")
                campo_numero.clear()
                campo_numero.send_keys(valor.strip())
                time.sleep(0.5)
            
            # Buscar
            boton_buscar = self.driver.find_element(By.XPATH, "//button[text()='Buscar']")
            boton_buscar.click()
            time.sleep(2)
            
            # Siguiente
            boton_siguiente = self.driver.find_element(By.XPATH, "//button[text()='Siguiente']")
            boton_siguiente.click()
            time.sleep(2)
            
            self.log("✅ Sección 1 completada")
            return True
            
        except Exception as e:
            self.log(f"❌ Error en Sección 1: {str(e)}")
            return False
    
    # ============================================
    # SECCIÓN 2: ATENCIÓN DE DENUNCIAS
    # ============================================
    
    def llenar_seccion2_atencion_denuncias(self, datos):
        """Llena Sección 2"""
        try:
            self.log("📝 Llenando Sección 2...")
            
            # Modalidad Evasión
            if 'Modalidad de evasion' in datos and pd.notna(datos['Modalidad de evasion']):
                valor = str(datos['Modalidad de evasion']).strip()
                select_modalidad = Select(self.driver.find_element(By.NAME, "modalidadEvasion"))
                select_modalidad.select_by_visible_text(valor)
                time.sleep(0.5)
            
            # Sub Modalidad (opcional)
            if 'Submodalidad' in datos and pd.notna(datos['Submodalidad']):
                valor = str(datos['Submodalidad']).strip()
                if valor != "" and valor != "-":
                    select_sub = Select(self.driver.find_element(By.NAME, "subModalidad"))
                    select_sub.select_by_visible_text(valor)
                    time.sleep(0.5)
            
            # Tipo de Denuncia
            if 'Tipo de denuncia' in datos and pd.notna(datos['Tipo de denuncia']):
                valor = str(datos['Tipo de denuncia']).strip()
                radio_xpath = f"//input[@type='radio' and contains(@value, '{valor}')]"
                radio_button = self.driver.find_element(By.XPATH, radio_xpath)
                radio_button.click()
                time.sleep(0.5)
            
            # Fecha SID (opcional)
            if 'Fecha SID' in datos and pd.notna(datos['Fecha SID']):
                fecha_formateada = self.convertir_fecha(datos['Fecha SID'])
                if fecha_formateada:
                    campo_fecha = self.driver.find_element(By.NAME, "fechaSID")
                    campo_fecha.clear()
                    campo_fecha.send_keys(fecha_formateada)
                    time.sleep(0.5)
            
            # Detalle
            if 'Detalle de la denuncia' in datos and pd.notna(datos['Detalle de la denuncia']):
                valor = str(datos['Detalle de la denuncia']).strip()
                campo_detalle = self.driver.find_element(By.NAME, "detalleDenuncia")
                campo_detalle.clear()
                campo_detalle.send_keys(valor)
                time.sleep(0.5)
            
            # Rango Desde
            if 'Desde' in datos and pd.notna(datos['Desde']):
                fecha = self.extraer_mes_anio(datos['Desde'])
                if fecha:
                    select_del_mes = Select(self.driver.find_element(By.NAME, "delMes"))
                    select_del_mes.select_by_visible_text(fecha['mes'])
                    
                    select_del_anio = Select(self.driver.find_element(By.NAME, "delAnio"))
                    select_del_anio.select_by_visible_text(str(fecha['anio']))
                    time.sleep(0.5)
            
            # Rango Hasta
            if 'Hasta' in datos and pd.notna(datos['Hasta']):
                fecha = self.extraer_mes_anio(datos['Hasta'])
                if fecha:
                    select_al_mes = Select(self.driver.find_element(By.NAME, "alMes"))
                    select_al_mes.select_by_visible_text(fecha['mes'])
                    
                    select_al_anio = Select(self.driver.find_element(By.NAME, "alAnio"))
                    select_al_anio.select_by_visible_text(str(fecha['anio']))
                    time.sleep(0.5)
            
            # Pruebas Ofrecidas
            if 'PRUEBA' in datos and pd.notna(datos['PRUEBA']):
                valor = str(datos['PRUEBA']).strip().lower()
                
                if valor in ["si", "sí"]:
                    radio_si = self.driver.find_element(By.XPATH, "//input[@type='radio' and @value='Si']")
                    radio_si.click()
                    time.sleep(1)
                    
                    # Tipo de Pruebas
                    if 'EN CASO DE SI' in datos and pd.notna(datos['EN CASO DE SI']):
                        valor_prueba = str(datos['EN CASO DE SI']).strip()
                        if valor_prueba != "" and valor_prueba != "-":
                            select_pruebas = Select(self.driver.find_element(By.NAME, "tipoPruebas"))
                            select_pruebas.select_by_visible_text(valor_prueba)
                            time.sleep(0.5)
                            
                            # Si es "Otros, detalle"
                            if "otros" in valor_prueba.lower():
                                if 'OTRO, DETALLE' in datos and pd.notna(datos['OTRO, DETALLE']):
                                    valor_otros = str(datos['OTRO, DETALLE']).strip()
                                    if valor_otros != "" and valor_otros != "-":
                                        campo_otros = self.driver.find_element(By.NAME, "otrosDetalle")
                                        campo_otros.clear()
                                        campo_otros.send_keys(valor_otros)
                                        time.sleep(0.5)
                else:
                    radio_no = self.driver.find_element(By.XPATH, "//input[@type='radio' and @value='No']")
                    radio_no.click()
                    time.sleep(0.5)
            
            # Siguiente
            boton_siguiente = self.driver.find_element(By.XPATH, "//button[text()='Siguiente']")
            boton_siguiente.click()
            time.sleep(2)
            
            self.log("✅ Sección 2 completada")
            return True
            
        except Exception as e:
            self.log(f"❌ Error en Sección 2: {str(e)}")
            return False
    
    # ============================================
    # SECCIÓN 3: IDENTIFICACIÓN DEL DENUNCIANTE
    # ============================================
    
    def llenar_seccion3_identificacion_denunciante(self, datos):
        """Llena Sección 3"""
        try:
            self.log("📝 Llenando Sección 3...")
            
            # Tipo Denunciante
            # Nota: Usar índice si hay dos columnas "TIPO"
            if 'TIPO' in datos:
                valor = str(datos['TIPO']).strip()
                select_tipo = Select(self.driver.find_element(By.NAME, "tipoDenunciante"))
                select_tipo.select_by_visible_text(valor)
                time.sleep(0.5)
            
            # RUC/DNI
            if 'ruc denunciante' in datos and pd.notna(datos['ruc denunciante']):
                valor = str(int(datos['ruc denunciante'])) if isinstance(datos['ruc denunciante'], float) else str(datos['ruc denunciante'])
                campo_numero = self.driver.find_element(By.NAME, "numeroNombre")
                campo_numero.clear()
                campo_numero.send_keys(valor.strip())
                time.sleep(0.5)
            
            # Teléfono
            if 'teléfono' in datos and pd.notna(datos['teléfono']):
                valor = str(int(datos['teléfono'])) if isinstance(datos['teléfono'], float) else str(datos['teléfono'])
                campo_telefono = self.driver.find_element(By.NAME, "telefono")
                campo_telefono.clear()
                campo_telefono.send_keys(valor.strip())
                time.sleep(0.5)
            
            # Correo
            if 'correo electrónico' in datos and pd.notna(datos['correo electrónico']):
                valor = str(datos['correo electrónico']).strip()
                campo_correo = self.driver.find_element(By.NAME, "correoElectronico")
                campo_correo.clear()
                campo_correo.send_keys(valor)
                time.sleep(0.5)
            
            # Departamento
            if 'Departamento' in datos and pd.notna(datos['Departamento']):
                valor = str(datos['Departamento']).strip()
                select_depto = Select(self.driver.find_element(By.NAME, "departamento"))
                select_depto.select_by_visible_text(valor)
                time.sleep(1)
            
            # Provincia
            if 'Provincia' in datos and pd.notna(datos['Provincia']):
                valor = str(datos['Provincia']).strip()
                select_prov = Select(self.driver.find_element(By.NAME, "provincia"))
                select_prov.select_by_visible_text(valor)
                time.sleep(1)
            
            # Distrito
            if 'Distrito' in datos and pd.notna(datos['Distrito']):
                valor = str(datos['Distrito']).strip()
                select_dist = Select(self.driver.find_element(By.NAME, "distrito"))
                select_dist.select_by_visible_text(valor)
                time.sleep(0.5)
            
            # Vía
            if 'Via' in datos and pd.notna(datos['Via']):
                valor = str(datos['Via']).strip()
                select_via = Select(self.driver.find_element(By.NAME, "via"))
                select_via.select_by_visible_text(valor)
                time.sleep(0.5)
            
            # Nombre de Vía
            if 'Relleno de Via' in datos and pd.notna(datos['Relleno de Via']):
                valor = str(datos['Relleno de Via']).strip()
                campo_nombre = self.driver.find_element(By.NAME, "nombreVia")
                campo_nombre.clear()
                campo_nombre.send_keys(valor)
                time.sleep(0.5)
            
            # N°/Mzn./Km.
            if 'N.°' in datos and pd.notna(datos['N.°']):
                valor = str(datos['N.°']).strip()
                if valor != "" and valor != "-":
                    campo_numero = self.driver.find_element(By.NAME, "numeroMznKm")
                    campo_numero.clear()
                    campo_numero.send_keys(valor)
                    time.sleep(0.5)
            
            # Dpto/Int
            if 'Dpto' in datos and pd.notna(datos['Dpto']):
                valor = str(datos['Dpto']).strip()
                if valor != "" and valor != "-":
                    campo_dpto = self.driver.find_element(By.NAME, "dptoIntLote")
                    campo_dpto.clear()
                    campo_dpto.send_keys(valor)
                    time.sleep(0.5)
            
            # Zona
            if 'Zona' in datos and pd.notna(datos['Zona']):
                valor = str(datos['Zona']).strip()
                if valor != "" and valor != "-":
                    try:
                        select_zona = Select(self.driver.find_element(By.NAME, "zona"))
                        select_zona.select_by_visible_text(valor)
                    except:
                        campo_zona = self.driver.find_element(By.NAME, "zona")
                        campo_zona.clear()
                        campo_zona.send_keys(valor)
                    time.sleep(0.5)
            
            # GRABAR
            boton_grabar = self.driver.find_element(By.XPATH, "//button[text()='Grabar']")
            boton_grabar.click()
            time.sleep(3)
            
            self.log("✅ Sección 3 completada - DENUNCIA GRABADA")
            return True
            
        except Exception as e:
            self.log(f"❌ Error en Sección 3: {str(e)}")
            return False
    
    # ============================================
    # MÉTODOS AUXILIARES
    # ============================================
    
    def convertir_fecha(self, fecha):
        """Convierte fecha de Excel a formato DD/MM/YYYY"""
        try:
            if isinstance(fecha, (int, float)):
                fecha_base = datetime(1899, 12, 30)
                fecha_real = fecha_base + timedelta(days=int(fecha))
                return fecha_real.strftime('%d/%m/%Y')
            elif isinstance(fecha, pd.Timestamp):
                return fecha.strftime('%d/%m/%Y')
            else:
                return str(fecha)
        except:
            return None
    
    def extraer_mes_anio(self, fecha):
        """Extrae mes y año"""
        try:
            if isinstance(fecha, (int, float)):
                fecha_base = datetime(1899, 12, 30)
                fecha_real = fecha_base + timedelta(days=int(fecha))
            elif isinstance(fecha, pd.Timestamp):
                fecha_real = fecha
            else:
                from dateutil import parser
                fecha_real = parser.parse(str(fecha))
            
            meses = {
                1: "Enero", 2: "Febrero", 3: "Marzo", 4: "Abril",
                5: "Mayo", 6: "Junio", 7: "Julio", 8: "Agosto",
                9: "Septiembre", 10: "Octubre", 11: "Noviembre", 12: "Diciembre"
            }
            
            return {'mes': meses[fecha_real.month], 'anio': fecha_real.year}
        except:
            return None
    
    # ============================================
    # PROCESO PRINCIPAL
    # ============================================
    
    def procesar_una_denuncia(self, datos_fila, numero_fila):
        """Procesa una denuncia completa"""
        try:
            # Verificar si el proceso fue cancelado
            if not self.interfaz.proceso_activo:
                self.log("⚠️ Proceso cancelado por el usuario")
                return False
            
            self.log(f"\n{'='*50}")
            self.log(f"📋 PROCESANDO DENUNCIA #{numero_fila}")
            self.log(f"{'='*50}")
            
            # Navegar al formulario
            if not self.navegar_a_formulario_registro():
                return False
            
            # Verificar cancelación
            if not self.interfaz.proceso_activo:
                return False
            
            # Sección 1
            if not self.llenar_seccion1_identificacion(datos_fila):
                return False
            
            # Verificar cancelación
            if not self.interfaz.proceso_activo:
                return False
            
            # Sección 2
            if not self.llenar_seccion2_atencion_denuncias(datos_fila):
                return False
            
            # Verificar cancelación
            if not self.interfaz.proceso_activo:
                return False
            
            # Sección 3
            if not self.llenar_seccion3_identificacion_denunciante(datos_fila):
                return False
            
            self.log(f"🎉 ¡DENUNCIA #{numero_fila} REGISTRADA EXITOSAMENTE!")
            self.denuncias_exitosas += 1
            
            # Volver al iframe principal
            self.driver.switch_to.default_content()
            time.sleep(2)
            
            return True
            
        except Exception as e:
            self.log(f"❌ Error en denuncia #{numero_fila}: {str(e)}")
            self.denuncias_fallidas += 1
            
            try:
                self.driver.switch_to.default_content()
            except:
                pass
            
            return False
    
    def ejecutar(self):
        """Método principal"""
        try:
            self.log("="*50)
            self.log("🤖 INICIANDO PROCESO")
            self.log("="*50)
            
            # Leer Excel
            self.log(f"📂 Leyendo archivo: {os.path.basename(self.archivo_excel)}")
            df = pd.read_excel(self.archivo_excel)
            total = len(df)
            self.log(f"✅ {total} denuncias encontradas\n")
            
            # Iniciar navegador
            if not self.iniciar_navegador():
                return
            
            # Login
            if not self.hacer_login():
                self.cerrar_navegador()
                return
            
            # Navegar a Denuncias
            if not self.navegar_a_denuncias():
                self.cerrar_navegador()
                return
            
            # Procesar cada denuncia
            for index, fila in df.iterrows():
                # Verificar cancelación
                if not self.interfaz.proceso_activo:
                    self.log("\n⚠️ PROCESO CANCELADO POR EL USUARIO")
                    break
                
                numero_fila = index + 2
                self.procesar_una_denuncia(fila, numero_fila)
                time.sleep(2)
            
            # Resumen
            self.log("\n" + "="*50)
            self.log("📊 RESUMEN FINAL")
            self.log("="*50)
            self.log(f"✅ Exitosas: {self.denuncias_exitosas}/{total}")
            self.log(f"❌ Fallidas: {self.denuncias_fallidas}/{total}")
            if total > 0:
                tasa = (self.denuncias_exitosas/total)*100
                self.log(f"📈 Tasa de éxito: {tasa:.2f}%")
            self.log("="*50)
            
            # Cerrar
            self.log("\nCerrando navegador...")
            time.sleep(3)
            self.cerrar_navegador()
            
            self.log("\n🏁 ¡PROCESO COMPLETADO!")
            
            # Mostrar mensaje final
            messagebox.showinfo(
                "Proceso Completado",
                f"Proceso finalizado exitosamente\n\n"
                f"✅ Denuncias exitosas: {self.denuncias_exitosas}\n"
                f"❌ Denuncias fallidas: {self.denuncias_fallidas}\n"
                f"📊 Total procesadas: {total}"
            )
            
        except Exception as e:
            self.log(f"\n❌ ERROR CRÍTICO: {str(e)}")
            messagebox.showerror("Error Crítico", f"Error:\n{str(e)}")
            self.cerrar_navegador()


# ============================================
# PUNTO DE ENTRADA
# ============================================
if __name__ == "__main__":
    app = InterfazBot()
    app.ejecutar()