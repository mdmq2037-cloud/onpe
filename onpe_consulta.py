#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
ONPE Consulta Electoral Masiva v2.0
Consulta masiva de DNIs en: https://consultaelectoral.onpe.gob.pe/inicio

Funcionalidades:
  - Carga masiva de DNIs desde CSV, TXT o Excel
  - Consulta automática con Playwright (navegador real)
  - Captura del endpoint API para consultas directas (sin navegador)
  - Almacenamiento en SQLite
  - Exportación a CSV
  - Interfaz gráfica con progreso en tiempo real
"""

import tkinter as tk
from tkinter import ttk, filedialog, messagebox
import threading
import sqlite3
import csv
import os
import time
import json
import re
from datetime import datetime
from pathlib import Path
import queue
import random

# ─────────────────────────────────────────────
# CONFIGURACIÓN
# ─────────────────────────────────────────────
ONPE_URL      = "https://consultaelectoral.onpe.gob.pe/inicio"
DB_FILE       = os.path.join(os.path.dirname(os.path.abspath(__file__)), "onpe_consultas.db")
DELAY_DEFAULT = 10.0  # segundos entre consultas (respetar rate-limit ONPE)

# ─────────────────────────────────────────────
# BASE DE DATOS
# ─────────────────────────────────────────────
class Database:
    def __init__(self, path=DB_FILE):
        self.path = path
        self._init()

    def _init(self):
        with sqlite3.connect(self.path, check_same_thread=False) as c:
            c.execute("""
                CREATE TABLE IF NOT EXISTS consultas (
                    id           INTEGER PRIMARY KEY AUTOINCREMENT,
                    dni          TEXT    NOT NULL UNIQUE,
                    nombres      TEXT,
                    region       TEXT,
                    provincia    TEXT,
                    distrito     TEXT,
                    miembro_mesa INTEGER DEFAULT 0,
                    local_vot    TEXT,
                    direccion    TEXT,
                    referencia   TEXT,
                    nro_mesa     TEXT,
                    nro_orden    TEXT,
                    estado       TEXT    DEFAULT 'pendiente',
                    error_msg    TEXT,
                    consultado   TEXT
                )
            """)
            c.commit()

    def upsert(self, r):
        with sqlite3.connect(self.path, check_same_thread=False) as c:
            c.execute("""
                INSERT OR REPLACE INTO consultas
                (dni, nombres, region, provincia, distrito, miembro_mesa,
                 local_vot, direccion, referencia, nro_mesa, nro_orden,
                 estado, error_msg, consultado)
                VALUES (?,?,?,?,?,?,?,?,?,?,?,?,?,?)
            """, (
                r.get('dni',''),        r.get('nombres',''),
                r.get('region',''),     r.get('provincia',''),
                r.get('distrito',''),   1 if r.get('miembro_mesa') else 0,
                r.get('local_vot',''),  r.get('direccion',''),
                r.get('referencia',''), r.get('nro_mesa',''),
                r.get('nro_orden',''),  r.get('estado','ok'),
                r.get('error_msg',''),  datetime.now().strftime('%Y-%m-%d %H:%M:%S'),
            ))
            c.commit()

    def get_all(self):
        with sqlite3.connect(self.path, check_same_thread=False) as c:
            c.row_factory = sqlite3.Row
            return [dict(r) for r in c.execute(
                "SELECT * FROM consultas ORDER BY id DESC"
            ).fetchall()]

    def stats(self):
        with sqlite3.connect(self.path, check_same_thread=False) as c:
            c.row_factory = sqlite3.Row
            row = c.execute("""
                SELECT COUNT(*) total,
                       SUM(CASE WHEN estado='ok'    THEN 1 ELSE 0 END) ok,
                       SUM(CASE WHEN miembro_mesa=1 THEN 1 ELSE 0 END) miembros,
                       SUM(CASE WHEN estado='error' THEN 1 ELSE 0 END) errores
                FROM consultas
            """).fetchone()
            return dict(row) if row else {}

    def export_csv(self, path):
        rows = self.get_all()
        if not rows:
            return False
        with open(path, 'w', newline='', encoding='utf-8-sig') as f:
            dw = csv.DictWriter(f, fieldnames=rows[0].keys())
            dw.writeheader()
            dw.writerows(rows)
        return True

    def clear(self):
        with sqlite3.connect(self.path, check_same_thread=False) as c:
            c.execute("DELETE FROM consultas")
            c.commit()

    def pending_dnis(self):
        """DNIs ya consultados, para evitar duplicados."""
        with sqlite3.connect(self.path, check_same_thread=False) as c:
            rows = c.execute("SELECT dni FROM consultas WHERE estado='ok'").fetchall()
            return {r[0] for r in rows}


# ─────────────────────────────────────────────
# SCRAPER CON UNDETECTED-CHROMEDRIVER
# ─────────────────────────────────────────────
class ONPEScraper:
    """
    Usa undetected-chromedriver para controlar Chrome real
    evitando la detección de reCAPTCHA v3.
    """

    def __init__(self, headless=False, log_fn=None):
        self.headless = headless
        self.log      = log_fn or print
        self._driver  = None

    # ── Inicio del navegador ──────────────────
    def start(self):
        import undetected_chromedriver as uc
        import subprocess
        self.log("Iniciando Google Chrome (undetected)...")

        # Auto-detectar versión de Chrome instalado (HKLM y HKCU)
        chrome_ver = None
        for reg_key in [
            r'HKLM\SOFTWARE\Google\Chrome\BLBeacon',
            r'HKCU\SOFTWARE\Google\Chrome\BLBeacon',
            r'HKLM\SOFTWARE\WOW6432Node\Google\Chrome\BLBeacon',
        ]:
            try:
                out = subprocess.check_output(
                    ['reg', 'query', reg_key, '/v', 'version'],
                    stderr=subprocess.DEVNULL).decode()
                ver_str = re.search(r'(\d+)\.\d+', out)
                if ver_str:
                    chrome_ver = int(ver_str.group(1))
                    break
            except Exception:
                continue

        if chrome_ver:
            self.log(f"Chrome versión detectada: {chrome_ver}")
        else:
            self.log("No se detectó versión de Chrome, se usará auto-detección.")

        # Perfil persistente: guarda cookies/historial entre sesiones (mejora reCAPTCHA)
        profile_dir = os.path.join(os.path.dirname(os.path.abspath(__file__)), "chrome_profile")

        # Crear options frescos (uc no permite reutilizar el mismo objeto)
        def _mk_options():
            o = uc.ChromeOptions()
            o.add_argument("--window-size=1366,768")
            o.add_argument("--lang=es-PE")
            if self.headless:
                o.add_argument("--headless=new")
            return o

        # Siempre incluir version_main para evitar mismatch ChromeDriver/Chrome
        ver_kwargs = {'version_main': chrome_ver} if chrome_ver else {}

        # Intento 1: con perfil persistente
        try:
            self._driver = uc.Chrome(options=_mk_options(), user_data_dir=profile_dir, **ver_kwargs)
        except Exception:
            self.log("Perfil bloqueado, iniciando sin perfil persistente...")
            # Intento 2: sin perfil (el perfil puede estar bloqueado por otro Chrome abierto)
            try:
                self._driver = uc.Chrome(options=_mk_options(), **ver_kwargs)
            except Exception as e2:
                raise RuntimeError(f"No se pudo iniciar Chrome: {e2}") from e2

        self._driver.get(ONPE_URL)
        time.sleep(random.uniform(4.0, 6.0))
        self._human_behavior()
        time.sleep(random.uniform(2.0, 3.0))
        self.log("Navegador listo.")

    # ── Consulta de un DNI ────────────────────
    def query_dni(self, dni):
        dni = str(dni).strip().zfill(8)
        return self._query_via_browser(dni)

    # ── Vía navegador ─────────────────────────
    def _query_via_browser(self, dni):
        from selenium.webdriver.common.by import By
        from selenium.webdriver.support.ui import WebDriverWait
        from selenium.webdriver.support import expected_conditions as EC
        from selenium.webdriver.common.keys import Keys

        r = self._empty_record(dni)
        try:
            wait = WebDriverWait(self._driver, 25)

            # ── Buscar campo DNI ─────────────
            # Intentar varios selectores con timeout generoso (ONPE SPA puede tardar)
            dni_input = None
            for css in ['input[type="tel"]', 'input[maxlength="8"]',
                        'input[placeholder*="DNI"]', 'input[placeholder*="dni"]',
                        'input']:
                try:
                    dni_input = WebDriverWait(self._driver, 25).until(
                        EC.element_to_be_clickable((By.CSS_SELECTOR, css)))
                    break
                except Exception:
                    continue

            if not dni_input:
                raise RuntimeError("No se encontró el campo DNI en la página")

            # ── Click JS en input (activa Angular, evita ElementClickInterceptedException) ──
            time.sleep(random.uniform(0.5, 1.0))
            self._driver.execute_script(
                "arguments[0].scrollIntoView({block:'center'}); arguments[0].click();",
                dni_input
            )
            time.sleep(random.uniform(0.2, 0.4))
            # Limpiar con teclas para que Angular detecte el cambio
            dni_input.send_keys(Keys.CONTROL + 'a')
            time.sleep(0.1)
            dni_input.send_keys(Keys.DELETE)
            time.sleep(0.2)
            # Escribir cada dígito con pausa natural
            for ch in dni:
                dni_input.send_keys(ch)
                time.sleep(random.uniform(0.08, 0.18))
            self.log(f"  DNI {dni} ingresado")

            # ── Clic en CONSULTAR ────────────
            time.sleep(random.uniform(0.5, 1.0))
            btn = None
            for xpath in [
                "//button[contains(translate(text(),'abcdefghijklmnopqrstuvwxyz','ABCDEFGHIJKLMNOPQRSTUVWXYZ'),'CONSULTAR')]",
                "//button[@type='submit']",
                "//button",
            ]:
                try:
                    btn = wait.until(EC.element_to_be_clickable((By.XPATH, xpath)))
                    break
                except Exception:
                    continue

            if not btn:
                raise RuntimeError("No se encontró el botón CONSULTAR")

            self._driver.execute_script("arguments[0].scrollIntoView({block:'center'});", btn)
            time.sleep(0.3)
            self._driver.execute_script("arguments[0].click();", btn)

            # Esperar dinámicamente a que ONPE cargue los resultados
            # (usar textContent para capturar elementos ocultos/animados de Angular)
            def _results_ready(d):
                txt = (d.execute_script("return document.body.textContent;") or '').lower()
                # SOLO frases únicas de resultados. La página inicial dice
                # "y si eres miembro de mesa" → NO usar 'si eres miembro de mesa'
                return ('no eres miembro de mesa' in txt
                        or 'sí eres miembro de mesa' in txt
                        or 'nombres y apellidos' in txt
                        or 'error interno del servidor' in txt)
            try:
                WebDriverWait(self._driver, 25).until(_results_ready)
                self.log(f"  Resultados detectados para {dni}")
            except Exception:
                self.log(f"  Timeout esperando resultados para {dni}, extrayendo lo que haya")

            # Esperar a que cargue también el local de votación (Angular renderiza en fases)
            try:
                WebDriverWait(self._driver, 8).until(
                    lambda d: 'tu local de votaci' in (
                        d.execute_script("return document.body.textContent;") or '').lower()
                )
            except Exception:
                pass

            # ── Detectar error y reintentar ──
            body_text = self._driver.execute_script("return document.body.textContent;") or ''
            if ('error interno' in body_text.lower()
                    or 'volver al inicio' in body_text.lower()):
                self.log(f"  Error servidor para {dni}, esperando 30s y reintentando...")
                time.sleep(30)
                self._driver.get(ONPE_URL)
                time.sleep(random.uniform(5, 8))
                self._human_behavior()
                time.sleep(random.uniform(2, 3))
                dni_input2 = WebDriverWait(self._driver, 20).until(
                    EC.element_to_be_clickable((By.CSS_SELECTOR, 'input[type="tel"]')))
                self._driver.execute_script(
                    "arguments[0].scrollIntoView({block:'center'}); arguments[0].click();",
                    dni_input2)
                time.sleep(0.3)
                dni_input2.send_keys(Keys.CONTROL + 'a')
                time.sleep(0.1)
                dni_input2.send_keys(Keys.DELETE)
                time.sleep(0.2)
                for ch in dni:
                    dni_input2.send_keys(ch)
                    time.sleep(random.uniform(0.07, 0.15))
                time.sleep(random.uniform(1.0, 2.0))
                btn2 = WebDriverWait(self._driver, 20).until(
                    EC.element_to_be_clickable((By.XPATH, "//button")))
                self._driver.execute_script("arguments[0].scrollIntoView({block:'center'});", btn2)
                time.sleep(0.3)
                self._driver.execute_script("arguments[0].click();", btn2)
                try:
                    WebDriverWait(self._driver, 20).until(_results_ready)
                except Exception:
                    pass

            # ── Extraer datos del DOM ────────
            r = self._extract_from_dom(dni)

        except Exception as e:
            self.log(f"  Error en navegador para {dni}: {e}")
            r['estado']    = 'error'
            r['error_msg'] = str(e)

        return r

    def _extract_from_dom(self, dni):
        """Extrae datos de ONPE usando CSS selectores y regex sobre textContent plano."""
        from selenium.webdriver.common.by import By
        r = self._empty_record(dni)
        try:
            # textContent captura TODO el texto del DOM incluyendo elementos ocultos/animados
            body_text = self._driver.execute_script("return document.body.textContent;") or ''

            if ('error interno' in body_text.lower()
                    or 'volver al inicio' in body_text.lower()):
                r['estado']    = 'error'
                r['error_msg'] = 'Error interno del servidor ONPE'
                return r

            def _css(selector):
                try:
                    el = self._driver.find_element(By.CSS_SELECTOR, selector)
                    return (el.get_attribute('textContent') or el.text or '').strip()
                except Exception:
                    return ''

            body_up = body_text.upper()

            # ── Miembro de mesa ──────────────────────────────────────────────────────
            if 'NO ERES MIEMBRO DE MESA' in body_up:
                r['miembro_mesa'] = False
            elif 'SÍ ERES MIEMBRO DE MESA' in body_up or 'SI ERES MIEMBRO DE MESA' in body_up:
                r['miembro_mesa'] = True

            # ── Nombre (selector CSS primario, regex fallback) ────────────────────────
            nombre = _css('.apellido') or _css('.nombre-completo')
            if not nombre:
                m = re.search(r'Nombres?\s*y\s*Apellidos?\s+(.+?)(?=Regi[oó]n|$)', body_text, re.I)
                if m:
                    nombre = m.group(1).strip()
            if nombre and not nombre.isdigit() and len(nombre) > 3:
                r['nombres'] = nombre

            # ── Región / Provincia / Distrito (selector CSS primario, regex fallback) ─
            geo = _css('.local')
            if not geo or '/' not in geo:
                m = re.search(
                    r'Regi[oó]n\s*/\s*Provincia\s*/\s*Distrito\s*(.+?)(?=Capac[íi]|Tu\s+local|Oficina|$)',
                    body_text, re.I)
                if m:
                    geo = m.group(1).strip()
            if '/' in geo:
                parts = geo.split('/')
                r['region']    = parts[0].strip()
                r['provincia'] = parts[1].strip() if len(parts) > 1 else ''
                r['distrito']  = parts[2].strip() if len(parts) > 2 else ''

            # ── Local de votación, Dirección, Referencia (regex sobre texto plano) ────
            # textContent de Angular es una sola línea: "...ver Mapa<LOCAL><DIR>Referencia:<REF>N°..."
            m_local = re.search(
                r'ver\s*Mapa\s*(.+?)(?=\s*(?:AV\.?\s|JR\.?\s|CALLE\s|PSJ\.?\s|Referencia:|N[°º]\s*de\s*Mesa|Oficina))',
                body_text, re.I)
            if m_local:
                r['local_vot'] = m_local.group(1).strip()

            m_dir = re.search(
                r'((?:AV\.?\s|JR\.?\s|CALLE\s|PSJ\.?\s|PJ\.\s)\S.+?)(?=Referencia:|N[°º]\s*de\s*Mesa|Oficina)',
                body_text, re.I)
            if m_dir:
                r['direccion'] = m_dir.group(1).strip()

            m_ref = re.search(
                r'Referencia:\s*(.+?)(?=N[°º]\s*de\s*Mesa|Oficina|$)',
                body_text, re.I)
            if m_ref:
                r['referencia'] = m_ref.group(1).strip()

            # ── N° Mesa y N° Orden (todos los votantes tienen mesa, no solo miembros) ──
            m_mesa = re.search(r'N[°º]\s*de\s*Mesa:\s*(\d+)', body_text, re.I)
            if m_mesa:
                r['nro_mesa'] = m_mesa.group(1)

            m_orden = re.search(r'N[°º]\s*de\s*Orden:\s*(\d+)', body_text, re.I)
            if m_orden:
                r['nro_orden'] = m_orden.group(1)

            r['estado'] = 'ok'

        except Exception as e:
            r['estado']    = 'error'
            r['error_msg'] = str(e)

        return r

    def _empty_record(self, dni):
        return {
            'dni': dni, 'nombres': '', 'region': '', 'provincia': '',
            'distrito': '', 'miembro_mesa': False, 'local_vot': '',
            'direccion': '', 'referencia': '', 'nro_mesa': '',
            'nro_orden': '', 'estado': 'pendiente', 'error_msg': '',
        }

    # ── Comportamiento humano ─────────────────
    def _human_behavior(self):
        """Simula movimientos reales de mouse via ActionChains para mejorar reCAPTCHA v3."""
        from selenium.webdriver.common.action_chains import ActionChains
        from selenium.webdriver.common.by import By
        try:
            body = self._driver.find_element(By.TAG_NAME, 'body')
            ac = ActionChains(self._driver)
            # Movimientos reales de mouse (no JS sintético — reCAPTCHA los diferencia)
            ac.move_to_element(body)
            ac.perform()
            time.sleep(random.uniform(0.3, 0.5))
            for _ in range(random.randint(4, 7)):
                x_off = random.randint(-200, 200)
                y_off = random.randint(-100, 100)
                try:
                    ActionChains(self._driver).move_to_element_with_offset(
                        body, x_off, y_off).perform()
                except Exception:
                    pass
                time.sleep(random.uniform(0.1, 0.3))
            self._driver.execute_script("window.scrollBy(0, 80)")
            time.sleep(random.uniform(0.3, 0.5))
            self._driver.execute_script("window.scrollBy(0, -80)")
            time.sleep(random.uniform(0.2, 0.4))
        except Exception:
            pass

    # ── Volver al formulario ──────────────────
    def back_to_form(self):
        """Cierra la sesión activa y vuelve al formulario."""
        from selenium.webdriver.common.by import By
        from selenium.webdriver.support.ui import WebDriverWait
        from selenium.webdriver.support import expected_conditions as EC
        try:
            for xpath in ["//a[contains(text(),'Salir') or contains(text(),'SALIR')]",
                          "//button[contains(text(),'Salir') or contains(text(),'SALIR')]"]:
                try:
                    el = WebDriverWait(self._driver, 3).until(
                        EC.element_to_be_clickable((By.XPATH, xpath)))
                    self._driver.execute_script("arguments[0].click();", el)
                    time.sleep(random.uniform(2.0, 3.0))
                    self._human_behavior()
                    return
                except Exception:
                    continue
            self._driver.get(ONPE_URL)
            time.sleep(random.uniform(2.5, 4.0))
        except Exception:
            try:
                self._driver.get(ONPE_URL)
                time.sleep(3)
            except Exception:
                pass
        self._human_behavior()

    def stop(self):
        try:
            if self._driver:
                self._driver.quit()
        except Exception:
            pass


# ─────────────────────────────────────────────
# INTERFAZ GRÁFICA
# ─────────────────────────────────────────────
class App:
    COLS = [
        ('dni',          'DNI',              85),
        ('nombres',      'Nombres y Apellidos', 210),
        ('region',       'Región',           75),
        ('provincia',    'Provincia',        100),
        ('distrito',     'Distrito',         120),
        ('miembro_mesa', 'Miembro Mesa',      95),
        ('local_vot',    'Local de Votación',180),
        ('nro_mesa',     'N° Mesa',           75),
        ('nro_orden',    'N° Orden',          65),
        ('estado',       'Estado',            60),
    ]

    def __init__(self, root):
        self.root      = root
        self.root.title("ONPE – Consulta Electoral Masiva  v2.0")
        self.root.geometry("1200x720")
        self.root.minsize(950, 600)

        self.db            = Database()
        self.dnis_queue    = []
        self.running       = False
        self._force_manual = False
        self._log_q        = queue.Queue()
        self._res_q        = queue.Queue()

        self._style()
        self._build()
        self._poll()           # arrancar el loop de actualización GUI

    # ── Estilos ──────────────────────────────
    def _style(self):
        s = ttk.Style()
        s.theme_use('clam')
        s.configure('H.TLabel', font=('Segoe UI', 13, 'bold'), foreground='#1a237e')
        s.configure('Stat.TLabel', font=('Segoe UI', 9))
        s.configure('OK.TLabel',   font=('Segoe UI', 9, 'bold'), foreground='#2e7d32')
        s.configure('ERR.TLabel',  font=('Segoe UI', 9, 'bold'), foreground='#b71c1c')

    # ── Construcción de la UI ─────────────────
    def _build(self):
        root = self.root
        main = ttk.Frame(root, padding=8)
        main.pack(fill=tk.BOTH, expand=True)

        # ── Header ──────────────────────────
        hdr = ttk.Frame(main)
        hdr.pack(fill=tk.X, pady=(0, 4))

        ttk.Label(hdr, text="ONPE – Consulta Electoral Masiva",
                  style='H.TLabel').pack(side=tk.LEFT)

        self._stat_total    = ttk.Label(hdr, text="Total: 0",          style='Stat.TLabel')
        self._stat_miembros = ttk.Label(hdr, text="Miembros: 0",       style='OK.TLabel')
        self._stat_errores  = ttk.Label(hdr, text="Errores: 0",        style='ERR.TLabel')
        for w in (self._stat_errores, self._stat_miembros, self._stat_total):
            w.pack(side=tk.RIGHT, padx=8)

        ttk.Separator(main).pack(fill=tk.X, pady=4)

        # ── Panel de control ─────────────────
        ctrl = ttk.LabelFrame(main, text="Control", padding=8)
        ctrl.pack(fill=tk.X, pady=4)

        # Fila 1 – Archivo
        r1 = ttk.Frame(ctrl)
        r1.pack(fill=tk.X, pady=2)
        ttk.Label(r1, text="Archivo DNIs (.csv/.txt/.xlsx):").pack(side=tk.LEFT, padx=4)
        self._file_var = tk.StringVar()
        ttk.Entry(r1, textvariable=self._file_var, width=52).pack(side=tk.LEFT, padx=4)
        ttk.Button(r1, text="Examinar…", command=self._pick_file).pack(side=tk.LEFT)
        self._lbl_count = ttk.Label(r1, text="  DNIs en cola: 0")
        self._lbl_count.pack(side=tk.LEFT, padx=12)

        # Fila 2 – DNI manual + opciones
        r2 = ttk.Frame(ctrl)
        r2.pack(fill=tk.X, pady=2)
        ttk.Label(r2, text="DNI manual:").pack(side=tk.LEFT, padx=4)
        self._ent_dni = ttk.Entry(r2, width=12)
        self._ent_dni.pack(side=tk.LEFT, padx=4)
        self._ent_dni.bind('<Return>', lambda e: self._add_manual())
        ttk.Button(r2, text="▶ Consultar", command=self._add_manual).pack(side=tk.LEFT)

        ttk.Separator(r2, orient='vertical').pack(side=tk.LEFT, padx=12, fill=tk.Y)

        self._headless_var = tk.BooleanVar(value=False)
        ttk.Checkbutton(r2, text="Navegador oculto",
                        variable=self._headless_var).pack(side=tk.LEFT, padx=4)

        self._skip_done_var = tk.BooleanVar(value=True)
        ttk.Checkbutton(r2, text="Saltar ya consultados",
                        variable=self._skip_done_var).pack(side=tk.LEFT, padx=4)

        ttk.Label(r2, text="  Pausa (s):").pack(side=tk.LEFT)
        self._delay_var = tk.DoubleVar(value=DELAY_DEFAULT)
        ttk.Spinbox(r2, from_=8, to=60, increment=1,
                    textvariable=self._delay_var, width=5).pack(side=tk.LEFT, padx=4)

        # Fila 3 – Botones de acción
        r3 = ttk.Frame(ctrl)
        r3.pack(fill=tk.X, pady=4)
        self._btn_start  = ttk.Button(r3, text="▶  INICIAR CONSULTAS", command=self._start)
        self._btn_start.pack(side=tk.LEFT, padx=4)
        self._btn_stop   = ttk.Button(r3, text="⏹  DETENER",
                                      command=self._stop, state=tk.DISABLED)
        self._btn_stop.pack(side=tk.LEFT, padx=4)
        ttk.Button(r3, text="📤 Exportar CSV",   command=self._export).pack(side=tk.LEFT, padx=4)
        ttk.Button(r3, text="🗑  Limpiar vista",  command=self._clear_view).pack(side=tk.LEFT, padx=4)

        # ── Barra de progreso ────────────────
        self._prog_var = tk.DoubleVar()
        ttk.Progressbar(main, variable=self._prog_var,
                        maximum=100).pack(fill=tk.X, pady=(4, 0))
        self._prog_lbl = ttk.Label(main, text="Listo.", anchor=tk.W)
        self._prog_lbl.pack(fill=tk.X)

        # ── Paned: tabla + log ───────────────
        paned = ttk.PanedWindow(main, orient=tk.VERTICAL)
        paned.pack(fill=tk.BOTH, expand=True, pady=4)

        # Tabla
        tf = ttk.LabelFrame(paned, text="Resultados", padding=4)
        paned.add(tf, weight=4)

        self.tree = ttk.Treeview(
            tf, columns=[c[0] for c in self.COLS], show='headings', height=14)
        for col, hd, w in self.COLS:
            self.tree.heading(col, text=hd,
                              command=lambda c=col: self._sort(c))
            self.tree.column(col, width=w, minwidth=40)

        vsb = ttk.Scrollbar(tf, orient='vertical',   command=self.tree.yview)
        hsb = ttk.Scrollbar(tf, orient='horizontal', command=self.tree.xview)
        self.tree.configure(yscrollcommand=vsb.set, xscrollcommand=hsb.set)
        vsb.pack(side=tk.RIGHT,  fill=tk.Y)
        hsb.pack(side=tk.BOTTOM, fill=tk.X)
        self.tree.pack(fill=tk.BOTH, expand=True)

        self.tree.tag_configure('miembro',    background='#e8f5e9')
        self.tree.tag_configure('no_miembro', background='#ffebee')
        self.tree.tag_configure('error',      background='#fff3e0')

        # Log
        lf = ttk.LabelFrame(paned, text="Log de actividad", padding=4)
        paned.add(lf, weight=1)

        self._log_txt = tk.Text(lf, height=6, font=('Consolas', 8),
                                state=tk.DISABLED, bg='#1e1e1e', fg='#d4d4d4',
                                wrap=tk.WORD)
        log_sb = ttk.Scrollbar(lf, orient='vertical', command=self._log_txt.yview)
        self._log_txt.configure(yscrollcommand=log_sb.set)
        log_sb.pack(side=tk.RIGHT, fill=tk.Y)
        self._log_txt.pack(fill=tk.BOTH, expand=True)

    # ── Helpers UI ───────────────────────────
    def _log(self, msg):
        ts = datetime.now().strftime('%H:%M:%S')
        self._log_q.put(f"[{ts}] {msg}")

    def _poll(self):
        """Procesa logs y resultados desde los hilos (100 ms)."""
        while not self._log_q.empty():
            msg = self._log_q.get_nowait()
            self._log_txt.config(state=tk.NORMAL)
            self._log_txt.insert(tk.END, msg + '\n')
            self._log_txt.see(tk.END)
            self._log_txt.config(state=tk.DISABLED)

        while not self._res_q.empty():
            rec = self._res_q.get_nowait()
            self._add_row(rec)
            self._refresh_stats()

        self.root.after(100, self._poll)

    def _add_row(self, r):
        miembro_txt = "SÍ ✓" if r.get('miembro_mesa') else "NO"
        tag = ('error' if r.get('estado') == 'error'
               else ('miembro' if r.get('miembro_mesa') else 'no_miembro'))
        self.tree.insert('', 0, values=(
            r.get('dni',''),        r.get('nombres',''),
            r.get('region',''),     r.get('provincia',''),
            r.get('distrito',''),   miembro_txt,
            r.get('local_vot',''),  r.get('nro_mesa',''),
            r.get('nro_orden',''),  r.get('estado',''),
        ), tags=(tag,))

    def _refresh_stats(self):
        s = self.db.stats()
        self._stat_total.config(   text=f"Total: {s.get('total',0)}")
        self._stat_miembros.config(text=f"Miembros: {s.get('miembros',0)}")
        self._stat_errores.config( text=f"Errores: {s.get('errores',0)}")

    def _sort(self, col):
        data = [(self.tree.set(c, col), c) for c in self.tree.get_children('')]
        data.sort(reverse=False)
        for i, (_, iid) in enumerate(data):
            self.tree.move(iid, '', i)

    def _load_existing(self):
        rows = self.db.get_all()
        for r in rows:
            self._add_row(r)
        self._refresh_stats()
        if rows:
            self._log(f"BD existente: {len(rows)} registros cargados.")

    # ── Cargar archivo ────────────────────────
    def _pick_file(self):
        path = filedialog.askopenfilename(
            parent=self.root,
            title="Seleccionar archivo con DNIs",
            initialdir=os.path.dirname(os.path.abspath(__file__)),
            filetypes=[("Todos los archivos", "*.*"),
                       ("Excel",              "*.xlsx"),
                       ("Excel 97-03",        "*.xls"),
                       ("CSV",                "*.csv"),
                       ("TXT",                "*.txt")])
        if path:
            self._file_var.set(path)
            self._load_dnis(path)

    def _load_dnis(self, path):
        dnis = []
        ext = Path(path).suffix.lower()
        try:
            if ext in ('.xlsx', '.xls'):
                import openpyxl
                wb = openpyxl.load_workbook(path, read_only=True, data_only=True)
                for ws in wb.worksheets:
                    for row in ws.iter_rows(values_only=True):
                        for cell in row:
                            if cell is None:
                                continue
                            # Celdas numéricas: convertir a entero para evitar "6793063.0"
                            if isinstance(cell, (int, float)):
                                cell_str = str(int(cell))
                            else:
                                cell_str = str(cell).strip()
                            m = re.search(r'\b(\d{7,9})\b', cell_str)
                            if m:
                                dnis.append(m.group(1).zfill(8))
            else:
                with open(path, 'r', encoding='utf-8-sig', errors='ignore') as f:
                    for line in f:
                        for m in re.finditer(r'\b(\d{7,9})\b', line):
                            dnis.append(m.group(1).zfill(8))

            # Deduplicar (orden preservado)
            seen, unique = set(), []
            for d in dnis:
                if d not in seen:
                    seen.add(d)
                    unique.append(d)

            # Preguntar si limpiar BD y vista para resultados frescos
            existing = self.db.get_all()
            if existing:
                limpiar = messagebox.askyesno(
                    "Base de datos existente",
                    f"Hay {len(existing)} registros previos en la BD.\n"
                    "¿Deseas limpiar la BD antes de iniciar?\n\n"
                    "SÍ = Borrar todo y empezar fresco\n"
                    "NO = Mantener registros y agregar nuevos")
                if limpiar:
                    self.db.clear()
                    for item in self.tree.get_children():
                        self.tree.delete(item)
                    self._refresh_stats()
                    self._log(f"BD limpiada por solicitud del usuario.")

            self.dnis_queue = unique
            self._lbl_count.config(text=f"  DNIs en cola: {len(unique)}")
            self._log(f"Archivo '{Path(path).name}': {len(unique)} DNIs únicos cargados.")
            messagebox.showinfo("Archivo cargado",
                                f"Se encontraron {len(unique)} DNIs únicos.")
        except Exception as e:
            messagebox.showerror("Error al leer archivo", str(e))

    def _add_manual(self):
        dni = self._ent_dni.get().strip()
        if not (7 <= len(dni) <= 8 and dni.isdigit()):
            messagebox.showwarning("DNI inválido", "El DNI debe tener 7 u 8 dígitos numéricos.")
            return
        dni = dni.zfill(8)
        # Reemplazar la cola (no acumular) y forzar re-consulta aunque ya exista en BD
        self.dnis_queue = [dni]
        self._lbl_count.config(text="  DNIs en cola: 1")
        self._log(f"Consultando DNI {dni}...")
        self._ent_dni.delete(0, tk.END)
        self._force_manual = True   # bypass "saltar ya consultados"
        if not self.running:
            self._start()

    # ── Iniciar / Detener ─────────────────────
    def _start(self):
        if not self.dnis_queue:
            messagebox.showwarning("Sin DNIs",
                                   "Carga un archivo o agrega DNIs manualmente.")
            return
        if self.running:
            return

        self.running = True
        self._btn_start.config(state=tk.DISABLED)
        self._btn_stop.config(state=tk.NORMAL)

        # Limpiar vista antes de cada nueva tanda de consultas
        for item in self.tree.get_children():
            self.tree.delete(item)

        t = threading.Thread(target=self._worker, daemon=True)
        t.start()

    def _stop(self):
        self.running = False
        self._log("Deteniendo… Se completará la consulta actual.")
        self._btn_stop.config(state=tk.DISABLED)

    def _worker(self):
        """Hilo principal de consultas."""
        import traceback
        scraper = None
        dnis = list(self.dnis_queue)
        force = getattr(self, '_force_manual', False)
        self._force_manual = False

        # Filtrar ya consultados (excepto cuando se lanzó desde el botón manual)
        if not force and self._skip_done_var.get():
            done = self.db.pending_dnis()
            before = len(dnis)
            dnis = [d for d in dnis if d not in done]
            if before != len(dnis):
                self._log(f"Saltados {before - len(dnis)} DNIs ya consultados.")

        total = len(dnis)
        if total == 0:
            self._log("Todos los DNIs ya están consultados.")
            self.root.after(0, self._done)
            return

        self._log(f"Iniciando {total} consultas...")

        try:
            scraper = ONPEScraper(
                headless=self._headless_var.get(),
                log_fn=self._log,
            )
            scraper.start()

            for i, dni in enumerate(dnis):
                if not self.running:
                    break

                pct = (i + 1) / total * 100
                self.root.after(0, lambda p=pct: self._prog_var.set(p))
                self.root.after(0, lambda i=i, t=total, d=dni:
                    self._prog_lbl.config(
                        text=f"Consultando {i+1}/{t} — DNI: {d}"))

                result = scraper.query_dni(dni)
                self.db.upsert(result)
                self._res_q.put(result)

                estado = result.get('estado', '?')
                miembro = "MIEMBRO" if result.get('miembro_mesa') else "NO miembro"
                nombre  = result.get('nombres', '')
                self._log(f"  {dni} → {estado.upper()} | {miembro} | {nombre}")

                if i < total - 1 and self.running:
                    scraper.back_to_form()
                    time.sleep(self._delay_var.get())

        except Exception as e:
            tb = traceback.format_exc()
            err_msg = str(e)
            self._log(f"ERROR CRÍTICO: {err_msg}\n{tb}")
            self.root.after(0, lambda m=err_msg: messagebox.showerror(
                "Error", f"El proceso falló:\n\n{m}"))
        finally:
            if scraper:
                scraper.stop()
            self.root.after(0, self._done)

    def _done(self):
        self.running = False
        self._btn_start.config(state=tk.NORMAL)
        self._btn_stop.config(state=tk.DISABLED)
        self._prog_var.set(100)
        self._prog_lbl.config(text="Proceso completado.")
        self._log("=== Proceso finalizado. ===")
        self._refresh_stats()
        messagebox.showinfo("Listo", "Todas las consultas han sido procesadas.")

    # ── Exportar / Limpiar ────────────────────
    def _export(self):
        path = filedialog.asksaveasfilename(
            defaultextension='.csv',
            filetypes=[('CSV', '*.csv'), ('Todos', '*.*')],
            initialfile=f"onpe_{datetime.now().strftime('%Y%m%d_%H%M%S')}.csv")
        if not path:
            return
        if self.db.export_csv(path):
            self._log(f"Exportado: {path}")
            messagebox.showinfo("Exportado", f"Archivo guardado en:\n{path}")
        else:
            messagebox.showwarning("Sin datos", "No hay registros para exportar.")

    def _clear_view(self):
        for item in self.tree.get_children():
            self.tree.delete(item)
        self._log("Vista limpiada (la base de datos no fue modificada).")


# ─────────────────────────────────────────────
# PUNTO DE ENTRADA
# ─────────────────────────────────────────────
def main():
    root = tk.Tk()
    App(root)
    root.mainloop()


if __name__ == '__main__':
    main()
