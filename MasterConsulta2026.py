import sqlite3
import pandas as pd
import hashlib
import json
from datetime import datetime
import streamlit as st

# ==============================================================================
# 1. BASE DE DATOS Y CARGA DEL PADRÓN ELECTORAL (CON CASTING Y CONEXIÓN MULTIHILO)
# ==============================================================================
class ElectoralDatabase:
    def __init__(self, db_name="ahora_nacion_electoral.db"):
        # check_same_thread=False previene errores de concurrencia en Streamlit Cloud
        self.conn = sqlite3.connect(db_name, check_same_thread=False)
        self.create_tables()

    def create_tables(self):
        cursor = self.conn.cursor()
        
        # Tabla: Locales de Votación (Estructura Correlacionada con el Excel Maestro)
        cursor.execute('''
        CREATE TABLE IF NOT EXISTS locales_votacion (
            codigo_lv INTEGER PRIMARY KEY,
            odpe TEXT,
            ubigeo INTEGER,
            departamento TEXT,
            provincia TEXT,
            distrito TEXT,
            nombre_local TEXT,
            direccion_local TEXT
        )
        ''')

        # Tabla: Mesas de Sufragio
        cursor.execute('''
        CREATE TABLE IF NOT EXISTS mesas_sufragio (
            mesa_sufragio INTEGER PRIMARY KEY,
            codigo_lv INTEGER,
            orden_ca INTEGER,
            numero_ca INTEGER,
            orden_mesas INTEGER,
            aula INTEGER,
            piso INTEGER,
            pabellon INTEGER,
            electores INTEGER,
            electores_discapacidad INTEGER,
            FOREIGN KEY (codigo_lv) REFERENCES locales_votacion(codigo_lv)
        )
        ''')

        # Tabla: Usuarios y Roles Jerárquicos
        cursor.execute('''
        CREATE TABLE IF NOT EXISTS usuarios (
            dni TEXT PRIMARY KEY,
            nombre TEXT,
            telefono TEXT,
            rol TEXT,
            password_hash TEXT,
            estado TEXT DEFAULT 'ACTIVO'
        )
        ''')

        # Tabla: Asignaciones de Personeros y Check-in GPS
        cursor.execute('''
        CREATE TABLE IF NOT EXISTS asignaciones_personero (
            id_asignacion INTEGER PRIMARY KEY AUTOINCREMENT,
            dni_personero TEXT,
            mesa_sufragio INTEGER,
            checkin_status INTEGER DEFAULT 0,
            checkin_timestamp TEXT,
            latitud REAL,
            longitud REAL,
            FOREIGN KEY (dni_personero) REFERENCES usuarios(dni),
            FOREIGN KEY (mesa_sufragio) REFERENCES mesas_sufragio(mesa_sufragio)
        )
        ''')

        # Tabla: Actas de Segunda Vuelta
        cursor.execute('''
        CREATE TABLE IF NOT EXISTS actas_segunda_vuelta (
            mesa_sufragio INTEGER PRIMARY KEY,
            dni_personero TEXT,
            votos_ahora_nacion INTEGER,
            votos_contendiente INTEGER,
            votos_blancos INTEGER,
            votos_nulos INTEGER,
            votos_impugnados INTEGER,
            total_emitidos INTEGER,
            foto_acta_url TEXT,
            estado_acta TEXT,
            fecha_registro TEXT,
            FOREIGN KEY (mesa_sufragio) REFERENCES mesas_sufragio(mesa_sufragio)
        )
        ''')

        # Tabla: Incidencias Legales y Pánico
        cursor.execute('''
        CREATE TABLE IF NOT EXISTS incidencias_legales (
            id_incidencia INTEGER PRIMARY KEY AUTOINCREMENT,
            mesa_sufragio INTEGER,
            dni_reporta TEXT,
            tipo_incidencia TEXT,
            descripcion TEXT,
            prioridad TEXT,
            fecha_hora TEXT
        )
        ''')

        self.conn.commit()

    def cargar_excel_maestro(self, excel_path="Maestro_EMR2026.xlsx"):
        """Carga y convierte explícitamente los datos de pandas/numpy a tipos nativos de Python"""
        df = pd.read_excel(excel_path, sheet_name='TacnaEMR2026')
        cursor = self.conn.cursor()

        # 1. Cargar Locales de Votación
        locales = df[['CÓDIGO DE LV', 'ODPE', 'UBIGEO', 'DEPARTAMENTO', 'PROVINCIA', 
                      'DISTRITO', 'NOMBRE DE LOCAL DE VOTACIÓN', 'DIRECCION DEL LOCAL DE VOTACIÓN']].drop_duplicates()
        
        for _, r in locales.iterrows():
            cursor.execute('''
            INSERT OR REPLACE INTO locales_votacion VALUES (?, ?, ?, ?, ?, ?, ?, ?)
            ''', (
                int(r['CÓDIGO DE LV']), 
                str(r['ODPE']), 
                int(r['UBIGEO']), 
                str(r['DEPARTAMENTO']), 
                str(r['PROVINCIA']), 
                str(r['DISTRITO']), 
                str(r['NOMBRE DE LOCAL DE VOTACIÓN']), 
                str(r['DIRECCION DEL LOCAL DE VOTACIÓN'])
            ))

        # 2. Cargar Mesas de Sufragio (Casting nativo `int()` para prevenir IntegrityError)
        mesas = df[['MESA DE SUFRAGIO', 'CÓDIGO DE LV', 'ORDEN DE CA', 'NÚMERO DE CA', 'ORDEN DE MESAS', 
                    'AULA', 'PISO', 'PABELLÓN', 'ELECTORES', 'ELECTORES CON DISCAPACIDAD']]
        
        for _, r in mesas.iterrows():
            cursor.execute('''
            INSERT OR REPLACE INTO mesas_sufragio VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?)
            ''', (
                int(r['MESA DE SUFRAGIO']), 
                int(r['CÓDIGO DE LV']), 
                int(r['ORDEN DE CA']),
                int(r['NÚMERO DE CA']),
                int(r['ORDEN DE MESAS']),
                int(r['AULA']), 
                int(r['PISO']), 
                int(r['PABELLÓN']), 
                int(r['ELECTORES']), 
                int(r['ELECTORES CON DISCAPACIDAD'])
            ))

        self.conn.commit()


# ==============================================================================
# 2. MOTOR DE LÓGICA DE NEGOCIO Y SERVICIOS
# ==============================================================================
class ServicioElectoral:
    def __init__(self, db: ElectoralDatabase):
        self.db = db

    def hash_password(self, password: str) -> str:
        return hashlib.sha256(str(password).encode()).hexdigest()

    def registrar_usuario(self, dni, nombre, telefono, rol, password):
        cursor = self.db.conn.cursor()
        pwd_hash = self.hash_password(password)
        cursor.execute("INSERT OR REPLACE INTO usuarios VALUES (?, ?, ?, ?, ?, 'ACTIVO')",
                       (str(dni), str(nombre), str(telefono), str(rol), pwd_hash))
        self.db.conn.commit()
        return {"status": "SUCCESS", "message": f"Usuario {nombre} registrado como {rol}."}

    def autenticar_usuario(self, dni, password):
        cursor = self.db.conn.cursor()
        pwd_hash = self.hash_password(password)
        cursor.execute("SELECT dni, nombre, rol FROM usuarios WHERE dni = ? AND password_hash = ?", (str(dni), pwd_hash))
        user = cursor.fetchone()
        if user:
            return {"autenticado": True, "user": {"dni": user[0], "nombre": user[1], "rol": user[2]}}
        return {"autenticado": False, "message": "Credenciales inválidas."}

    def consultar_mesa(self, mesa_sufragio):
        """Módulo: Consulta de Mesas y Ubicación Electoral"""
        try:
            mesa_id = int(mesa_sufragio)
        except (ValueError, TypeError):
            return None

        cursor = self.db.conn.cursor()
        query = '''
        SELECT m.mesa_sufragio, m.aula, m.piso, m.pabellon, m.electores, m.electores_discapacidad,
               m.orden_ca, m.numero_ca, l.nombre_local, l.direccion_local, l.distrito, l.provincia, l.departamento
        FROM mesas_sufragio m
        JOIN locales_votacion l ON m.codigo_lv = l.codigo_lv
        WHERE m.mesa_sufragio = ?
        '''
        cursor.execute(query, (mesa_id,))
        res = cursor.fetchone()
        if res:
            return {
                "mesa": res[0], "aula": res[1], "piso": res[2], "pabellon": res[3],
                "electores": res[4], "discapacidad": res[5],
                "orden_ca": res[6], "numero_ca": res[7],
                "local": res[8], "direccion": res[9], "distrito": res[10],
                "provincia": res[11], "departamento": res[12]
            }
        return None

    def asignar_personero_mesa(self, dni_personero, mesa_sufragio):
        """Módulo: Asignación de Personeros"""
        try:
            mesa_id = int(mesa_sufragio)
        except (ValueError, TypeError):
            return {"status": "ERROR", "message": "Mesa inválida."}

        cursor = self.db.conn.cursor()
        cursor.execute("INSERT INTO asignaciones_personero (dni_personero, mesa_sufragio) VALUES (?, ?)",
                       (str(dni_personero), mesa_id))
        self.db.conn.commit()
        return {"status": "SUCCESS", "message": f"Personero DNI {dni_personero} asignado a Mesa {mesa_id}."}

    def checkin_personero(self, dni_personero, latitud, longitud):
        """Módulo: Check-in de Asistencia vía GPS"""
        cursor = self.db.conn.cursor()
        now = datetime.now().isoformat()
        cursor.execute('''
        UPDATE asignaciones_personero 
        SET checkin_status = 1, checkin_timestamp = ?, latitud = ?, longitud = ?
        WHERE dni_personero = ?
        ''', (now, float(latitud), float(longitud), str(dni_personero)))
        self.db.conn.commit()
        return {"status": "CHECKIN_OK", "timestamp": now}

    def registrar_acta_segunda_vuelta(self, dni_personero, mesa_sufragio, v_ahora_nacion, v_contendiente, v_blancos, v_nulos, v_impugnados, foto_url):
        """Módulo: Digitalización, Conteo Rápido y Validación (Previene InterfaceError)"""
        try:
            mesa_id = int(mesa_sufragio)
            v_an = int(v_ahora_nacion)
            v_co = int(v_contendiente)
            v_bl = int(v_blancos)
            v_nu = int(v_nulos)
            v_im = int(v_impugnados)
        except (ValueError, TypeError):
            return {"status": "ERROR", "message": "Parámetros numéricos inválidos o no nulos."}

        cursor = self.db.conn.cursor()
        cursor.execute("SELECT electores FROM mesas_sufragio WHERE mesa_sufragio = ?", (mesa_id,))
        mesa_info = cursor.fetchone()
        
        if not mesa_info:
            return {"status": "ERROR", "message": f"La mesa {mesa_id} no existe en el padrón."}
        
        max_electores = int(mesa_info[0])
        total_calculado = v_an + v_co + v_bl + v_nu + v_im

        # Regla de Validación Algorítmica
        estado = "VALIDADA"
        if total_calculado > max_electores:
            estado = "OBSERVADA_EXCESO_VOTOS"
        elif total_calculado == 0:
            estado = "OBSERVADA_SIN_VOTOS"

        now = datetime.now().isoformat()
        cursor.execute('''
        INSERT OR REPLACE INTO actas_segunda_vuelta VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?)
        ''', (mesa_id, str(dni_personero), v_an, v_co, v_bl, v_nu, v_im, total_calculado, str(foto_url), estado, now))
        
        self.db.conn.commit()
        return {
            "status": "REGISTRADO",
            "estado_acta": estado,
            "total_emitidos": total_calculado,
            "max_electores_mesa": max_electores
        }

    def registrar_incidencia(self, mesa_sufragio, dni_reporta, tipo, descripcion, prioridad="ALTA"):
        """Módulo: Registro de Incidencias Legales y Pánico"""
        try:
            mesa_id = int(mesa_sufragio)
        except (ValueError, TypeError):
            return {"status": "ERROR", "message": "Mesa inválida."}

        cursor = self.db.conn.cursor()
        now = datetime.now().isoformat()
        cursor.execute('''
        INSERT INTO incidencias_legales (mesa_sufragio, dni_reporta, tipo_incidencia, descripcion, prioridad, fecha_hora)
        VALUES (?, ?, ?, ?, ?, ?)
        ''', (mesa_id, str(dni_reporta), str(tipo), str(descripcion), str(prioridad), now))
        self.db.conn.commit()
        return {"status": "INCIDENCIA_REGISTRADA", "id": cursor.lastrowid}

    def obtener_resumen_computo_regional(self):
        """Módulo: Cómputo Regional en Tiempo Real"""
        cursor = self.db.conn.cursor()
        cursor.execute('''
        SELECT 
            COUNT(mesa_sufragio) as actas_procesadas,
            SUM(votos_ahora_nacion) as total_ahora_nacion,
            SUM(votos_contendiente) as total_contendiente,
            SUM(votos_blancos) as total_blancos,
            SUM(votos_nulos) as total_nulos,
            SUM(total_emitidos) as total_votantes
        FROM actas_segunda_vuelta
        WHERE estado_acta = 'VALIDADA'
        ''')
        res = cursor.fetchone()
        
        cursor.execute("SELECT COUNT(*), SUM(electores) FROM mesas_sufragio")
        totales = cursor.fetchone()
        total_mesas = totales[0] if totales else 0
        total_electores_region = totales[1] if totales else 0

        actas_proc = res[0] if res and res[0] else 0
        return {
            "total_mesas_region": total_mesas,
            "total_electores_region": total_electores_region,
            "actas_procesadas": actas_proc,
            "porcentaje_escrutado": round((actas_proc / total_mesas) * 100, 2) if total_mesas > 0 else 0,
            "resultados": {
                "AHORA_NACION": res[1] if res and res[1] else 0,
                "CONTENDIENTE": res[2] if res and res[2] else 0,
                "BLANCOS": res[3] if res and res[3] else 0,
                "NULOS": res[4] if res and res[4] else 0,
                "TOTAL_EMITIDOS": res[5] if res and res[5] else 0
            }
        }


# ==============================================================================
# 3. GESTIÓN Y CACHÉ DE RECURSOS STREAMLIT
# ==============================================================================
@st.cache_resource
def inicializar_sistema():
    """Garantiza la instanciación única de la BD y la carga protegida del Excel Maestro"""
    db = ElectoralDatabase()
    cursor = db.conn.cursor()
    cursor.execute("SELECT COUNT(*) FROM mesas_sufragio")
    if cursor.fetchone()[0] == 0:
        db.cargar_excel_maestro("Maestro_EMR2026.xlsx")
    return db


# Instanciación segura de la app
db = inicializar_sistema()
app = ServicioElectoral(db)


# ==============================================================================
# 4. INTERFAZ WEB STREAMLIT
# ==============================================================================
st.set_page_config(page_title="Ahora Nación - Segunda Vuelta", page_icon="🗳️", layout="wide")

st.title("🗳️ Partido Político Ahora Nación - Segunda Vuelta")
st.caption("Sistema Electoral Regional de Cómputo, Personeros y Control de Mesas")

# Menú de Navegación Lateral
opcion = st.sidebar.selectbox(
    "Seleccione Módulo",
    ["Consulta de Mesas", "Registro de Actas", "Incidencias Legales", "Cómputo Regional"]
)

if opcion == "Consulta de Mesas":
    st.subheader("🔍 Consulta de Local y Mesa de Votación")
    mesa_input = st.number_input("Número de Mesa de Sufragio:", min_value=1, value=78981, step=1)
    if st.button("Buscar Mesa"):
        info = app.consultar_mesa(int(mesa_input))
        if info:
            st.success(f"Local: {info['local']}")
            col1, col2 = st.columns(2)
            with col1:
                st.write(f"**Provincia:** {info['provincia']}")
                st.write(f"**Distrito:** {info['distrito']}")
                st.write(f"**Dirección:** {info['direccion']}")
            with col2:
                st.write(f"**Aula:** {info['aula']} | **Piso:** {info['piso']} | **Pabellón:** {info['pabellon']}")
                st.write(f"**Electores Hábiles:** {info['electores']}")
        else:
            st.error("Mesa no encontrada en el padrón maestro.")

elif opcion == "Registro de Actas":
    st.subheader("📝 Conteo Rápido y Digitalización de Actas")
    col_a, col_b = st.columns(2)
    with col_a:
        dni_p = st.text_input("DNI del Personero:", "45879632")
        mesa_p = st.number_input("Mesa de Sufragio:", min_value=1, value=78981)
        v_an = st.number_input("Votos AHORA NACIÓN:", min_value=0, value=0)
        v_co = st.number_input("Votos Contendiente:", min_value=0, value=0)
    with col_b:
        v_bl = st.number_input("Votos Blancos:", min_value=0, value=0)
        v_nu = st.number_input("Votos Nulos:", min_value=0, value=0)
        v_im = st.number_input("Votos Impugnados:", min_value=0, value=0)
        foto_url = st.text_input("URL / Foto del Acta:", "https://storage.ahoranacion.pe/actas/78981.jpg")

    if st.button("Enviar y Validar Acta"):
        res = app.registrar_acta_segunda_vuelta(dni_p, mesa_p, v_an, v_co, v_bl, v_nu, v_im, foto_url)
        if res["status"] == "REGISTRADO":
            if res["estado_acta"] == "VALIDADA":
                st.success(f"Acta Registrada y Validada. Total Votos: {res['total_emitidos']}")
            else:
                st.warning(f"Acta Registrada con Estado: {res['estado_acta']}. Total Votos ({res['total_emitidos']}) supera el límite de la mesa ({res['max_electores_mesa']}).")
        else:
            st.error(res["message"])

elif opcion == "Incidencias Legales":
    st.subheader("⚠️ Botón de Pánico e Incidencias")
    mesa_i = st.number_input("Mesa de Sufragio:", min_value=1, value=78981)
    dni_i = st.text_input("DNI que Reporta:", "45879632")
    tipo_i = st.selectbox("Tipo de Incidencia", ["PROPAGANDA_PROHIBIDA", "AUSENCIA_PERSONERO", "SUPLANTACION", "OTRO"])
    desc_i = st.text_area("Descripción de la Incidencia")
    
    if st.button("Reportar Incidencia"):
        res_i = app.registrar_incidencia(mesa_i, dni_i, tipo_i, desc_i)
        st.error(f"Alerta registrada con ID #{res_i['id']}. Equipo legal notificado.")

elif opcion == "Cómputo Regional":
    st.subheader("📊 Tablero de Cómputo Regional en Tiempo Real")
    resumen = app.obtener_resumen_computo_regional()
    
    m1, m2, m3 = st.columns(3)
    m1.metric("Actas Procesadas", f"{resumen['actas_procesadas']} / {resumen['total_mesas_region']}")
    m2.metric("Porcentaje Escrutado", f"{resumen['porcentaje_escrutado']}%")
    m3.metric("Electores Totales", f"{resumen['total_electores_region']}")

    st.write("### Resultados Proyectados")
    st.json(resumen["resultados"])