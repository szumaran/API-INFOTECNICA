import asyncio
import aiohttp
import logging
import os
import sys
import locale
import re
from dataclasses import dataclass
from itertools import chain
from typing import Any, Dict, List, Optional
import win32com.client as win32

# Configuración de logging básico
logging.basicConfig(level=logging.INFO, format='%(asctime)s - %(levelname)s - %(message)s')
locale.setlocale(locale.LC_NUMERIC, '')

BASE_URLS = {
    'subestaciones': 'https://api-infotecnica.coordinador.cl/v1/subestaciones',
    'interruptores': 'https://api-infotecnica.coordinador.cl/v1/interruptores',
    'transformadores_2d': 'https://api-infotecnica.coordinador.cl/v1/transformadores-2d',
    'transformadores_3d': 'https://api-infotecnica.coordinador.cl/v1/transformadores-3d',
    'lineas': 'https://api-infotecnica.coordinador.cl/v1/lineas',
    'tramos': 'https://api-infotecnica.coordinador.cl/v1/tramos',
    'secciones_tramos': 'https://api-infotecnica.coordinador.cl/v1/secciones-tramos',
    'desconectadores': 'https://api-infotecnica.coordinador.cl/v1/desconectadores',
    'transformadores_corriente': 'https://api-infotecnica.coordinador.cl/v1/transformadores-corrientes',
    'trampas_ondas': 'https://api-infotecnica.coordinador.cl/v1/trampas-ondas',
    'unidades_generadoras': 'https://api-infotecnica.coordinador.cl/v1/unidades-generadoras',
    'panos': 'https://api-infotecnica.coordinador.cl/v1/panos',
}

HEADERS = {
    "User-Agent": "Mozilla/5.0",
    "Accept": "application/json, text/plain, */*",
    "Referer": "https://infotecnica.coordinador.cl/",
    "Origin": "https://infotecnica.coordinador.cl",
    "Connection": "keep-alive",
}

@dataclass
class EquipoDatosTecnicos:
    estados_certificacion: Dict[str, str]

@dataclass
class InterruptorDatosTecnicos(EquipoDatosTecnicos):
    tension_nominal: Optional[str] = None
    corriente_nominal: Optional[str] = None
    capacidad_simetrica: Optional[str] = None
    fabricante: Optional[str] = None
    capacidad_asimetrica: Optional[str] = None
    capacidad_cierre: Optional[str] = None
    modelo: Optional[str] = None

@dataclass
class TransformadorDatosTecnicos(EquipoDatosTecnicos):
    tension_at: Optional[str] = None
    tension_mt: Optional[str] = None
    capacidad_at: Optional[str] = None
    capacidad_at_onaf1: Optional[str] = None
    capacidad_at_onaf2: Optional[str] = None
    imp_sec_pos: Optional[str] = None
    pot_sec_pos: Optional[str] = None
    imp_sec_cero: Optional[str] = None
    pot_sec_cero: Optional[str] = None

@dataclass
class Transformador3DDatosTecnicos(EquipoDatosTecnicos):
    tension_at: Optional[str] = None
    tension_mt: Optional[str] = None
    tension_bt: Optional[str] = None
    capacidad_at: Optional[str] = None
    capacidad_mt: Optional[str] = None
    capacidad_bt: Optional[str] = None
    capacidad_at_onaf1: Optional[str] = None
    capacidad_mt_onaf1: Optional[str] = None
    capacidad_bt_onaf1: Optional[str] = None
    capacidad_at_onaf2: Optional[str] = None
    capacidad_mt_onaf2: Optional[str] = None
    capacidad_bt_onaf2: Optional[str] = None
    imp_sec_pos_AT_MT: Optional[str] = None
    pot_sec_pos_AT_MT: Optional[str] = None
    imp_sec_pos_MT_BT: Optional[str] = None
    pot_sec_pos_MT_BT: Optional[str] = None
    imp_sec_pos_BT_AT: Optional[str] = None
    pot_sec_pos_BT_AT: Optional[str] = None
    imp_sec_cero_AT_MT: Optional[str] = None
    pot_sec_cero_AT_MT: Optional[str] = None
    imp_sec_cero_MT_BT: Optional[str] = None
    pot_sec_cero_MT_BT: Optional[str] = None
    imp_sec_cero_BT_AT: Optional[str] = None
    pot_sec_cero_BT_AT: Optional[str] = None

@dataclass
class SeccionTramoDatosTecnicos(EquipoDatosTecnicos):
    tension_nominal: Optional[str] = None
    longitud_conductor: Optional[str] = None
    resistencia_sec_pos: Optional[str] = None
    reactancia_sec_pos: Optional[str] = None
    susceptancia_sec_pos: Optional[str] = None
    resistencia_sec_cero: Optional[str] = None
    reactancia_sec_cero: Optional[str] = None
    susceptancia_sec_cero: Optional[str] = None
    limites_termicos: Dict[str, Any] = None

@dataclass
class DesconectadorDatosTecnicos(EquipoDatosTecnicos):
    nivel_tension: Optional[str] = None
    corriente_nominal: Optional[str] = None
    tipo_desconectador: Optional[str] = None

@dataclass
class TransformadorCorrienteDatosTecnicos(EquipoDatosTecnicos):
    relacion_transformacion: Optional[str] = None
    precision: Optional[str] = None
    corriente_primaria: Optional[str] = None
    elemento_asociado: Optional[str] = None
    proteccion_asociada: Optional[str] = None

@dataclass
class TrampasOndasDatosTecnicos(EquipoDatosTecnicos):
    corriente_nominal: Optional[str] = None
    linea_asociada: Optional[str] = None

@dataclass
class UnidadesDatosTecnicos(EquipoDatosTecnicos):
    tension_nominal: Optional[str] = None
    tecnologia: Optional[str] = None
    potencia_neta: Optional[str] = None
    minimo_tecnico: Optional[str] = None
    subestacion_nombre: Optional[str] = None

@dataclass
class Equipo:
    tipo: str
    id_equipo: int
    nombre_equipo: str
    subestacion_nombre: Optional[str]
    propietario_nombre: Optional[str]
    pano_coordinado_nombre: Optional[str]
    datos_tecnicos: EquipoDatosTecnicos

@dataclass
class Interruptor(Equipo):
    datos_tecnicos: InterruptorDatosTecnicos

@dataclass
class Transformador(Equipo):
    datos_tecnicos: TransformadorDatosTecnicos

@dataclass
class Transformador3D(Equipo):
    datos_tecnicos: Transformador3DDatosTecnicos

@dataclass
class SeccionTramo(Equipo):
    id_linea: int
    nombre_linea: str
    id_tramo: int
    nombre_tramo: str
    extremo1: str
    extremo2: str
    datos_tecnicos: SeccionTramoDatosTecnicos

@dataclass
class Desconectador(Equipo):
    datos_tecnicos: DesconectadorDatosTecnicos

@dataclass
class TransformadorCorriente(Equipo):
    datos_tecnicos: TransformadorCorrienteDatosTecnicos

@dataclass
class TrampaOnda(Equipo):
    datos_tecnicos: TrampasOndasDatosTecnicos

@dataclass
class Unidad(Equipo):
    central_nombre: Optional[str]
    datos_tecnicos: UnidadesDatosTecnicos

class ApiClientFactory:
    @staticmethod
    def create_client(session: aiohttp.ClientSession) -> 'ApiClient':
        return ApiClient(session)

class ApiClient:
    def __init__(self, session: aiohttp.ClientSession):
        self.session = session
        self.nombre_subestacion_cache = {}
        self.datos_tecnicos_cache = {}

    async def hacer_solicitud(self, url: str) -> Optional[Dict[str, Any]]:
        try:
            async with self.session.get(url, headers=HEADERS) as response:
                response.raise_for_status()
                return await response.json()
        except Exception as e:
            logging.error(f"Error al realizar solicitud a {url}: {e}")
            return None

    async def obtener_nombre_subestacion(self, id_subestacion: int) -> Optional[str]:
        if id_subestacion in self.nombre_subestacion_cache:
            return self.nombre_subestacion_cache[id_subestacion]
        url = f"{BASE_URLS['subestaciones']}/{id_subestacion}"
        datos = await self.hacer_solicitud(url)
        nombre = datos.get('nombre') if datos else None
        self.nombre_subestacion_cache[id_subestacion] = nombre
        return nombre

    async def buscar_equipos_por_subestacion_nombre(self, nombre_subestacion: str, tipo_equipo: str) -> List[Dict[str, Any]]:
        if nombre_subestacion.startswith("S/E "):
            nombre_subestacion = nombre_subestacion[4:]
        url = f"{BASE_URLS[tipo_equipo]}?nombre__icontains={nombre_subestacion}"
        datos = await self.hacer_solicitud(url)
        return datos if datos else []
    
    async def buscar_nemotecnico_pano_por_extremo_tramo(self, nombre_extremo: str) -> Optional[str]:
        if nombre_extremo.startswith("S/E "):
            nombre_extremo = nombre_extremo[4:]
        url = f"{BASE_URLS['panos']}?nombre__icontains={nombre_extremo}"
        datos = await self.hacer_solicitud(url)
        if datos:
            for pano in datos:
                if pano.get('nombre', '').lower() == nombre_extremo.lower():
                    return pano.get('nemotecnico', '')
            return datos[0].get('nemotecnico', '')
        else:
            return None
        
    async def buscar_equipos_por_nemotecnico(self, nemotecnico_pano: str, tipo_equipo: str) -> List[Dict[str, Any]]:
        url = f"{BASE_URLS[tipo_equipo]}/?search={nemotecnico_pano}"
        datos = await self.hacer_solicitud(url)
        return datos if datos else []

    async def obtener_datos_tecnicos(self, id_equipo: int, tipo_equipo: str) -> Optional[EquipoDatosTecnicos]:
        cache_key = (id_equipo, tipo_equipo)
        if cache_key in self.datos_tecnicos_cache:
            return self.datos_tecnicos_cache[cache_key]

        if tipo_equipo == 'secciones_tramos':
            endpoint = f"{BASE_URLS['secciones_tramos']}/{id_equipo}/fichas-tecnicas/general/"
            datos = await self.hacer_solicitud(endpoint)
            limites_termicos_endpoint = f"{BASE_URLS['secciones_tramos']}/{id_equipo}/fichas-tecnicas/limites-termicos/"
            limites_termicos = await self.hacer_solicitud(limites_termicos_endpoint)
            result = None
            if datos:
                estados_certificacion = {key: value.get('estado_certificacion_nombre', '') for key, value in datos.items()}
                result = SeccionTramoDatosTecnicos(
                    tension_nominal=datos.get('5895', {}).get('valor_texto', '').replace(",", "."),
                    longitud_conductor=datos.get('1005', {}).get('valor_texto', '').replace(",", "."),
                    resistencia_sec_pos=datos.get('1076', {}).get('valor_texto', '').replace(",", "."),
                    reactancia_sec_pos=datos.get('1009', {}).get('valor_texto', '').replace(",", "."),
                    susceptancia_sec_pos=datos.get('1010', {}).get('valor_texto', '').replace(",", "."),
                    resistencia_sec_cero=datos.get('1090', {}).get('valor_texto', '').replace(",", "."),
                    reactancia_sec_cero=datos.get('1013', {}).get('valor_texto', '').replace(",", "."),
                    susceptancia_sec_cero=datos.get('1014', {}).get('valor_texto', '').replace(",", "."),
                    limites_termicos=limites_termicos,
                    estados_certificacion=estados_certificacion
                )
        else:
            endpoint = f"{BASE_URLS[tipo_equipo]}/{id_equipo}/fichas-tecnicas/general/"
            datos = await self.hacer_solicitud(endpoint)
            result = None
            if datos:
                estados_certificacion = {key: value.get('estado_certificacion_nombre', '') for key, value in datos.items()}
                if tipo_equipo == 'interruptores':
                    result = InterruptorDatosTecnicos(
                        tension_nominal=datos.get('6018', {}).get('valor_texto', '').replace(",", "."),
                        corriente_nominal=datos.get('6019', {}).get('valor_texto', '').replace(",", "."),
                        capacidad_simetrica=datos.get('326', {}).get('valor_texto', '').replace(",", "."),
                        fabricante=datos.get('6023', {}).get('valor_texto', ''),
                        capacidad_asimetrica=datos.get('327', {}).get('valor_texto', '').replace(",", "."),
                        capacidad_cierre=datos.get('328', {}).get('valor_texto', '').replace(",", "."),
                        modelo=datos.get('6022', {}).get('valor_texto', ''),
                        estados_certificacion=estados_certificacion
                    )
                elif tipo_equipo == 'transformadores_2d':
                    result = TransformadorDatosTecnicos(
                        tension_at=datos.get('132', {}).get('valor_texto', '').replace(",", "."),
                        tension_mt=datos.get('133', {}).get('valor_texto', '').replace(",", "."),
                        capacidad_at=datos.get('129', {}).get('valor_texto', '').replace(",", "."),
                        capacidad_at_onaf1=datos.get('130', {}).get('valor_texto', '').replace(",", "."),
                        capacidad_at_onaf2=datos.get('1917', {}).get('valor_texto', '').replace(",", "."),
                        imp_sec_pos=datos.get('136', {}).get('valor_texto', '').replace(",", "."),
                        pot_sec_pos=datos.get('137', {}).get('valor_texto', '').replace(",", "."),
                        imp_sec_cero=datos.get('6503', {}).get('valor_texto', '').replace(",", "."),
                        pot_sec_cero=datos.get('6504', {}).get('valor_texto', '').replace(",", "."),
                        estados_certificacion=estados_certificacion
                    )
                elif tipo_equipo == 'transformadores_3d':
                    result = Transformador3DDatosTecnicos(
                        tension_at=datos.get('5978', {}).get('valor_texto', '').replace(",", "."),
                        tension_mt=datos.get('33', {}).get('valor_texto', '').replace(",", "."),
                        tension_bt=datos.get('5979', {}).get('valor_texto', '').replace(",", "."),
                        capacidad_at=datos.get('5952', {}).get('valor_texto', '').replace(",", "."),
                        capacidad_mt=datos.get('24', {}).get('valor_texto', '').replace(",", "."),
                        capacidad_bt=datos.get('5953', {}).get('valor_texto', '').replace(",", "."),
                        capacidad_at_onaf1=datos.get('5954', {}).get('valor_texto', '').replace(",", "."),
                        capacidad_mt_onaf1=datos.get('5955', {}).get('valor_texto', '').replace(",", "."),
                        capacidad_bt_onaf1=datos.get('5956', {}).get('valor_texto', '').replace(",", "."),
                        capacidad_at_onaf2=datos.get('5957', {}).get('valor_texto', '').replace(",", "."),
                        capacidad_mt_onaf2=datos.get('5971', {}).get('valor_texto', '').replace(",", "."),
                        capacidad_bt_onaf2=datos.get('5958', {}).get('valor_texto', '').replace(",", "."),
                        imp_sec_pos_AT_MT=datos.get('37', {}).get('valor_texto', '').replace(",", "."),
                        pot_sec_pos_AT_MT=datos.get('38', {}).get('valor_texto', '').replace(",", "."),
                        imp_sec_pos_MT_BT=datos.get('43', {}).get('valor_texto', '').replace(",", "."),
                        pot_sec_pos_MT_BT=datos.get('44', {}).get('valor_texto', '').replace(",", "."),
                        imp_sec_pos_BT_AT=datos.get('49', {}).get('valor_texto', '').replace(",", "."),
                        pot_sec_pos_BT_AT=datos.get('50', {}).get('valor_texto', '').replace(",", "."),
                        imp_sec_cero_AT_MT=datos.get('6279', {}).get('valor_texto', '').replace(",", "."),
                        pot_sec_cero_AT_MT=datos.get('6280', {}).get('valor_texto', '').replace(",", "."),
                        imp_sec_cero_MT_BT=datos.get('6478', {}).get('valor_texto', '').replace(",", "."),
                        pot_sec_cero_MT_BT=datos.get('6479', {}).get('valor_texto', '').replace(",", "."),
                        imp_sec_cero_BT_AT=datos.get('6281', {}).get('valor_texto', '').replace(",", "."),
                        pot_sec_cero_BT_AT=datos.get('6282', {}).get('valor_texto', '').replace(",", "."),
                        estados_certificacion=estados_certificacion
                    )
                elif tipo_equipo == 'desconectadores':
                    result = DesconectadorDatosTecnicos(
                        nivel_tension=datos.get('5699', {}).get('valor_texto', '').replace(",", "."),
                        corriente_nominal=datos.get('6216', {}).get('valor_texto', '').replace(",", "."),
                        tipo_desconectador=datos.get('7986', {}).get('valor_texto', ''),
                        estados_certificacion=estados_certificacion
                    )
                elif tipo_equipo == 'transformadores_corriente':
                    result = TransformadorCorrienteDatosTecnicos(
                        relacion_transformacion=datos.get('458', {}).get('valor_texto', ''),
                        precision=datos.get('460', {}).get('valor_texto', ''),
                        corriente_primaria=datos.get('6177', {}).get('valor_texto', '').replace(",", "."),
                        elemento_asociado=datos.get('4813', {}).get('valor_texto', ''),
                        proteccion_asociada=datos.get('5651', {}).get('valor_texto', ''),
                        estados_certificacion=estados_certificacion
                    )
                elif tipo_equipo == 'trampas_ondas':
                    result = TrampasOndasDatosTecnicos(
                        linea_asociada=datos.get('4868', {}).get('valor_texto', ''),
                        corriente_nominal=datos.get('469', {}).get('valor_texto', '').replace(",", "."),
                        estados_certificacion=estados_certificacion
                    )
                elif tipo_equipo == 'unidades_generadoras':
                    result = UnidadesDatosTecnicos(
                        subestacion_nombre=datos.get('7875', {}).get('valor_texto', ''),
                        tecnologia=datos.get('4582', {}).get('valor_texto', ''),
                        tension_nominal=datos.get('7247', {}).get('valor_texto', '').replace(",", "."),
                        potencia_neta=datos.get('590', {}).get('valor_texto', '').replace(",", "."),
                        minimo_tecnico=datos.get('7928', {}).get('valor_texto', '').replace(",", "."),
                        estados_certificacion=estados_certificacion
                    )
        self.datos_tecnicos_cache[cache_key] = result
        return result

    async def obtener_datos_tramo(self, id_tramo: int) -> Dict[str, Any]:
        url = f"{BASE_URLS['tramos']}/{id_tramo}/"
        datos_tramo = await self.hacer_solicitud(url)
        if datos_tramo:
            return {
                'id_tramo': id_tramo,
                'nombre_tramo': datos_tramo.get('nombre', ''),
                'extremo1': self.limpiar_extremos(datos_tramo.get('extremo1_descripcion', '')),
                'extremo2': self.limpiar_extremos(datos_tramo.get('extremo2_descripcion', '')),
            }
        return {'id_tramo': id_tramo, 'nombre_tramo': '', 'extremo1': '', 'extremo2': ''}

    def limpiar_extremos(self, texto: str) -> str:
        if texto is None: return ''
        patrones = [r'^Paño\s*:\s*', r'^Tap\s*:\s*']
        for patron in patrones:
            texto = re.sub(patron, '', texto, flags=re.IGNORECASE)
        return texto.strip()

    async def obtener_secciones_por_linea(self, id_linea: int) -> List[int]:
        url_linea = f"{BASE_URLS['lineas']}/{id_linea}/"
        datos_linea = await self.hacer_solicitud(url_linea)
        if not datos_linea: return []
        nemotecnico_linea = datos_linea.get('nemotecnico', '')
        if not nemotecnico_linea: return []
        nemotecnico_linea_clave = nemotecnico_linea[:5] + nemotecnico_linea[-5:]
        
        url_secciones = f"{BASE_URLS['secciones_tramos']}"
        datos_secciones = await self.hacer_solicitud(url_secciones)
        if not datos_secciones: return []

        ids_secciones_tramos = []
        for seccion in datos_secciones:
            nemotecnico_seccion = seccion.get('nemotecnico', '')
            if nemotecnico_seccion:
                nemotecnico_seccion_clave = nemotecnico_seccion[:5] + nemotecnico_seccion[12:17]
                if nemotecnico_seccion_clave == nemotecnico_linea_clave:
                    ids_secciones_tramos.append(seccion['id'])
        return ids_secciones_tramos

async def procesar_interruptor(id_interruptor: int, api_client: ApiClient) -> Optional[Interruptor]:
    url_interruptor = f"{BASE_URLS['interruptores']}/{id_interruptor}"
    datos_interruptor = await api_client.hacer_solicitud(url_interruptor)
    if not datos_interruptor: return None
    datos_tecnicos = await api_client.obtener_datos_tecnicos(id_interruptor, 'interruptores')
    return Interruptor('Interruptor', id_interruptor, datos_interruptor.get('nombre', ''), datos_interruptor.get('subestacion_nombre', ''), datos_interruptor.get('propietario_nombre', ''), datos_interruptor.get('pano_nombre', ''), datos_tecnicos)

async def procesar_transformador(id_transformador: int, api_client: ApiClient) -> Optional[Transformador]:
    url_transformador = f"{BASE_URLS['transformadores_2d']}/{id_transformador}"
    datos_transformador = await api_client.hacer_solicitud(url_transformador)
    if not datos_transformador: return None
    datos_tecnicos = await api_client.obtener_datos_tecnicos(id_transformador, 'transformadores_2d')
    return Transformador('Transformador 2D', id_transformador, datos_transformador.get('nombre', ''), datos_transformador.get('subestacion_nombre', ''), datos_transformador.get('propietario_nombre', ''), datos_transformador.get('coordinado_nombre', ''), datos_tecnicos)

async def procesar_transformador_3d(id_transformador: int, api_client: ApiClient) -> Optional[Transformador3D]:
    url_transformador = f"{BASE_URLS['transformadores_3d']}/{id_transformador}"
    datos_transformador = await api_client.hacer_solicitud(url_transformador)
    if not datos_transformador: return None
    datos_tecnicos = await api_client.obtener_datos_tecnicos(id_transformador, 'transformadores_3d')
    return Transformador3D('Transformador 3D', id_transformador, datos_transformador.get('nombre', ''), datos_transformador.get('subestacion_nombre', ''), datos_transformador.get('propietario_nombre', ''), datos_transformador.get('coordinado_nombre', ''), datos_tecnicos)

async def procesar_desconectador(id_desconectador: int, api_client: ApiClient) -> Optional[Desconectador]:
    url_desconectador = f"{BASE_URLS['desconectadores']}/{id_desconectador}"
    datos_desconectador = await api_client.hacer_solicitud(url_desconectador)
    if not datos_desconectador: return None
    datos_tecnicos = await api_client.obtener_datos_tecnicos(id_desconectador, 'desconectadores')
    return Desconectador('Desconectador', id_desconectador, datos_desconectador.get('nombre', ''), datos_desconectador.get('subestacion_nombre', ''), datos_desconectador.get('propietario_nombre', ''), datos_desconectador.get('pano_nombre', ''), datos_tecnicos)

async def procesar_transformador_corriente(id_transformador_corriente: int, api_client: ApiClient) -> Optional[TransformadorCorriente]:
    url_transformador_corriente = f"{BASE_URLS['transformadores_corriente']}/{id_transformador_corriente}"
    datos_transformador_corriente = await api_client.hacer_solicitud(url_transformador_corriente)
    if not datos_transformador_corriente: return None
    datos_tecnicos = await api_client.obtener_datos_tecnicos(id_transformador_corriente, 'transformadores_corriente')
    return TransformadorCorriente('Transformador Corriente', id_transformador_corriente, datos_transformador_corriente.get('nombre', ''), datos_transformador_corriente.get('subestacion_nombre', ''), datos_transformador_corriente.get('propietario_nombre', ''), datos_transformador_corriente.get('pano_nombre', ''), datos_tecnicos)

async def procesar_trampa_onda(id_trampa_onda: int, api_client: ApiClient) -> Optional[TrampaOnda]:
    url_trampa_onda = f"{BASE_URLS['trampas_ondas']}/{id_trampa_onda}"
    datos_trampa_onda = await api_client.hacer_solicitud(url_trampa_onda)
    if not datos_trampa_onda: return None
    datos_tecnicos = await api_client.obtener_datos_tecnicos(id_trampa_onda, 'trampas_ondas')
    return TrampaOnda('Trampa Onda', id_trampa_onda, datos_trampa_onda.get('nombre', ''), datos_trampa_onda.get('subestacion_nombre', ''), datos_trampa_onda.get('propietario_nombre', ''), datos_trampa_onda.get('pano_nombre', ''), datos_tecnicos)

async def procesar_unidades(id_unidad: int, api_client: ApiClient) -> Optional[Unidad]:
    url_unidad = f"{BASE_URLS['unidades_generadoras']}/{id_unidad}"
    datos_unidad_generadora = await api_client.hacer_solicitud(url_unidad)
    if not datos_unidad_generadora: return None
    datos_tecnicos = await api_client.obtener_datos_tecnicos(id_unidad, 'unidades_generadoras')
    return Unidad('Unidad Generadora', id_unidad, datos_unidad_generadora.get('nombre', ''), None, datos_unidad_generadora.get('propietario_nombre', ''), None, datos_unidad_generadora.get('central_nombre', ''), datos_tecnicos)

async def procesar_seccion_tramo(id_seccion_tramo: int, api_client: ApiClient) -> Optional[SeccionTramo]:
    url_seccion_tramo = f"{BASE_URLS['secciones_tramos']}/{id_seccion_tramo}/"
    datos_seccion_tramo = await api_client.hacer_solicitud(url_seccion_tramo)
    if not datos_seccion_tramo: return None
    datos_tecnicos = await api_client.obtener_datos_tecnicos(id_seccion_tramo, 'secciones_tramos')
    tramo_info = await api_client.obtener_datos_tramo(datos_seccion_tramo.get('id_tramo', 0))
    return SeccionTramo('Sección Tramo', id_seccion_tramo, datos_seccion_tramo.get('nombre', ''), None, datos_seccion_tramo.get('propietario_nombre', ''), None, datos_seccion_tramo.get('id_linea', 0), datos_seccion_tramo.get('linea_nombre', ''), tramo_info.get('id_tramo', 0), tramo_info.get('nombre_tramo', ''), tramo_info.get('extremo1', ''), tramo_info.get('extremo2', ''), datos_tecnicos)

async def obtener_equipos_por_subestacion(id_subestacion: int, api_client: ApiClient, tipo_equipo: str) -> List[Equipo]:
    nombre_subestacion = await api_client.obtener_nombre_subestacion(id_subestacion)
    if not nombre_subestacion: return []
    equipos_data = await api_client.buscar_equipos_por_subestacion_nombre(nombre_subestacion, tipo_equipo)
    tareas = []
    if tipo_equipo == 'interruptores': tareas = [procesar_interruptor(e['id'], api_client) for e in equipos_data]
    elif tipo_equipo == 'transformadores_2d': tareas = [procesar_transformador(e['id'], api_client) for e in equipos_data]
    elif tipo_equipo == 'transformadores_3d': tareas = [procesar_transformador_3d(e['id'], api_client) for e in equipos_data]
    elif tipo_equipo == 'desconectadores': tareas = [procesar_desconectador(e['id'], api_client) for e in equipos_data]
    elif tipo_equipo == 'transformadores_corriente': tareas = [procesar_transformador_corriente(e['id'], api_client) for e in equipos_data]
    elif tipo_equipo == 'trampas_ondas': tareas = [procesar_trampa_onda(e['id'], api_client) for e in equipos_data]
    resultados = await asyncio.gather(*tareas)
    return [r for r in resultados if r is not None]

def convertir_a_flotante(valor: Optional[str]) -> float:
    try:
        return float(valor.replace(",", "."))
    except (ValueError, AttributeError):
        return 0.0

def calcular_max_valores(*valores: float) -> Optional[float]:
    valores_flotantes = [v for v in valores if v]
    return max(valores_flotantes) if valores_flotantes else None

def calcular_min_valores(*valores: float) -> Optional[float]:
    valores_flotantes = [v for v in valores if v]
    return min(valores_flotantes) if valores_flotantes else None

def realizar_operaciones(datos: List[Transformador]) -> List[Dict[str, Any]]:
    resultados = []
    for transformador in datos:
        dt = transformador.datos_tecnicos
        if dt:
            capacidad_max_at = calcular_max_valores(convertir_a_flotante(dt.capacidad_at), convertir_a_flotante(dt.capacidad_at_onaf1), convertir_a_flotante(dt.capacidad_at_onaf2))
            pot_sec_pos = convertir_a_flotante(dt.pot_sec_pos)
            pot_sec_cero = convertir_a_flotante(dt.pot_sec_cero)
            resultados.append({
                'ID Transformador': transformador.id_equipo,
                'Nombre Transformador': transformador.nombre_equipo,
                'Z_pos_[%]': convertir_a_flotante(dt.imp_sec_pos) * (capacidad_max_at / pot_sec_pos) if capacidad_max_at and pot_sec_pos else None,
                'Z_cero_[%]': convertir_a_flotante(dt.imp_sec_cero) * (capacidad_max_at / pot_sec_cero) if capacidad_max_at and pot_sec_cero else None,
                'Pot_base_[MVA]': capacidad_max_at,
            })
    return resultados

def realizar_operaciones_3d(datos: List[Transformador3D]) -> List[Dict[str, Any]]:
    resultados = []
    for transformador in datos:
        dt = transformador.datos_tecnicos
        if dt:
            capacidad_max_at = calcular_max_valores(convertir_a_flotante(dt.capacidad_at), convertir_a_flotante(dt.capacidad_at_onaf1), convertir_a_flotante(dt.capacidad_at_onaf2))
            capacidad_max_mt = calcular_max_valores(convertir_a_flotante(dt.capacidad_mt), convertir_a_flotante(dt.capacidad_mt_onaf1), convertir_a_flotante(dt.capacidad_mt_onaf2))
            capacidad_max_bt = calcular_max_valores(convertir_a_flotante(dt.capacidad_bt), convertir_a_flotante(dt.capacidad_bt_onaf1), convertir_a_flotante(dt.capacidad_bt_onaf2))
            min_at_mt = calcular_min_valores(capacidad_max_at, capacidad_max_mt)
            min_mt_bt = calcular_min_valores(capacidad_max_mt, capacidad_max_bt)
            min_bt_at = calcular_min_valores(capacidad_max_bt, capacidad_max_at)
            resultados.append({
                'ID Transformador': transformador.id_equipo,
                'Nombre Transformador': transformador.nombre_equipo,
                'Z_pos_AT-MT_[%]': convertir_a_flotante(dt.imp_sec_pos_AT_MT) * (min_at_mt / convertir_a_flotante(dt.pot_sec_pos_AT_MT)) if min_at_mt and dt.pot_sec_pos_AT_MT else None,
                'Z_pos_MT-BT_[%]': convertir_a_flotante(dt.imp_sec_pos_MT_BT) * (min_mt_bt / convertir_a_flotante(dt.pot_sec_pos_MT_BT)) if min_mt_bt and dt.pot_sec_pos_MT_BT else None,
                'Z_pos_BT-AT_[%]': convertir_a_flotante(dt.imp_sec_pos_BT_AT) * (min_bt_at / convertir_a_flotante(dt.pot_sec_pos_BT_AT)) if min_bt_at and dt.pot_sec_pos_BT_AT else None,
                'Pot_base_AT_MT_[MVA]': min_at_mt,
                'Pot_base_MT_BT_[MVA]': min_mt_bt,
                'Pot_base_BT_AT_[MVA]': min_bt_at,
            })
    return resultados

def colorear_celda(celda: Any, estado_certificacion: str) -> None:
    colores = {'Validado': (204, 255, 204), 'Rechazado': (244, 176, 132), 'En Uso': (255, 230, 153)}
    if estado_certificacion in colores:
        r, g, b = colores[estado_certificacion]
        celda.Interior.Color = (b << 16) + (g << 8) + r

def crear_archivo_excel(datos_im, datos_t2d, datos_t3d, datos_st, datos_des, datos_tc, datos_to, datos_ug) -> str:
    """Crea y guarda el archivo Excel devolviendo la ruta absoluta del mismo."""
    excel = win32.DispatchEx('Excel.Application')
    excel.Visible = False
    workbook = excel.Workbooks.Add()
    while workbook.Sheets.Count > 1: workbook.Sheets(1).Delete()

    # --- Generación de pestañas condicionales ---
    if datos_im:
        sheet = workbook.Sheets.Add(); sheet.Name = "IM"
        sheet.Cells(1,1).Value = "ID Interruptor"; sheet.Cells(1,2).Value = "Nombre Interruptor"
        for idx, item in enumerate(datos_im, start=2):
            sheet.Cells(idx, 1).Value = item.id_equipo; sheet.Cells(idx, 2).Value = item.nombre_equipo
    if datos_t2d:
        sheet = workbook.Sheets.Add(); sheet.Name = "T2D"
        sheet.Cells(1,1).Value = "ID Transformador"; sheet.Cells(1,2).Value = "Nombre Transformador"
        for idx, item in enumerate(datos_t2d, start=2):
            sheet.Cells(idx, 1).Value = item.id_equipo; sheet.Cells(idx, 2).Value = item.nombre_equipo
    if datos_st:
        sheet = workbook.Sheets.Add(); sheet.Name = "Secciones Tramos"
        sheet.Cells(1,1).Value = "ID Sección Tramo"; sheet.Cells(1,2).Value = "Nombre Sección"
        for idx, item in enumerate(datos_st, start=2):
            sheet.Cells(idx, 1).Value = item.id_equipo; sheet.Cells(idx, 2).Value = item.nombre_equipo

    excel.DisplayAlerts = False
    filepath = os.path.abspath(os.path.join(os.getcwd(), "equipos_datos.xlsx"))
    workbook.SaveAs(filepath)
    workbook.Close()
    excel.Quit()
    return filepath

async def exportar_datos_async(ids_sub=None, ids_int=None, ids_t2d=None, ids_t3d=None, ids_st=None, ids_lin=None, ids_des=None, ids_tc=None, ids_to=None, ids_ug=None, por_secciones=False) -> str:
    async with aiohttp.ClientSession() as session:
        api_client = ApiClientFactory.create_client(session)
        
        r_int, r_t2d, r_t3d, r_st, r_des, r_tc, r_to, r_ug = [], [], [], [], [], [], [], []

        if ids_sub:
            for s_id in ids_sub:
                r_int.extend(await obtener_equipos_por_subestacion(s_id, api_client, 'interruptores'))
                r_t2d.extend(await obtener_equipos_por_subestacion(s_id, api_client, 'transformadores_2d'))
        if ids_int:
            r_int.extend([await procesar_interruptor(i, api_client) for i in ids_int if i])
        if ids_t2d:
            r_t2d.extend([await procesar_transformador(i, api_client) for i in ids_t2d if i])
        if ids_st:
            r_st.extend([await procesar_seccion_tramo(i, api_client) for i in ids_st if i])

        r_int = [x for x in r_int if x]; r_t2d = [x for x in r_t2d if x]; r_st = [x for x in r_st if x]
        
        if not any([r_int, r_t2d, r_t3d, r_st, r_des, r_tc, r_to, r_ug]):
            raise ValueError("No se encontraron registros válidos en Infotécnica.")

        return crear_archivo_excel(r_int, r_t2d, r_t3d, r_st, r_des, r_tc, r_to, r_ug)
