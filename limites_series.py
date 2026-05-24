import asyncio
import aiohttp
import logging
import os
import re
from typing import Any, Dict, List, Optional
from openpyxl import Workbook
from openpyxl.styles import PatternFill

BASE_URLS = {
    'interruptores': 'https://api-infotecnica.coordinador.cl/v1/interruptores',
    'transformadores_2d': 'https://api-infotecnica.coordinador.cl/v1/transformadores-2d',
    'transformadores_3d': 'https://api-infotecnica.coordinador.cl/v1/transformadores-3d',
    'secciones_tramos': 'https://api-infotecnica.coordinador.cl/v1/secciones-tramos',
    'desconectadores': 'https://api-infotecnica.coordinador.cl/v1/desconectadores',
    'transformadores_corriente': 'https://api-infotecnica.coordinador.cl/v1/transformadores-corrientes',
    'trampas_ondas': 'https://api-infotecnica.coordinador.cl/v1/trampas-ondas',
    'panos': 'https://api-infotecnica.coordinador.cl/v1/panos',
}

HEADERS = {
    "User-Agent": "Mozilla/5.0",
    "Accept": "application/json, text/plain, */*",
    "Referer": "https://infotecnica.coordinador.cl/",
    "Origin": "https://infotecnica.coordinador.cl",
    "Connection": "keep-alive",
}

def limpiar_valor_corriente(texto_valor: Any) -> float:
    try:
        if not texto_valor: return float('inf')
        num_str = re.sub(r'[^\d,\.]', '', str(texto_valor)).replace(",", ".")
        return float(num_str) if num_str else float('inf')
    except Exception:
        return float('inf')

async def buscar_limites_series_motor(list_ids: List[int], es_modo_tramo: bool) -> str:
    async with aiohttp.ClientSession() as session:
        resumen_panos = []

        for eq_id in list_ids:
            pano_nombres_a_buscar = []
            subestacion = 'Desconocida'

            if es_modo_tramo:
                # ==============================================================
                # LÓGICA MODO TRAMO: El ID corresponde a Secciones de Líneas
                # ==============================================================
                url_tramo = f"{BASE_URLS['secciones_tramos']}/{eq_id}/"
                async with session.get(url_tramo, headers=HEADERS) as resp:
                    if resp.status == 200:
                        data_tramo = await resp.json()
                        subestacion = data_tramo.get('subestacion_nombre') or 'Línea de Transmisión'
                        
                        # Consultamos los paños asociados a los extremos de esta línea
                        url_panos_tramo = f"{BASE_URLS['secciones_tramos']}/{eq_id}/panos/"
                        async with session.get(url_panos_tramo, headers=HEADERS) as resp_p:
                            if resp_p.status == 200:
                                panos_asociados = await resp_p.json()
                                for p in panos_asociados:
                                    if p.get('nemotecnico'):
                                        pano_nombres_a_buscar.append(p.get('nemotecnico'))
            else:
                # ==============================================================
                # LÓGICA MODO DIRECTO: El ID corresponde a Equipos Directos
                # ==============================================================
                # 1. Intentar como Interruptor
                url_eq = f"{BASE_URLS['interruptores']}/{eq_id}"
                async with session.get(url_eq, headers=HEADERS) as resp:
                    data_eq = await resp.json() if resp.status == 200 else None
                if data_eq:
                    subestacion = data_eq.get('subestacion_nombre', 'Desconocida')
                    if data_eq.get('pano_nombre'):
                        pano_nombres_a_buscar.append(data_eq.get('pano_nombre'))

                # 2. Intentar como Trafo 2D
                if not pano_nombres_a_buscar:
                    url_trafo2d = f"{BASE_URLS['transformadores_2d']}/{eq_id}"
                    async with session.get(url_trafo2d, headers=HEADERS) as resp:
                        data_eq = await resp.json() if resp.status == 200 else None
                    if data_eq:
                        subestacion = data_eq.get('subestacion_nombre', 'Desconocida')
                        p_nom = data_eq.get('pano_nombre') or data_eq.get('coordinado_nombre')
                        if p_nom:
                            match = re.search(r'Paño\s+([A-Za-z0-9_-]+)', p_nom, re.IGNORECASE)
                            pano_nombres_a_buscar.append(match.group(1) if match else p_nom)

                # 3. Intentar como Trafo 3D
                if not pano_nombres_a_buscar:
                    url_trafo3d = f"{BASE_URLS['transformadores_3d']}/{eq_id}"
                    async with session.get(url_trafo3d, headers=HEADERS) as resp:
                        data_eq = await resp.json() if resp.status == 200 else None
                    if data_eq:
                        subestacion = data_eq.get('subestacion_nombre', 'Desconocida')
                        p_nom = data_eq.get('pano_nombre') or data_eq.get('coordinado_nombre')
                        if p_nom:
                            match = re.search(r'Paño\s+([A-Za-z0-9_-]+)', p_nom, re.IGNORECASE)
                            pano_nombres_a_buscar.append(match.group(1) if match else p_nom)

            if not pano_nombres_a_buscar: continue

            # ==============================================================
            # DETERMINACIÓN FINALIZADA: BUSCAR ELEMENTOS EN SERIE DEL PAÑO
            # ==============================================================
            for pano_nombre in set(pano_nombres_a_buscar):
                if not pano_nombre: continue
                
                endpoints_series = ['interruptores', 'desconectadores', 'transformadores_corriente', 'trampas_ondas']
                sub_equipos_encontrados = []

                for tipo in endpoints_series:
                    url_search = f"{BASE_URLS[tipo]}/?search={pano_nombre}"
                    async with session.get(url_search, headers=HEADERS) as resp:
                        datos_api = await resp.json() if resp.status == 200 else []
                    
                    for item in datos_api:
                        url_ficha = f"{BASE_URLS[tipo]}/{item['id']}/fichas-tecnicas/general/"
                        async with session.get(url_ficha, headers=HEADERS) as resp:
                            ficha = await resp.json() if resp.status == 200 else None
                        
                        if ficha:
                            id_campo_corr = '6019' if tipo == 'interruptores' else '6216' if tipo == 'desconectadores' else '6177' if tipo == 'transformadores_corriente' else '469'
                            txt_corr = ficha.get(id_campo_corr, {}).get('valor_texto', '')
                            valor_amp = limpiar_valor_corriente(txt_corr)
                            
                            if valor_amp != float('inf'):
                                sub_equipos_encontrados.append({
                                    'id': item['id'],
                                    'nombre': item.get('nombre', f"{tipo}_{item['id']}"),
                                    'tipo': tipo.replace("_", " ").upper(),
                                    'corriente': valor_amp
                                })

                if sub_equipos_encontrados:
                    limitante = min(sub_equipos_encontrados, key=lambda x: x['corriente'])
                    resumen_panos.append({
                        'id_consultado': eq_id,
                        'subestacion': subestacion,
                        'paño': pano_nombre,
                        'limite_paño_amp': limitante['corriente'],
                        'equipo_limitante': f"{limitante['nombre']} ({limitante['tipo']})",
                        'total_elementos_serie': len(sub_equipos_encontrados)
                    })

        if not resumen_panos:
            raise ValueError("No se encontraron elementos en serie para los IDs y el modo especificado.")

        wb = Workbook()
        ws = wb.active
        ws.title = "Límites de Corriente del Paño"
        ws.append(['ID Consultado', 'Subestación / Elemento', 'Nombre del Paño Coordinado', 'Límite Térmico del Paño [A]', 'Componente Limitante (Eslabón Más Débil)', 'Elementos Evaluados en Serie'])
        
        for p in resumen_panos:
            ws.append([p['id_consultado'], p['subestacion'], p['paño'], p['limite_paño_amp'], p['equipo_limitante'], p['total_elementos_serie']])
            for col in range(1, 7):
                ws.cell(row=ws.max_row, column=col).fill = PatternFill(start_color='FFE699', end_color='FFE699', fill_type='solid')

        filepath = os.path.abspath(os.path.join(os.getcwd(), "equipos_datos.xlsx"))
        wb.save(filepath)
        return filepath
