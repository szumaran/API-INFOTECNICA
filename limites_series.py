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

async def buscar_limites_series_motor(list_ids: List[int]) -> str:
    async with aiohttp.ClientSession() as session:
        resumen_panos = []

        for eq_id in list_ids:
            pano_nombres_a_buscar = []
            subestacion = 'Desconocida'

            # 1. INTENTAR COMO TRANSFORMADOR 2D
            url_trafo2d = f"{BASE_URLS['transformadores_2d']}/{eq_id}"
            async with session.get(url_trafo2d, headers=HEADERS) as resp:
                if resp.status == 200:
                    data_trafo = await resp.json()
                    subestacion = data_trafo.get('subestacion_nombre', 'Desconocida')
                    p_nom = data_trafo.get('pano_nombre') or data_trafo.get('coordinado_nombre')
                    if p_nom:
                        match = re.search(r'Paño\s+([A-Za-z0-9_-]+)', p_nom, re.IGNORECASE)
                        pano_nombres_a_buscar.append(match.group(1) if match else p_nom)

            # 2. INTENTAR COMO TRANSFORMADOR 3D (NUEVO)
            if not pano_nombres_a_buscar:
                url_trafo3d = f"{BASE_URLS['transformadores_3d']}/{eq_id}"
                async with session.get(url_trafo3d, headers=HEADERS) as resp:
                    if resp.status == 200:
                        data_trafo3d = await resp.json()
                        subestacion = data_trafo3d.get('subestacion_nombre', 'Desconocida')
                        p_nom = data_trafo3d.get('pano_nombre') or data_trafo3d.get('coordinado_nombre')
                        if p_nom:
                            match = re.search(r'Paño\s+([A-Za-z0-9_-]+)', p_nom, re.IGNORECASE)
                            pano_nombres_a_buscar.append(match.group(1) if match else p_nom)

            # 3. INTENTAR COMO SECCIÓN DE TRAMO / LÍNEA
            if not pano_nombres_a_buscar:
                url_tramo = f"{BASE_URLS['secciones_tramos']}/{eq_id}/"
                async with session.get(url_tramo, headers=HEADERS) as resp:
                    if resp.status == 200:
                        data_tramo = await resp.json()
                        subestacion = data_tramo.get('subestacion_nombre') or 'Línea de Transmisión'
                        
                        # Consultar el endpoint de paños asociados a esta sección de tramo
                        url_panos_tramo = f"{BASE_URLS['secciones_tramos']}/{eq_id}/panos/"
                        async with session.get(url_panos_tramo, headers=HEADERS) as resp_p:
                            if resp_p.status == 200:
                                panos_asociados = await resp_p.json()
                                for p in panos_asociados:
                                    if p.get('nemotecnico'):
                                        pano_nombres_a_buscar.append(p.get('nemotecnico'))

            # Si no se encontró ningún paño por ninguna vía, saltamos al siguiente ID
            if not pano_nombres_a_buscar: continue

            # 4. PROCESAR LOS PAÑOS ENCONTRADOS PARA BUSCAR SUS ELEMENTOS EN SERIE
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
            raise ValueError("No se pudieron levantar los elementos en serie para los IDs provistos en las APIs.")

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
