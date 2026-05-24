import asyncio
import aiohttp
import logging
import os
import re
from typing import Any, Dict, List, Optional
from openpyxl import Workbook
from openpyxl.styles import PatternFill, Font, Alignment

BASE_URLS = {
    'interruptores': 'https://api-infotecnica.coordinador.cl/v1/interruptores',
    'transformadores_2d': 'https://api-infotecnica.coordinador.cl/v1/transformadores-2d',
    'transformadores_3d': 'https://api-infotecnica.coordinador.cl/v1/transformadores-3d',
    'secciones_tramos': 'https://api-infotecnica.coordinador.cl/v1/secciones-tramos',
    'tramos': 'https://api-infotecnica.coordinador.cl/v1/tramos',
    'desconectadores': 'https://api-infotecnica.coordinador.cl/v1/desconectadores',
    'transformadores_corriente': 'https://api-infotecnica.coordinador.cl/v1/transformadores-corrientes',
    'trampas_ondas': 'https://api-infotecnica.coordinador.cl/v1/trampas-ondas',
    'panos': 'https://api-infotecnica.coordinador.cl/v1/panos',
}

HEADERS = {
    "User-Agent": "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko)",
    "Accept": "application/json, text/plain, */*",
    "Referer": "https://infotecnica.coordinador.cl/",
    "Origin": "https://infotecnica.coordinador.cl",
    "Connection": "keep-alive",
}

def limpiar_valor_float(texto_valor: Any) -> float:
    try:
        if not texto_valor: return float('inf')
        num_str = re.sub(r'[^\d,\.]', '', str(texto_valor)).replace(",", ".")
        return float(num_str) if num_str else float('inf')
    except Exception:
        return float('inf')

async def hacer_solicitud(session: aiohttp.ClientSession, url: str) -> Optional[Any]:
    timeout = aiohttp.ClientTimeout(total=8.0)
    try:
        async with session.get(url, headers=HEADERS, timeout=timeout) as response:
            if response.status == 200:
                return await response.json()
    except Exception:
        pass
    return None

def limpiar_nombre_instalacion(texto: str) -> str:
    if not texto: return ""
    texto_limpio = re.sub(r'^(Paño\s*:\s*|Tap\s*:\s*|Paño\s+|S/E\s+)', '', texto, flags=re.IGNORECASE)
    if " - " in texto_limpio:
        texto_limpio = texto_limpio.split(" - ")[0]
    return texto_limpio.strip()

async def buscar_limites_series_motor(list_ids: List[int], es_modo_tramo: bool) -> str:
    connector = aiohttp.TCPConnector(limit_per_host=2)
    async with aiohttp.ClientSession(connector=connector) as session:
        
        wb = Workbook()
        ws = wb.active
        ws.title = "Análisis de Elementos en Serie"
        
        headers = [
            'ID Consultado', 'Subestación / Elemento', 'Paño Coordinado', 
            'Equipo en Serie', 'Tipo de Equipo', 'Capacidad Corriente [A]', 
            'Cap. Ruptura Simétrica [kA]', 'Elemento Limitante Corriente', 'Elemento Limitante Ruptura'
        ]
        ws.append(headers)
        
        header_fill = PatternFill(start_color='1F4E78', end_color='1F4E78', fill_type='solid')
        header_font = Font(name='Arial', size=11, bold=True, color='FFFFFF')
        for col_num in range(1, len(headers) + 1):
            cell = ws.cell(row=1, column=col_num)
            cell.fill = header_fill
            cell.font = header_font
            cell.alignment = Alignment(horizontal='center', vertical='center', wrap_text=True)
            
        datos_agregados = False
        paños_ya_procesados = set()

        for eq_id in list_ids:
            pano_nombres_a_buscar = []
            subestacion = 'Desconocida'

            if es_modo_tramo:
                # ==============================================================
                # MODO TRAMO
                # ==============================================================
                url_seccion = f"{BASE_URLS['secciones_tramos']}/{eq_id}/"
                data_seccion = await hacer_solicitud(session, url_seccion)
                
                if data_seccion and isinstance(data_seccion, dict):
                    subestacion = data_seccion.get('linea_nombre') or data_seccion.get('nombre') or 'Línea de Transmisión'
                    
                    for llave_txt in ['linea_nombre', 'nombre', 'extremo1_descripcion', 'extremo2_descripcion']:
                        val_txt = data_seccion.get(llave_txt, '')
                        if val_txt:
                            txt_limpio = limpiar_nombre_instalacion(str(val_txt))
                            if txt_limpio and len(txt_limpio) > 2:
                                url_p = f"{BASE_URLS['panos']}/?search={txt_limpio}"
                                panos_res = await hacer_solicitud(session, url_p)
                                if panos_res and isinstance(panos_res, list):
                                    for p in panos_res:
                                        if p.get('nemotecnico'):
                                            pano_nombres_a_buscar.append(p.get('nemotecnico'))
            else:
                # ==============================================================
                # MODO DIRECTO
                # ==============================================================
                url_eq = f"{BASE_URLS['interruptores']}/{eq_id}"
                data_eq = await hacer_solicitud(session, url_eq)
                if data_eq and isinstance(data_eq, dict):
                    subestacion = data_eq.get('subestacion_nombre', 'Desconocida')
                    if data_eq.get('pano_nombre'):
                        pano_nombres_a_buscar.append(data_eq.get('pano_nombre'))

                if not pano_nombres_a_buscar:
                    url_trafo2d = f"{BASE_URLS['transformadores_2d']}/{eq_id}"
                    data_eq = await hacer_solicitud(session, url_trafo2d)
                    if data_eq and isinstance(data_eq, dict):
                        subestacion = data_eq.get('subestacion_nombre', 'Desconocida')
                        p_nom = data_eq.get('pano_nombre') or data_eq.get('coordinado_nombre')
                        if p_nom:
                            pano_nombres_a_buscar.append(limpiar_nombre_instalacion(p_nom))

                if not pano_nombres_a_buscar:
                    url_trafo3d = f"{BASE_URLS['transformadores_3d']}/{eq_id}"
                    data_eq = await hacer_solicitud(session, url_trafo3d)
                    if data_eq and isinstance(data_eq, dict):
                        subestacion = data_eq.get('subestacion_nombre', 'Desconocida')
                        p_nom = data_eq.get('pano_nombre') or data_eq.get('coordinado_nombre')
                        if p_nom:
                            pano_nombres_a_buscar.append(limpiar_nombre_instalacion(p_nom))

            pano_nombres_a_buscar = list(set(pano_nombres_a_buscar))
            if not pano_nombres_a_buscar: continue

            # ==============================================================
            # COSECHA VERTICAL DE ELEMENTOS EN SERIE
            # ==============================================================
            for pano_nombre in pano_nombres_a_buscar:
                if not pano_nombre or pano_nombre in paños_ya_procesados: continue
                paños_ya_procesados.add(pano_nombre)
                
                endpoints_series = ['interruptores', 'desconectadores', 'transformadores_corriente', 'trampas_ondas']
                sub_equipos_encontrados = []

                for tipo in endpoints_series:
                    url_search = f"{BASE_URLS[tipo]}/?search={pano_nombre}"
                    datos_api = await hacer_solicitud(session, url_search)
                    
                    if datos_api and isinstance(datos_api, list):
                        # Reemplazamos el filtro estricto de texto por un límite físico seguro (máximo 10 fichas por endpoint)
                        # Esto permite capturar equipos con sutiles variaciones de nombre sin generar ráfagas infinitas
                        for item in datos_api[:10]:
                            
                            await asyncio.sleep(0.15)
                            url_ficha = f"{BASE_URLS[tipo]}/{item['id']}/fichas-tecnicas/general/"
                            ficha = await hacer_solicitud(session, url_ficha)
                            
                            if ficha and isinstance(ficha, dict):
                                id_campo_corr = '6019' if tipo == 'interruptores' else '6216' if tipo == 'desconectadores' else '6177' if tipo == 'transformadores_corriente' else '469'
                                txt_corr = ficha.get(id_campo_corr, {}).get('valor_texto', '')
                                valor_amp = limpiar_valor_float(txt_corr)
                                
                                valor_ruptura = float('inf')
                                if tipo == 'interruptores':
                                    txt_rup = ficha.get('326', {}).get('valor_texto', '')
                                    valor_ruptura = limpiar_valor_float(txt_rup)

                                if valor_amp != float('inf') or valor_ruptura != float('inf'):
                                    sub_equipos_encontrados.append({
                                        'id': item['id'],
                                        'nombre': item.get('nombre', f"{tipo}_{item['id']}"),
                                        'tipo': tipo.replace("_", " ").upper(),
                                        'corriente': valor_amp,
                                        'ruptura': valor_ruptura
                                    })

                if sub_equipos_encontrados:
                    datos_agregados = True
                    
                    limitante_corriente = min(sub_equipos_encontrados, key=lambda x: x['corriente'])
                    equipos_con_ruptura = [x for x in sub_equipos_encontrados if x['ruptura'] != float('inf')]
                    limitante_ruptura = min(equipos_con_ruptura, key=lambda x: x['ruptura']) if equipos_con_ruptura else None

                    inicio_bloque_row = ws.max_row + 1
                    
                    for eq in sub_equipos_encontrados:
                        corr_display = eq['corriente'] if eq['corriente'] != float('inf'] else 'N/A'
                        rup_display = eq['ruptura'] if eq['ruptura'] != float('inf') else 'N/A'
                        
                        ws.append([
                            eq_id, subestacion, pano_nombre, 
                            eq['nombre'], eq['tipo'], corr_display, rup_display,
                            "", ""
                        ])
                    
                    fin_bloque_row = ws.max_row
                    
                    txt_lim_corr = f"{limitante_corriente['nombre']} ({limitante_corriente['corriente']} A)"
                    txt_lim_rup = f"{limitante_ruptura['nombre']} ({limitante_ruptura['ruptura']} kA)" if limitante_ruptura else "N/A"
                    
                    ws.cell(row=inicio_bloque_row, column=8, value=txt_lim_corr)
                    ws.cell(row=inicio_bloque_row, column=9, value=txt_lim_rup)
                    
                    block_fill = PatternFill(start_color='FFF2CC', end_color='FFF2CC', fill_type='solid')
                    
                    for r in range(inicio_bloque_row, fin_bloque_row + 1):
                        for c in range(1, 10):
                            cell = ws.cell(row=r, column=c)
                            if c in [8, 9]:
                                cell.fill = block_fill
                                cell.font = Font(name='Arial', size=10, bold=True, color='C00000')
                            cell.alignment = Alignment(vertical='center', horizontal='left' if c==4 else 'center')

        if not datos_agregados:
            raise ValueError("No se encontraron elementos en serie para los IDs ingresados. Verifica el modo de consulta.")

        for col in ws.columns:
            max_len = max(len(str(cell.value or '')) for cell in col)
            col_letter = col[0].column_letter
            ws.column_dimensions[col_letter].width = max(max_len + 3, 12)

        filepath = os.path.abspath(os.path.join(os.getcwd(), "equipos_datos.xlsx"))
        wb.save(filepath)
        return filepath
