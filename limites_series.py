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
    "User-Agent": "Mozilla/5.0",
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

async def hacer_solicitud(session: aiohttp.ClientSession, url: str) -> Optional[Dict[str, Any]]:
    try:
        async with session.get(url, headers=HEADERS) as response:
            if response.status == 200:
                return await response.json()
    except Exception:
        pass
    return None

async def buscar_nemotecnico_pano_por_extremo_tramo(session: aiohttp.ClientSession, nombre_extremo: str) -> Optional[str]:
    if not nombre_extremo: return None
    if nombre_extremo.startswith("S/E "):
        nombre_extremo = nombre_extremo[4:]
    url = f"{BASE_URLS['panos']}?nombre__icontains={nombre_extremo}"
    datos = await hacer_solicitud(session, url)
    if datos:
        for pano in datos:
            if pano.get('nombre', '').lower() == nombre_extremo.lower():
                return pano.get('nemotecnico', '')
        return datos[0].get('nemotecnico', '')
    return None

async def buscar_limites_series_motor(list_ids: List[int], es_modo_tramo: bool) -> str:
    async with aiohttp.ClientSession() as session:
        wb = Workbook()
        ws = wb.active
        ws.title = "Análisis de Elementos en Serie"
        
        filas_reporte = []
        max_equipos_detectados = 0

        for eq_id in list_ids:
            pano_nombres_a_buscar = []
            subestacion = 'Desconocida'

            if es_modo_tramo:
                # ==============================================================
                # MODO TRAMO (RESTAURADO AL MOTOR SEGURO DEL EXTRACTOR 1)
                # ==============================================================
                url_seccion = f"{BASE_URLS['secciones_tramos']}/{eq_id}/"
                data_seccion = await hacer_solicitud(session, url_seccion)
                
                if data_seccion and data_seccion.get('id_tramo'):
                    subestacion = data_seccion.get('linea_nombre') or 'Línea de Transmisión'
                    url_tramo = f"{BASE_URLS['tramos']}/{data_seccion['id_tramo']}/"
                    data_tramo = await hacer_solicitud(session, url_tramo)
                    
                    if data_tramo:
                        extremo1 = data_tramo.get('extremo1_descripcion', '')
                        extremo2 = data_tramo.get('extremo2_descripcion', '')
                        for ext in [extremo1, extremo2]:
                            if ext:
                                ext_limpio = re.sub(r'^(Paño\s*:\s*|Tap\s*:\s*|S/E\s*)', '', ext, flags=re.IGNORECASE).strip()
                                nemotecnico = await buscar_nemotecnico_pano_por_extremo_tramo(session, ext_limpio)
                                if nemotecnico:
                                    pano_nombres_a_buscar.append(nemotecnico)
            else:
                # ==============================================================
                # MODO DIRECTO (RESTAURADO AL MOTOR SEGURO DEL EXTRACTOR 1)
                # ==============================================================
                url_eq = f"{BASE_URLS['interruptores']}/{eq_id}"
                data_eq = await hacer_solicitud(session, url_eq)
                if data_eq:
                    subestacion = data_eq.get('subestacion_nombre', 'Desconocida')
                    if data_eq.get('pano_nombre'):
                        pano_nombres_a_buscar.append(data_eq.get('pano_nombre'))

                if not pano_nombres_a_buscar:
                    url_trafo2d = f"{BASE_URLS['transformadores_2d']}/{eq_id}"
                    data_eq = await hacer_solicitud(session, url_trafo2d)
                    if data_eq:
                        subestacion = data_eq.get('subestacion_nombre', 'Desconocida')
                        p_nom = data_eq.get('pano_nombre') or data_eq.get('coordinado_nombre')
                        if p_nom:
                            p_limpio = re.sub(r'^Paño\s*:\s*', '', p_nom, flags=re.IGNORECASE)
                            match = re.search(r'Paño\s+([A-Za-z0-9_-]+)', p_limpio, re.IGNORECASE)
                            pano_nombres_a_buscar.append(match.group(1) if match else p_limpio)

                if not pano_nombres_a_buscar:
                    url_trafo3d = f"{BASE_URLS['transformadores_3d']}/{eq_id}"
                    data_eq = await hacer_solicitud(session, url_trafo3d)
                    if data_eq:
                        subestacion = data_eq.get('subestacion_nombre', 'Desconocida')
                        p_nom = data_eq.get('pano_nombre') or data_eq.get('coordinado_nombre')
                        if p_nom:
                            p_limpio = re.sub(r'^Paño\s*:\s*', '', p_nom, flags=re.IGNORECASE)
                            match = re.search(r'Paño\s+([A-Za-z0-9_-]+)', p_limpio, re.IGNORECASE)
                            pano_nombres_a_buscar.append(match.group(1) if match else p_limpio)

            if not pano_nombres_a_buscar: continue

            # ==============================================================
            # COSECHA EN SERIE CON CONTROL DE FLUJO SEGURO
            # ==============================================================
            for pano_nombre in set(pano_nombres_a_buscar):
                if not pano_nombre: continue
                
                endpoints_series = ['interruptores', 'desconectadores', 'transformadores_corriente', 'trampas_ondas']
                sub_equipos_encontrados = []

                for tipo in endpoints_series:
                    url_search = f"{BASE_URLS[tipo]}/?search={pano_nombre}"
                    datos_api = await hacer_solicitud(session, url_search)
                    
                    if datos_api and isinstance(datos_api, list):
                        for item in datos_api:
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
                    max_equipos_detectados = max(max_equipos_detectados, len(sub_equipos_encontrados))
                    
                    limitante_corriente = min(sub_equipos_encontrados, key=lambda x: x['corriente'])
                    equipos_con_ruptura = [x for x in sub_equipos_encontrados if x['ruptura'] != float('inf')]
                    limitante_ruptura = min(equipos_con_ruptura, key=lambda x: x['ruptura']) if equipos_con_ruptura else None

                    filas_reporte.append({
                        'id_consultado': eq_id,
                        'subestacion': subestacion,
                        'pano_nombre': pano_nombre,
                        'equipos': sub_equipos_encontrados,
                        'lim_corriente': f"{limitante_corriente['nombre']} (ID: {limitante_corriente['id']}) - {limitante_corriente['corriente']} A",
                        'lim_ruptura': f"{limitante_ruptura['nombre']} (ID: {limitante_ruptura['id']}) - {limitante_ruptura['ruptura']} kA" if limitante_ruptura else "N/A"
                    })

        if not filas_reporte:
            raise ValueError("No se encontraron elementos en serie para los IDs ingresados. Verifica el modo seleccionado (Directo para Equipos, Tramo para Líneas).")

        # ==============================================================
        # ESTRUCTURA HORIZONTAL EN COLUMNAS DEFINITIVA
        # ==============================================================
        headers = ['ID Consultado', 'Subestación / Elemento', 'Paño Coordinado']
        for i in range(1, max_equipos_detectados + 1):
            headers.extend([
                f'Equipo {i} ID', f'Equipo {i} Nombre', f'Equipo {i} Tipo', f'Equipo {i} Corr [A]', f'Equipo {i} Rup [kA]'
            ])
        headers.extend(['Elemento Limitante Corriente', 'Elemento Limitante Ruptura'])
        ws.append(headers)
        
        header_fill = PatternFill(start_color='1F4E78', end_color='1F4E78', fill_type='solid')
        header_font = Font(name='Arial', size=11, bold=True, color='FFFFFF')
        for col_num in range(1, len(headers) + 1):
            cell = ws.cell(row=1, column=col_num)
            cell.fill = header_fill
            cell.font = header_font
            cell.alignment = Alignment(horizontal='center', vertical='center', wrap_text=True)

        block_fill = PatternFill(start_color='FFF2CC', end_color='FFF2CC', fill_type='solid')
        
        for r_data in filas_reporte:
            fila_completa = [r_data['id_consultado'], r_data['subestacion'], r_data['pano_nombre']]
            
            for eq in r_data['equipos']:
                corr_display = eq['corriente'] if eq['corriente'] != float('inf') else 'N/A'
                rup_display = eq['ruptura'] if eq['ruptura'] != float('inf') else 'N/A'
                fila_completa.extend([eq['id'], eq['nombre'], eq['tipo'], corr_display, rup_display])
                
            celdas_faltantes = (max_equipos_detectados - len(r_data['equipos'])) * 5
            if celdas_faltantes > 0:
                fila_completa.extend([""] * celdas_faltantes)
                
            fila_completa.extend([r_data['lim_corriente'], r_data['lim_ruptura']])
            ws.append(fila_completa)
            
            row_idx = ws.max_row
            total_cols = len(headers)
            
            celda_m_corr = ws.cell(row=row_idx, column=total_cols - 1)
            celda_m_rup = ws.cell(row=row_idx, column=total_cols)
            
            for c_res in [celda_m_corr, celda_m_rup]:
                c_res.fill = block_fill
                c_res.font = Font(name='Arial', size=10, bold=True, color='C00000')
                
            for col_idx in range(1, total_cols + 1):
                ws.cell(row=row_idx, column=col_idx).alignment = Alignment(vertical='center', horizontal='center')

        for col in ws.columns:
            max_len = max(len(str(cell.value or '')) for cell in col)
            col_letter = col[0].column_letter
            ws.column_dimensions[col_letter].width = max(max_len + 3, 12)

        filepath = os.path.abspath(os.path.join(os.getcwd(), "equipos_datos.xlsx"))
        wb.save(filepath)
        return filepath
