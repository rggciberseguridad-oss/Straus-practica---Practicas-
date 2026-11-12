#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
Script de automatización para generar y rellenar los partes mensuales de trabajo efectivo
del Alumno Rolando Gómez Gutiérrez en formato Word (.docx).
Proyecto: FPSE/2026/13/023 Ciberseguridad (IFCT0109)
"""

import os
import shutil
import zipfile
import argparse
import xml.etree.ElementTree as ET

W_NS = 'http://schemas.openxmlformats.org/wordprocessingml/2006/main'
W14_NS = 'http://schemas.microsoft.com/office/word/2010/wordml'

ET.register_namespace('w', W_NS)
ET.register_namespace('w14', W14_NS)
ET.register_namespace('r', 'http://schemas.openxmlformats.org/officeDocument/2006/relationships')
ET.register_namespace('m', 'http://schemas.openxmlformats.org/officeDocument/2006/math')
ET.register_namespace('v', 'urn:schemas-microsoft-com:vml')
ET.register_namespace('wp', 'http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing')
ET.register_namespace('w10', 'urn:schemas-microsoft-com:office:word')
ET.register_namespace('a', 'http://schemas.openxmlformats.org/drawingml/2006/main')
ET.register_namespace('pic', 'http://schemas.openxmlformats.org/drawingml/2006/picture')

# Preset de tareas verosímiles y ajustadas a MF0487_3 (Auditoría) para Noviembre 2026
TAREAS_NOVIEMBRE_2026 = [
    ("1ª", "02/11/2026", "Atención al público en tienda y recepción de equipos averiados en mostrador. Revisión de tickets de soporte en el CRM interno y planificación de los trabajos de auditoría y campo de la semana."),
    ("2ª", "03/11/2026", "Jornada de trabajo técnico en campo: desplazamiento e intervención en instalaciones de cliente para revisión física de cableado estructurado, verificación de puntos de acceso y comprobación de seguridad en periféricos de red."),
    ("3ª", "04/11/2026", "Soporte técnico presencial en tienda y taller. Auditoría post-instalación en el servidor HPE ProLiant ML30 Gen11 de Caminero & Prado Asociados SL: comprobación de la correcta ejecución de copias de seguridad automáticas en Windows Server Backup."),
    ("4ª", "05/11/2026", "Asistencia técnica remota (Supremo) para resolución de incidencias en clientes corporativos. Verificación del estado de los volúmenes lógicos en RAID 1 y comprobación de alertas en el controlador Smart Storage Administrator."),
    ("5ª", "06/11/2026", "Segunda jornada de trabajo técnico en campo: despliegue de equipamiento informático, comprobación de tomas de datos, etiquetado de red y validación de conectividad segura de puestos de trabajo con el cliente."),
    ("6ª", "09/11/2026", "Recepción de portátiles para mantenimiento preventivo en taller. Revisión de la base de datos del CRM propio, actualización de fichas relacionales de clientes y archivo de albaranes de servicio de campo."),
    ("7ª", "10/11/2026", "Soporte Helpdesk de primer nivel a usuarios finales. Auditoría de cuentas y directivas de seguridad en Windows Server: análisis de eventos de inicio de sesión fallidos en el Visor de Eventos (Event Viewer) para detectar posibles intentos de intrusión."),
    ("8ª", "11/11/2026", "Atención al público en mostrador asesorando sobre sustitución de componentes y periféricos. Escaneo de vulnerabilidades en equipos locales del taller utilizando herramientas de análisis de puertos y servicios activos."),
    ("9ª", "12/11/2026", "Diagnóstico de hardware en banco de trabajo (revisión de fuentes de alimentación y pruebas de integridad de memoria RAM). Soporte remoto para solventar problemas de sincronización de perfiles en clientes de oficina."),
    ("10ª", "13/11/2026", "Soporte presencial en tienda. Auditoría de seguridad lógica en sistemas operativos: verificación del estado de actualización de definiciones antivirus y comprobación de que el cortafuegos de Windows se encuentra habilitado en todos los perfiles de red."),
    ("11ª", "16/11/2026", "Atención en mostrador y gestión de garantías con mayoristas informáticos. Revisión del árbol de carpetas compartidas y auditoría de permisos efectivos NTFS en el servidor para asegurar el cumplimiento del principio de mínimo privilegio."),
    ("12ª", "17/11/2026", "Asistencia técnica remota con Supremo. Comprobación y verificación de la integridad del firewall perimetral y revisión de registros de acceso web en el entorno administrativo de clientes municipales."),
    ("13ª", "18/11/2026", "Recepción de equipos en taller para desinfección de software malicioso. Diagnóstico de persistencia de amenazas en el registro de Windows y saneamiento de navegadores y tareas programadas no autorizadas."),
    ("14ª", "19/11/2026", "Soporte técnico en mostrador resolviendo consultas operativas de usuarios. Comprobación del estado de las licencias de acceso de clientes (CAL) y verificación de la correcta asignación de perfiles estándar locales."),
    ("15ª", "20/11/2026", "Alcanzado el hito del 50% del proyecto formativo según cronograma. En trabajo efectivo: auditoría global de rendimiento en el servidor del taller, revisión de espacio libre en discos de copia de seguridad y optimización de servicios en segundo plano."),
    ("16ª", "23/11/2026", "Atención general al público en tienda. Elaboración de lista de verificación (checklist) de auditoría técnica de seguridad para evaluar vulnerabilidades comunes en puestos de trabajo ofimáticos."),
    ("17ª", "24/11/2026", "Soporte Helpdesk presencial. Ejecución de pruebas de restauración selectiva a partir de copias de seguridad de Windows Server Backup para verificar la fiabilidad y consistencia del plan de recuperación ante desastres."),
    ("18ª", "25/11/2026", "Recepción de equipos para ampliación de discos de estado sólido (SSD). Alta y actualización de registros técnicos en el CRM propio de la empresa y emisión de presupuestos de reparación."),
    ("19ª", "26/11/2026", "Asistencia remota a empresas clientes solucionando incidencias de conectividad de red local y configuración de recursos compartidos. Comprobación de directivas de bloqueo de cuenta por intentos erróneos de contraseña."),
    ("20ª", "27/11/2026", "Soporte en taller y banco de pruebas. Redacción del borrador del informe de auditoría técnica con la recopilación de hallazgos de seguridad y propuestas de mejora para los equipos evaluados durante el mes."),
    ("21ª", "30/11/2026", "Organización y limpieza del área de taller técnico. Ejecución de la copia de seguridad mensual de la base de datos del CRM propio, cierre administrativo del período formativo de noviembre y preparación de la evaluación del módulo MF0487_3.")
]

def make_task_cell(date_str, desc_str):
    tc = ET.Element(f'{{{W_NS}}}tc')
    tcPr = ET.SubElement(tc, f'{{{W_NS}}}tcPr')
    ET.SubElement(tcPr, f'{{{W_NS}}}tcW', {f'{{{W_NS}}}w': '8080', f'{{{W_NS}}}type': 'dxa'})
    borders = ET.SubElement(tcPr, f'{{{W_NS}}}tcBorders')
    for b in ['top', 'left', 'bottom', 'right']:
        ET.SubElement(borders, f'{{{W_NS}}}{b}', {
            f'{{{W_NS}}}val': 'single', f'{{{W_NS}}}sz': '4', f'{{{W_NS}}}space': '0', f'{{{W_NS}}}color': '000080'
        })

    # P1: Date
    p1 = ET.SubElement(tc, f'{{{W_NS}}}p')
    pPr1 = ET.SubElement(p1, f'{{{W_NS}}}pPr')
    ET.SubElement(pPr1, f'{{{W_NS}}}jc', {f'{{{W_NS}}}val': 'both'})
    rPr1_p = ET.SubElement(pPr1, f'{{{W_NS}}}rPr')
    ET.SubElement(rPr1_p, f'{{{W_NS}}}color', {f'{{{W_NS}}}val': '808080'})

    r1 = ET.SubElement(p1, f'{{{W_NS}}}r')
    rPr1 = ET.SubElement(r1, f'{{{W_NS}}}rPr')
    ET.SubElement(rPr1, f'{{{W_NS}}}color', {f'{{{W_NS}}}val': '808080'})
    t1 = ET.SubElement(r1, f'{{{W_NS}}}t')
    t1.text = date_str

    # P2: Description
    p2 = ET.SubElement(tc, f'{{{W_NS}}}p')
    pPr2 = ET.SubElement(p2, f'{{{W_NS}}}pPr')
    ET.SubElement(pPr2, f'{{{W_NS}}}jc', {f'{{{W_NS}}}val': 'both'})
    rPr2_p = ET.SubElement(pPr2, f'{{{W_NS}}}rPr')
    ET.SubElement(rPr2_p, f'{{{W_NS}}}color', {f'{{{W_NS}}}val': '808080'})

    r2 = ET.SubElement(p2, f'{{{W_NS}}}r')
    rPr2 = ET.SubElement(r2, f'{{{W_NS}}}rPr')
    ET.SubElement(rPr2, f'{{{W_NS}}}color', {f'{{{W_NS}}}val': '808080'})
    t2 = ET.SubElement(r2, f'{{{W_NS}}}t', {'{http://www.w3.org/XML/1998/namespace}space': 'preserve'})
    t2.text = desc_str

    # P3: Empty trailing paragraph
    p3 = ET.SubElement(tc, f'{{{W_NS}}}p')
    pPr3 = ET.SubElement(p3, f'{{{W_NS}}}pPr')
    ET.SubElement(pPr3, f'{{{W_NS}}}jc', {f'{{{W_NS}}}val': 'both'})
    rPr3_p = ET.SubElement(pPr3, f'{{{W_NS}}}rPr')
    ET.SubElement(rPr3_p, f'{{{W_NS}}}color', {f'{{{W_NS}}}val': '808080'})

    return tc

def make_obs_cell():
    tc = ET.Element(f'{{{W_NS}}}tc')
    tcPr = ET.SubElement(tc, f'{{{W_NS}}}tcPr')
    ET.SubElement(tcPr, f'{{{W_NS}}}tcW', {f'{{{W_NS}}}w': '8080', f'{{{W_NS}}}type': 'dxa'})
    borders = ET.SubElement(tcPr, f'{{{W_NS}}}tcBorders')
    for b in ['top', 'left', 'bottom', 'right']:
        ET.SubElement(borders, f'{{{W_NS}}}{b}', {
            f'{{{W_NS}}}val': 'single', f'{{{W_NS}}}sz': '4', f'{{{W_NS}}}space': '0', f'{{{W_NS}}}color': '000080'
        })
    ET.SubElement(tcPr, f'{{{W_NS}}}vAlign', {f'{{{W_NS}}}val': 'center'})

    p = ET.SubElement(tc, f'{{{W_NS}}}p')
    pPr = ET.SubElement(p, f'{{{W_NS}}}pPr')
    ET.SubElement(pPr, f'{{{W_NS}}}jc', {f'{{{W_NS}}}val': 'both'})
    rPr = ET.SubElement(pPr, f'{{{W_NS}}}rPr')
    ET.SubElement(rPr, f'{{{W_NS}}}rFonts', {
        f'{{{W_NS}}}ascii': 'Arial', f'{{{W_NS}}}cs': 'Arial', f'{{{W_NS}}}eastAsia': 'Arial', f'{{{W_NS}}}hAnsi': 'Arial'
    })
    ET.SubElement(rPr, f'{{{W_NS}}}b', {f'{{{W_NS}}}val': '1'})
    ET.SubElement(rPr, f'{{{W_NS}}}bCs', {f'{{{W_NS}}}val': '1'})
    ET.SubElement(rPr, f'{{{W_NS}}}color', {f'{{{W_NS}}}val': '000080'})
    ET.SubElement(rPr, f'{{{W_NS}}}sz', {f'{{{W_NS}}}val': '20'})
    ET.SubElement(rPr, f'{{{W_NS}}}szCs', {f'{{{W_NS}}}val': '20'})

    r1 = ET.SubElement(p, f'{{{W_NS}}}r')
    rPr1 = ET.SubElement(r1, f'{{{W_NS}}}rPr')
    ET.SubElement(rPr1, f'{{{W_NS}}}rFonts', {
        f'{{{W_NS}}}ascii': 'Arial', f'{{{W_NS}}}cs': 'Arial', f'{{{W_NS}}}eastAsia': 'Arial', f'{{{W_NS}}}hAnsi': 'Arial'
    })
    ET.SubElement(rPr1, f'{{{W_NS}}}b', {f'{{{W_NS}}}val': '1'})
    ET.SubElement(rPr1, f'{{{W_NS}}}bCs', {f'{{{W_NS}}}val': '1'})
    ET.SubElement(rPr1, f'{{{W_NS}}}color', {f'{{{W_NS}}}val': '000080'})
    ET.SubElement(rPr1, f'{{{W_NS}}}sz', {f'{{{W_NS}}}val': '20'})
    ET.SubElement(rPr1, f'{{{W_NS}}}szCs', {f'{{{W_NS}}}val': '20'})
    t1 = ET.SubElement(r1, f'{{{W_NS}}}t', {'{http://www.w3.org/XML/1998/namespace}space': 'preserve'})
    t1.text = "Observaciones: "

    r2 = ET.SubElement(p, f'{{{W_NS}}}r')
    rPr2 = ET.SubElement(r2, f'{{{W_NS}}}rPr')
    ET.SubElement(rPr2, f'{{{W_NS}}}color', {f'{{{W_NS}}}val': '808080'})
    t2 = ET.SubElement(r2, f'{{{W_NS}}}t')
    t2.text = "…"

    return tc

def make_day_label_cell(label_str):
    tc = ET.Element(f'{{{W_NS}}}tc')
    tcPr = ET.SubElement(tc, f'{{{W_NS}}}tcPr')
    ET.SubElement(tcPr, f'{{{W_NS}}}tcW', {f'{{{W_NS}}}w': '1556', f'{{{W_NS}}}type': 'dxa'})
    borders = ET.SubElement(tcPr, f'{{{W_NS}}}tcBorders')
    for b in ['top', 'left', 'bottom', 'right']:
        ET.SubElement(borders, f'{{{W_NS}}}{b}', {
            f'{{{W_NS}}}val': 'single', f'{{{W_NS}}}sz': '4', f'{{{W_NS}}}space': '0', f'{{{W_NS}}}color': '000080'
        })
    ET.SubElement(tcPr, f'{{{W_NS}}}vAlign', {f'{{{W_NS}}}val': 'center'})

    p = ET.SubElement(tc, f'{{{W_NS}}}p')
    pPr = ET.SubElement(p, f'{{{W_NS}}}pPr')
    ET.SubElement(pPr, f'{{{W_NS}}}jc', {f'{{{W_NS}}}val': 'center'})
    rPr = ET.SubElement(pPr, f'{{{W_NS}}}rPr')
    ET.SubElement(rPr, f'{{{W_NS}}}color', {f'{{{W_NS}}}val': '808080'})

    r = ET.SubElement(p, f'{{{W_NS}}}r')
    rPr_r = ET.SubElement(r, f'{{{W_NS}}}rPr')
    ET.SubElement(rPr_r, f'{{{W_NS}}}color', {f'{{{W_NS}}}val': '808080'})
    t = ET.SubElement(r, f'{{{W_NS}}}t')
    t.text = label_str

    return tc

def make_empty_label_cell():
    tc = ET.Element(f'{{{W_NS}}}tc')
    tcPr = ET.SubElement(tc, f'{{{W_NS}}}tcPr')
    ET.SubElement(tcPr, f'{{{W_NS}}}tcW', {f'{{{W_NS}}}w': '1556', f'{{{W_NS}}}type': 'dxa'})
    borders = ET.SubElement(tcPr, f'{{{W_NS}}}tcBorders')
    for b in ['top', 'left', 'bottom', 'right']:
        ET.SubElement(borders, f'{{{W_NS}}}{b}', {
            f'{{{W_NS}}}val': 'single', f'{{{W_NS}}}sz': '4', f'{{{W_NS}}}space': '0', f'{{{W_NS}}}color': '000080'
        })
    ET.SubElement(tcPr, f'{{{W_NS}}}vAlign', {f'{{{W_NS}}}val': 'center'})
    ET.SubElement(tc, f'{{{W_NS}}}p')
    return tc

def generar_parte(doc_template_path, doc_output_path, mes_num, primer_dia, ultimo_dia, lista_tareas):
    if not os.path.exists(doc_template_path):
        raise FileNotFoundError(f"Plantilla no encontrada: {doc_template_path}")

    # Backup si el archivo de salida ya existía
    if os.path.exists(doc_output_path):
        bak = doc_output_path + ".bak"
        shutil.copy2(doc_output_path, bak)
        print(f"Copia de seguridad creada en: {bak}")

    with zipfile.ZipFile(doc_template_path, 'r') as z:
        all_files = {name: z.read(name) for name in z.namelist()}

    root = ET.fromstring(all_files['word/document.xml'])
    ns = {'w': W_NS}
    table = root.find('.//w:tbl', ns)
    rows = table.findall('./w:tr', ns)

    # Actualizar encabezado (Row 0)
    r0 = rows[0]
    t_nodes = r0.findall('.//w:t', ns)
    for t in t_nodes:
        if 'Mes ' in t.text or 'Mes' == t.text.strip():
            pass

    # Reconstrucción limpia de la fila 0 para mes_num y fechas
    # Eliminar párrafos viejos de r0 y agregar el formateado
    c0_hdr = r0.findall('./w:tc', ns)[0]
    for p_child in list(c0_hdr.findall('./w:p', ns)):
        c0_hdr.remove(p_child)

    p_hdr = ET.SubElement(c0_hdr, f'{{{W_NS}}}p')
    pPr_h = ET.SubElement(p_hdr, f'{{{W_NS}}}pPr')
    ET.SubElement(pPr_h, f'{{{W_NS}}}jc', {f'{{{W_NS}}}val': 'both'})
    rPr_h = ET.SubElement(pPr_h, f'{{{W_NS}}}rPr')
    ET.SubElement(rPr_h, f'{{{W_NS}}}rFonts', {
        f'{{{W_NS}}}ascii': 'Arial', f'{{{W_NS}}}cs': 'Arial', f'{{{W_NS}}}eastAsia': 'Arial', f'{{{W_NS}}}hAnsi': 'Arial'
    })
    ET.SubElement(rPr_h, f'{{{W_NS}}}b', {f'{{{W_NS}}}val': '1'})
    ET.SubElement(rPr_h, f'{{{W_NS}}}color', {f'{{{W_NS}}}val': '000080'})
    ET.SubElement(rPr_h, f'{{{W_NS}}}sz', {f'{{{W_NS}}}val': '20'})

    r_h = ET.SubElement(p_hdr, f'{{{W_NS}}}r')
    rPr_r_h = ET.SubElement(r_h, f'{{{W_NS}}}rPr')
    ET.SubElement(rPr_r_h, f'{{{W_NS}}}rFonts', {
        f'{{{W_NS}}}ascii': 'Arial', f'{{{W_NS}}}cs': 'Arial', f'{{{W_NS}}}eastAsia': 'Arial', f'{{{W_NS}}}hAnsi': 'Arial'
    })
    ET.SubElement(rPr_r_h, f'{{{W_NS}}}b', {f'{{{W_NS}}}val': '1'})
    ET.SubElement(rPr_r_h, f'{{{W_NS}}}color', {f'{{{W_NS}}}val': '000080'})
    ET.SubElement(rPr_r_h, f'{{{W_NS}}}sz', {f'{{{W_NS}}}val': '20'})
    t_h = ET.SubElement(r_h, f'{{{W_NS}}}t', {'{http://www.w3.org/XML/1998/namespace}space': 'preserve'})
    t_h.text = f"Mes {mes_num}:  \u2002{primer_dia}\u2002 al \u2002\u2002{ultimo_dia}\u2002\u2002\u2002\u2002 (indicar el primer y último día trabajado)"

    # Limpiar filas previas de días (conservar Row 0 y Row 1)
    while len(table.findall('./w:tr', ns)) > 2:
        table.remove(table.findall('./w:tr', ns)[-1])

    # Inyectar jornadas
    for day_lbl, date_val, desc_val in lista_tareas:
        tr_task = ET.Element(f'{{{W_NS}}}tr')
        trPr1 = ET.SubElement(tr_task, f'{{{W_NS}}}trPr')
        ET.SubElement(trPr1, f'{{{W_NS}}}cantSplit')
        tr_task.append(make_day_label_cell(day_lbl))
        tr_task.append(make_task_cell(date_val, desc_val))
        table.append(tr_task)

        tr_obs = ET.Element(f'{{{W_NS}}}tr')
        trPr2 = ET.SubElement(tr_obs, f'{{{W_NS}}}trPr')
        ET.SubElement(trPr2, f'{{{W_NS}}}cantSplit')
        tr_obs.append(make_empty_label_cell())
        tr_obs.append(make_obs_cell())
        table.append(tr_obs)

    os.makedirs(os.path.dirname(os.path.abspath(doc_output_path)), exist_ok=True)
    all_files['word/document.xml'] = ET.tostring(root, encoding='utf-8', xml_declaration=True)

    with zipfile.ZipFile(doc_output_path, 'w', zipfile.ZIP_DEFLATED) as z_out:
        for name, content in all_files.items():
            z_out.writestr(name, content)

    print(f"Documento generado exitosamente en: {doc_output_path} ({len(lista_tareas)} jornadas procesadas)")

if __name__ == '__main__':
    parser = argparse.ArgumentParser(description="Generador de partes mensuales de trabajo efectivo")
    parser.add_argument('--mes', type=int, default=11, help="Número de mes (por defecto: 11 para Noviembre)")
    args = parser.parse_args()

    if args.mes == 11:
        # Noviembre 2026 (Mes 5)
        tpl = r"10-26/Rolando Gomez Gutierrez - 023-CT1-Alumno-Octubre.docx"
        out = r"11-26/Rolando Gomez Gutierrez - 023-CT1-Alumno-Noviembre.docx"
        generar_parte(tpl, out, mes_num=5, primer_dia="02/11/2026", ultimo_dia="30/11/2026", lista_tareas=TAREAS_NOVIEMBRE_2026)
    else:
        print(f"Mes {args.mes} no implementado con preset directo. Personalizar o ampliar.")
