import time
import keyboard
import pandas as pd
from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.common.keys import Keys
from openpyxl.styles import Alignment, PatternFill

KEYWORDS_TO_SEARCH = [
    "Machine Learning",
    "Artificial Intelligence",
    "Java",
    "Software Engineering",
    "Database",
    "Data"
]

TARGET_CITIES = [
    "Montreal",   # Québec (Azul Claro Pastel)
    "Calgary",    # Alberta (Verde Claro Pastel)
    "Edmonton",   # Alberta
    "Ottawa",     # Ontario (Rosa Claro Pastel)
    "Kingston",   # Ontario
    "London"      # Ontario
]

MAX_PER_CITY_PER_KEYWORD = 10

options = webdriver.ChromeOptions()
driver = webdriver.Chrome(options=options)

try:
    print("Abriendo el portal de Mitacs Globalink...")
    driver.get("https://globalink.mitacs.ca/#/student/application/student-login")
    
    time.sleep(3)
    print("Por favor realiza el login manual en la ventana de Chrome...")
    print("Esperando 15 segundos para completar el login...")
    time.sleep(15) 

    # 1. Navegar al catálogo de proyectos
    print("Navegando al catálogo de proyectos...")
    driver.get("https://globalink.mitacs.ca/#/student/application/projects")
    time.sleep(8) 

    # Diccionario para agrupar los datos por Palabra Clave (pestaña en Excel)
    projects_by_keyword = {}
    seen_ids = set()

    # -------------------------------------------------------------
    # BUCLE DE BÚSQUEDA POR PALABRAS CLAVE
    # -------------------------------------------------------------
    for kw in KEYWORDS_TO_SEARCH:
        print(f"\n=======================================================")
        print(f"---> BUSCANDO PARA KEYWORD: '{kw}' (MÁX {MAX_PER_CITY_PER_KEYWORD} POR CIUDAD)")
        print(f"=======================================================")
        print("💡 CONSEJO: Presiona la 'FLECHA DERECHA (->)' en cualquier momento para saltar al siguiente tema.")

        current_kw_projects = []

        try:
            # A. Escribir Keyword en Keyword search
            search_box = driver.find_element(By.CSS_SELECTOR, "input.p-inputtext")
            search_box.clear()
            search_box.send_keys(kw)
            time.sleep(1)

            # B. Seleccionar Idioma 'English'
            try:
                lang_dropdowns = driver.find_elements(By.XPATH, "//div[contains(., 'Language') and .//p-dropdown]")
                if lang_dropdowns:
                    lang_dd = lang_dropdowns[0].find_element(By.TAG_NAME, "p-dropdown")
                    lang_dd.click()
                    time.sleep(1)
                    
                    option_eng = driver.find_element(By.XPATH, "//li[contains(@class, 'p-dropdown-item') and contains(., 'English')]")
                    option_eng.click()
                    time.sleep(1)
            except Exception as e_lang:
                print(f"Nota: Ajustando filtro de Idioma a English: {e_lang}")

            # C. Seleccionar 'Computer Science' en la disciplina y trasfondo académico
            try:
                dropdowns = driver.find_elements(By.CSS_SELECTOR, "p-dropdown")
                for dropdown in dropdowns[-2:]:
                    dropdown.click()
                    time.sleep(1)
                    
                    filter_input = driver.find_element(By.CSS_SELECTOR, "input.p-dropdown-filter")
                    filter_input.clear()
                    filter_input.send_keys("Computer Science")
                    time.sleep(1)

                    option_cs = driver.find_element(By.XPATH, "//li[contains(@class, 'p-dropdown-item') and contains(., 'Computer Science')]")
                    option_cs.click()
                    time.sleep(1)
            except Exception:
                pass

            # D. Enviar búsqueda
            try:
                btn_search = driver.find_element(By.XPATH, "//button[contains(., 'Search and Filter')]")
                btn_search.click()
            except Exception:
                search_box.send_keys(Keys.ENTER)

            time.sleep(6) 

        except Exception as e:
            print(f"Error aplicando la keyword '{kw}': {e}")
            continue

        city_counts = {city.lower(): 0 for city in TARGET_CITIES}
        page_num = 1
        previous_active_page = ""
        no_new_data_counter = 0

        while True:
            # 1. Detección manual con flecha derecha
            if keyboard.is_pressed('right'):
                print(f"\n[FLECHA DERECHA DETECTADA] Saltando '{kw}' a petición tuya...")
                time.sleep(1)
                break

            # 2. Parada si 3 ciudades alcanzaron el límite de 10 proyectos
            cities_at_max = sum(1 for count in city_counts.values() if count >= MAX_PER_CITY_PER_KEYWORD)
            if cities_at_max >= 3:
                print(f"¡Se alcanzaron 10 proyectos en al menos 3 ciudades! Pasando a la siguiente keyword...")
                break

            # 3. VERIFICACIÓN 1: Leer el número de página resaltado en el Paginator de PrimeNG
            try:
                active_page_elem = driver.find_element(By.CSS_SELECTOR, ".p-paginator-page.p-highlight, .p-paginator-page.p-state-active")
                current_active_page = active_page_elem.text.strip()
            except Exception:
                current_active_page = str(page_num)

            # Si le dimos clic a 'Siguiente' pero el número de página en la web NO cambió
            if current_active_page == previous_active_page and page_num > 1:
                print(f"Fin real alcanzado automáticamente en la página {current_active_page} para '{kw}'. Se pasa al siguiente tema.")
                break
            
            previous_active_page = current_active_page

            print(f"-> Analizando página real {current_active_page} para '{kw}'... Conteo por ciudad: {city_counts}")

            cards = driver.find_elements(By.XPATH, "//div[contains(., 'Project ID') and .//button[contains(text(), 'View Detail')]]")
            if not cards:
                cards = driver.find_elements(By.XPATH, "//*[contains(text(), 'Project ID')]/ancestor::div[2]")

            new_projects_in_this_page = 0

            for card in cards:
                try:
                    raw_text = card.text.strip()
                    lines = [line.strip() for line in raw_text.split('\n') if line.strip()]

                    if raw_text.count("Project ID") > 1:
                        continue

                    project_id = ""
                    title = ""
                    university = ""
                    campus = ""
                    province = ""
                    city = ""
                    language = ""
                    description = ""

                    for i, line in enumerate(lines):
                        if "Project ID" in line:
                            project_id = line.replace("Project ID", "").strip()
                            if i + 1 < len(lines):
                                title = lines[i + 1]
                        elif "Faculty University:" in line:
                            university = line.replace("Faculty University:", "").strip()
                        elif "Faculty Campus:" in line:
                            campus = line.replace("Faculty Campus:", "").strip()
                        elif "Faculty Province:" in line:
                            province = line.replace("Faculty Province:", "").strip()
                        elif "Project Location:" in line:
                            city = line.replace("Project Location:", "").strip()
                        elif "Language:" in line:
                            language = line.replace("Language:", "").strip()

                    # Validar idioma inglés por seguridad
                    if language and "english" not in language.lower():
                        continue

                    candidates = [l for l in lines if len(l) > 70 and not any(k in l for k in ["Faculty", "Project Location", "Language", "Preferred start", "Project ID"])]
                    if candidates:
                        description = candidates[0]

                    matched_city = None
                    for target in TARGET_CITIES:
                        if target.lower() in city.lower() or target.lower() in campus.lower():
                            matched_city = target.lower()
                            break

                    if project_id and project_id not in seen_ids and matched_city:
                        if city_counts[matched_city] < MAX_PER_CITY_PER_KEYWORD:
                            seen_ids.add(project_id)
                            city_counts[matched_city] += 1
                            new_projects_in_this_page += 1
                            
                            current_kw_projects.append({
                                "Keyword Usada": kw,
                                "ID Proyecto": project_id,
                                "Título": title,
                                "Universidad": university,
                                "Campus": campus,
                                "Provincia": province,
                                "Ciudad": city,
                                "Idioma": language,
                                "Descripción": description
                            })
                            print(f"  + Proyecto [{project_id}] registrado en {city} ({city_counts[matched_city]}/{MAX_PER_CITY_PER_KEYWORD} para '{kw}')")

                except Exception:
                    continue

            # VERIFICACIÓN 2: Si no hubo proyectos nuevos en 2 páginas consecutivas
            if new_projects_in_this_page == 0:
                no_new_data_counter += 1
                if no_new_data_counter >= 2 and page_num > 5:
                    print(f"No se detectaron proyectos nuevos en 2 páginas consecutivas. Fin de paginación para '{kw}'.")
                    break
            else:
                no_new_data_counter = 0

            # VERIFICACIÓN 3: Avanzar a la siguiente página o romper si está deshabilitado
            try:
                next_btn = driver.find_element(By.XPATH, "//button[contains(@class, 'p-paginator-next')]")
                btn_class = next_btn.get_attribute("class") or ""
                is_disabled = next_btn.get_attribute("disabled") or next_btn.get_attribute("aria-disabled")

                if "p-disabled" in btn_class or is_disabled == "true" or is_disabled is not None or not next_btn.is_enabled():
                    print(f"El botón de 'Siguiente' está deshabilitado. Fin de paginación para '{kw}'.")
                    break
                else:
                    next_btn.click()
                    page_num += 1
                    time.sleep(4)
            except Exception:
                print(f"No hay más botones de paginación para '{kw}'.")
                break

        if current_kw_projects:
            projects_by_keyword[kw] = current_kw_projects

    # -------------------------------------------------------------
    # EXPORTAR A EXCEL CON PESTAÑAS Y COLORES PASTELES POR PROVINCIA
    # -------------------------------------------------------------
    if projects_by_keyword:
        excel_filename = "mitacs_proyectos_ciudades_pastel.xlsx"
        
        with pd.ExcelWriter(excel_filename, engine='openpyxl') as writer:
            
            fill_quebec = PatternFill(start_color="D0E0E3", end_color="D0E0E3", fill_type="solid")   # Azul Claro
            fill_alberta = PatternFill(start_color="E2EFDA", end_color="E2EFDA", fill_type="solid")  # Verde Claro
            fill_ontario = PatternFill(start_color="F4CCCC", end_color="F4CCCC", fill_type="solid")  # Rosa Claro
            fill_other = PatternFill(start_color="F3F3F3", end_color="F3F3F3", fill_type="solid")
            wrap_alignment = Alignment(wrap_text=True, vertical='top')

            col_widths = {
                'A': 20,  # Keyword Usada
                'B': 12,  # ID
                'C': 30,  # Título
                'D': 25,  # Universidad
                'E': 18,  # Campus
                'F': 15,  # Provincia
                'G': 20,  # Ciudad
                'H': 12,  # Idioma
                'I': 65   # Descripción
            }

            total_projects_count = 0

            for kw_name, data_list in projects_by_keyword.items():
                df = pd.DataFrame(data_list)
                df.sort_values(by=["Provincia", "Ciudad"], inplace=True)
                
                sheet_title = kw_name[:30].replace(":", "").replace("/", "")
                df.to_excel(writer, index=False, sheet_name=sheet_title)
                
                worksheet = writer.sheets[sheet_title]

                for col_letter, width in col_widths.items():
                    worksheet.column_dimensions[col_letter].width = width

                for row in worksheet.iter_rows(min_row=2, max_col=9, max_row=len(data_list)+1):
                    prov_val = str(row[5].value).lower() if row[5].value else ""
                    
                    if "qu&eacute;bec" in prov_val or "quebec" in prov_val:
                        current_fill = fill_quebec
                    elif "alberta" in prov_val:
                        current_fill = fill_alberta
                    elif "ontario" in prov_val:
                        current_fill = fill_ontario
                    else:
                        current_fill = fill_other

                    for cell in row:
                        cell.alignment = wrap_alignment
                        cell.fill = current_fill

                total_projects_count += len(data_list)

        print(f"\n¡ÉXITO TOTAL! Se guardaron {total_projects_count} proyectos distribuidos en {len(projects_by_keyword)} pestañas (Sheets) en '{excel_filename}'.")
    else:
        print("No se encontraron proyectos para las ciudades seleccionadas.")

finally:
    # driver.quit()
    pass