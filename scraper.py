import time
from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.chrome.service import Service

# Inicializar Selenium
options = webdriver.ChromeOptions()
# options.add_argument('--headless') # Descomenta si no quieres que se abra la ventana visualmente
driver = webdriver.Chrome(options=options)

try:
    print("Abriendo el portal de Mitacs Globalink...")
    # 1. Ir a la página de login o inicio de sesión
    driver.get("https://globalink.mitacs.ca/#/student/application/student-login")
    
    time.sleep(3) # Espera a que cargue la interfaz

    # 2. (Opcional) Localizar el botón de Login o campos de correo/contraseña
    # Ejemplo de estructura si Selenium necesita rellenar el usuario:
    # email_field = driver.find_element(By.ID, "id_del_input_correo")
    # email_field.send_keys("tu_correo@universidad.edu.co")
    
    print("Por seguridad con tus credenciales, puedes hacer el login manual en la ventana de Chrome que se abrió...")
    print("Una vez hayas iniciado sesión y estés viendo el buscador de proyectos, avísame en la terminal.")
    
    # Pausa amplia para que te loguees de forma manual la primera vez si lo prefieres
    time.sleep(30) 

    # 3. Navegar directamente a la URL de proyectos que mencionaste
    driver.get("https://globalink.mitacs.ca/#/student/application/projects")
    time.sleep(5)

    print("¡Listo en la sección de proyectos! Procederemos a extraer las tarjetas.")

finally:
    # driver.quit()
    pass