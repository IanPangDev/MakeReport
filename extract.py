from json import loads
from base64 import b64decode
from io import BytesIO
from PIL import Image
from playwright.sync_api import sync_playwright

def extract_image(cell) -> tuple[list[list], str]:
    """
    Extrae la imagen y el código fuente de una celda de un archivo Jupyter Notebook.

    @param cell: Celda del notebook que contiene la salida y el código fuente.

    @return: Una tupla con dos elementos:
        - Una imagen decodificada en formato PIL (si existe).
        - El código fuente de la celda.
    """
    # Si la celda contiene alguna salida (outputs)
    if len(cell['outputs']) != 0:
        # Extraemos la imagen que está codificada en base64
        img = Image.open(BytesIO(b64decode(cell['outputs'][0]['data']['image/png'])))
        # Regresamos la imagen decodificada y el código fuente de la celda
        return img, cell['source']
    
    # Si no hay imagen, solo regresamos el código fuente
    return None, cell['source']

def extract_all(path, file_name) -> None:
    """
    Extrae todas las imágenes y el código fuente de un archivo Jupyter Notebook.

    @param path: Ruta del archivo Jupyter Notebook (.ipynb).
    @param file_name: Nombre de la carpeta donde se guardarán las imágenes extraídas.

    @return: No devuelve nada. Guarda las imágenes extraídas y el código fuente procesado.
    """
    # Usamos Playwright para interactuar con la web
    with sync_playwright() as p:
        # Iniciamos el navegador Chromium en modo no headless (con interfaz gráfica)
        browser = p.chromium.launch(headless=False)
        context = browser.new_context()

        # Abrimos una nueva pestaña
        page = context.new_page()

        # Leemos el archivo .ipynb como un archivo JSON
        with open(path, mode="r", encoding="utf-8") as f:
            file = loads(f.read())
        
        # Recorremos todas las celdas del notebook (ignorando la primera celda que es de markdown con las librerías)
        for i, cell in enumerate(file['cells'][1:]):
            # Si la celda tiene salidas (imágenes o resultados de ejecución)
            if 'outputs' in cell.keys():
                # Extraemos la imagen y el código fuente de la celda
                img, code = extract_image(cell)
                
                # Si existe una imagen, la almacenamos en la carpeta indicada
                if img:
                    img.save(f'{file_name}/{i}img.png', 'PNG')

                # Ahora nos dirigimos a la página web para generar una imagen a partir del código fuente
                url = "https://carbon.now.sh/?bg=rgba%2887%2C136%2C178%2C0%29&t=monokai&wt=none&l=python&width=680&ds=true&dsyoff=20px&dsblur=68px&wc=true&wa=true&pv=56px&ph=56px&ln=false&fl=1&fm=Hack&fs=14px&lh=133%25&si=false&es=2x&wm=false&code=%250Adef%2520remove_image_out%28cell%29%253A%250A%2520%2520%2520%2520if%2520len%28cell%255B%27outputs%27%255D%29%2520%21%253D%25200%253A%250A%2520%2520%2520%2520%2520%2520%2520%2520cell%255B%27outputs%27%255D%255B0%255D%255B%27data%27%255D%255B%27image%252Fpng%27%255D%2520%253D%2520None%250A%2520%2520%2520%2520return%2520cell"
                page.goto(url)

                # Esperamos a que la página cargue y damos clic en el área para insertar el código
                page.wait_for_selector(
                    '#export-container > div > div.react-codemirror2.CodeMirror__container.window-theme__none > div > div.CodeMirror-scroll > div.CodeMirror-sizer > div',
                    timeout=1000
                )
                page.click(
                    '#export-container > div > div.react-codemirror2.CodeMirror__container.window-theme__none > div > div.CodeMirror-scroll > div.CodeMirror-sizer > div'
                )

                # Seleccionamos el texto existente y pegamos el código fuente extraído
                page.keyboard.press('Control+A')
                page.keyboard.type(''.join([l.lstrip() for l in code]))

                # Configuramos la descarga de la imagen generada del código fuente
                with page.expect_download() as download_info:
                    page.click(
                        '#__next > main > div.jsx-1200824569.page > div.jsx-2146564046.editor > div.jsx-e3415c9b8e46575.toolbar > div.jsx-2146564046.toolbar-second-row > div.jsx-2146564046.share-buttons > div.jsx-3153731481.export-menu-container > div > button.jsx-2184717013'
                    )
                    download = download_info.value
                    download.save_as(f"{file_name}/{i}code.png")
        
        # Cerramos el navegador una vez finalizado el proceso
        browser.close()