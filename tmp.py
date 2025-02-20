import os
import cv2
def resource_path(relative_path):
    """Retorna o caminho absoluto para recursos no pacote"""
    try:
        # PyInstaller cria uma pasta temporária para armazenar arquivos
        base_path = sys._MEIPASS
    except Exception:
        base_path = os.path.abspath(".")
    return os.path.join(base_path, relative_path)


print(resource_path('haarcascade_frontalface_default.xml'))