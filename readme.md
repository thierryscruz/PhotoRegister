Comando para compilar
pyinstaller --onefile --windowed --icon=logo.ico PhotoReg.py --version-file version.txt
pyinstaller --onefile --windowed --icon=logo.ico --add-data "D:/Clientes/RAPIDONET/PhotoRegister/venv/Lib/site-packages/cv2/data/haarcascade_frontalface_default.xml;cv2/data" PhotoReg.py --version-file version.txt

