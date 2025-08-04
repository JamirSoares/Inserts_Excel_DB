import time
import sys
import os
from watchdog.observers import Observer
from watchdog.events import FileSystemEventHandler

# Ajuste o path para conseguir importar seu módulo
sys.path.append(os.path.abspath(os.path.join(os.path.dirname(__file__), '..')))
from ProcessarExcel import processar_excel  # importa a função correta

class MyEventHandler(FileSystemEventHandler):
    def on_created(self, event):
        if not event.is_directory:
            print(f"Arquivo criado: {event.src_path}, iniciando processamento...")
            processar_excel()

if __name__ == "__main__":
    path = r"C:/Users/jamir.rodrigues/Documents/processar"
    event_handler = MyEventHandler()
    observer = Observer()
    observer.schedule(event_handler, path, recursive=False)  # geralmente não precisa ser recursivo
    observer.start()

    print(f"Monitorando a pasta {path} para novos arquivos...")

    try:
        while True:
            time.sleep(1)
    except KeyboardInterrupt:
        observer.stop()
    observer.join()
