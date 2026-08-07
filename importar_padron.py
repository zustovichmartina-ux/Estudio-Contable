import pandas as pd
import re
import time
from ddgs import DDGS
from database import crear_cliente, inicializar_bd
import concurrent.futures
import threading

db_lock = threading.Lock()

def buscar_cuit(nombre):
    ddgs = DDGS()
    for attempt in range(2):
        try:
            results = list(ddgs.text(f"cuitonline {nombre}", max_results=2))
            for r in results:
                text = r.get('title', '') + " " + r.get('body', '')
                match = re.search(r'\b(20|23|24|27|30|33|34)-?\d{8}-?\d\b', text)
                if match:
                    return match.group(0).replace('-', '')
            return None
        except Exception as e:
            time.sleep(2)
    return None

def categorizar_tipo(nombre):
    nom = nombre.strip().upper()
    if re.search(r'\b(S\.?A\.?|S\.?R\.?L\.?|S\.A\.S\.?)$', nom) or "SOCIEDAD" in nom:
        return "Persona Jurídica"
    return "Persona Física"

def procesar_empresa(args):
    i, nom, total = args
    nom_busqueda = re.sub(r'[^\w\s]', '', nom)
    cuit = buscar_cuit(nom_busqueda)
    
    if not cuit:
        print(f"[{i}/{total}] [-] Sin CUIT: {nom}", flush=True)
        return "no_encontrado"
        
    tipo = categorizar_tipo(nom)
    
    with db_lock:
        try:
            crear_cliente(nom, cuit, tipo)
            print(f"[{i}/{total}] [+] Agregado: {nom} | CUIT: {cuit} ({tipo})", flush=True)
            return "agregado"
        except Exception as e:
            if "UNIQUE" in str(e):
                print(f"[{i}/{total}] [*] Ya existe: {nom} | CUIT: {cuit}", flush=True)
                return "ya_existente"
            else:
                print(f"[{i}/{total}] [!] Error DB: {nom} -> {e}", flush=True)
                return "error"

def main():
    print("Iniciando importación con concurrencia...")
    inicializar_bd()
    
    df = pd.read_excel('Lista_empresas.xlsx')
    nombres = df['Nombre'].dropna().astype(str).unique()
    total = len(nombres)
    
    agregados = 0
    ya_existentes = 0
    no_encontrados = 0
    
    print(f"Total a procesar: {total} empresas.")
    
    args_list = [(i+1, nom, total) for i, nom in enumerate(nombres)]
    
    with concurrent.futures.ThreadPoolExecutor(max_workers=5) as executor:
        for result in executor.map(procesar_empresa, args_list):
            if result == "agregado":
                agregados += 1
            elif result == "ya_existente":
                ya_existentes += 1
            elif result == "no_encontrado":
                no_encontrados += 1
                
    print("\n--- RESUMEN ---")
    print(f"Nuevos agregados: {agregados}")
    print(f"Ya existentes: {ya_existentes}")
    print(f"No encontrados/Error: {no_encontrados}")

if __name__ == '__main__':
    main()
