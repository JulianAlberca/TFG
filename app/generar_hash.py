from whitelist import calcular_hash_vba_code

ruta = r"C:\Users\alber\Desktop\TFG\source\subcarpeta\confianza2.xlsm"
hash = calcular_hash_vba_code(ruta)
print(f"Hash del documento: {hash}")
