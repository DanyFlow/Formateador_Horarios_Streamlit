"""Genera la entrada de un usuario para .streamlit/secrets.toml.

Uso:  python generar_hash.py
Pide usuario y contraseña (la contraseña no se muestra) e imprime el bloque
que hay que pegar en secrets.toml (local) o en los Secrets de Streamlit Cloud.
Repite el proceso por cada usuario.
"""

import getpass
import hashlib
import re
import secrets

ITERACIONES = 200_000  # debe coincidir con _ITERACIONES en FormateadorV3.py

usuario = input("Usuario (minúsculas, sin espacios): ").strip().lower()
if not re.fullmatch(r"[a-z0-9_-]+", usuario):
  raise SystemExit("Usa solo letras minúsculas, números, guion o guion bajo.")

password = getpass.getpass("Contraseña: ")
if getpass.getpass("Repite la contraseña: ") != password:
  raise SystemExit("Las contraseñas no coinciden.")
if len(password) < 10:
  raise SystemExit("Usa al menos 10 caracteres.")

salt = secrets.token_bytes(16)
digest = hashlib.pbkdf2_hmac("sha256", password.encode("utf-8"), salt, ITERACIONES)

print("\nPega este bloque en secrets.toml:\n")
print(f"[usuarios.{usuario}]")
print(f'salt = "{salt.hex()}"')
print(f'hash = "{digest.hex()}"')
