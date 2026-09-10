"""
Genera el hash de una contraseña nueva para AUTH_USERS (login real de Vercel).

La contraseña se escribe en un prompt oculto (no queda en el historial de la
terminal ni se envía a ningún lado) y el script solo imprime el hash — eso es
lo único que se pega en la variable de entorno AUTH_USERS en Vercel. La
contraseña en texto plano solo la sabe la persona que la escribió; guárdala
en un gestor de contraseñas (1Password/Bitwarden), no en este repo ni en chat.

Uso:
    python3 scripts/generar_hash_password.py
"""
import getpass
import hashlib
import os

SCRYPT_N = 16384
SCRYPT_R = 8
SCRYPT_P = 1
DKLEN = 64


def main():
    password = getpass.getpass("Contraseña nueva: ")
    confirm = getpass.getpass("Repite la contraseña: ")
    if password != confirm:
        print("Las contraseñas no coinciden.")
        return
    salt = os.urandom(16)
    derived = hashlib.scrypt(
        password.encode("utf-8"),
        salt=salt,
        n=SCRYPT_N,
        r=SCRYPT_R,
        p=SCRYPT_P,
        dklen=DKLEN,
    )
    print("\nHash (pegar en AUTH_USERS, nunca la contraseña en texto plano):")
    print(f"{salt.hex()}:{derived.hex()}")


if __name__ == "__main__":
    main()
