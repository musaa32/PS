"""
AES-256-GCM File Encryption/Decryption
Uses: cryptography library (pip install cryptography)

Encryption scheme:
  - AES-256-GCM (authenticated encryption — detects tampering)
  - Random 256-bit key derived via PBKDF2-HMAC-SHA256 from a password
  - Random 128-bit salt (stored in output file)
  - Random 96-bit nonce (stored in output file)

Output file layout (binary):
  [4 bytes magic] [16 bytes salt] [12 bytes nonce] [16 bytes GCM tag] [ciphertext]
"""

import os
import sys
import argparse
import getpass
from cryptography.hazmat.primitives.ciphers.aead import AESGCM
from cryptography.hazmat.primitives.kdf.pbkdf2 import PBKDF2HMAC
from cryptography.hazmat.primitives import hashes

# --- Constants ---
MAGIC = b"AESG"          # 4-byte file signature
SALT_LEN = 16            # bytes
NONCE_LEN = 12           # bytes (96-bit, recommended for GCM)
TAG_LEN = 16             # bytes (GCM authentication tag, appended by library)
KDF_ITERATIONS = 600_000 # OWASP recommended minimum for PBKDF2-SHA256


def derive_key(password: str, salt: bytes) -> bytes:
    """Derive a 256-bit AES key from a password using PBKDF2-HMAC-SHA256."""
    kdf = PBKDF2HMAC(
        algorithm=hashes.SHA256(),
        length=32,
        salt=salt,
        iterations=KDF_ITERATIONS,
    )
    return kdf.derive(password.encode())


def encrypt_file(input_path: str, output_path: str, password: str) -> None:
    """Encrypt a file with AES-256-GCM."""
    salt = os.urandom(SALT_LEN)
    nonce = os.urandom(NONCE_LEN)
    key = derive_key(password, salt)

    with open(input_path, "rb") as f:
        plaintext = f.read()

    aesgcm = AESGCM(key)
    # encrypt() returns ciphertext + 16-byte GCM tag appended at the end
    ciphertext_with_tag = aesgcm.encrypt(nonce, plaintext, associated_data=None)

    with open(output_path, "wb") as f:
        f.write(MAGIC)
        f.write(salt)
        f.write(nonce)
        f.write(ciphertext_with_tag)  # tag is the last TAG_LEN bytes

    print(f"[+] Encrypted → {output_path}")
    print(f"    Salt  : {salt.hex()}")
    print(f"    Nonce : {nonce.hex()}")


def decrypt_file(input_path: str, output_path: str, password: str) -> None:
    """Decrypt a file encrypted with encrypt_file()."""
    with open(input_path, "rb") as f:
        data = f.read()

    # Validate magic header
    if data[:4] != MAGIC:
        raise ValueError("Not a valid encrypted file (bad magic header).")

    offset = 4
    salt = data[offset:offset + SALT_LEN];        offset += SALT_LEN
    nonce = data[offset:offset + NONCE_LEN];      offset += NONCE_LEN
    ciphertext_with_tag = data[offset:]           # rest is ciphertext + tag

    key = derive_key(password, salt)
    aesgcm = AESGCM(key)

    # decrypt() raises InvalidTag if password is wrong or file was tampered
    plaintext = aesgcm.decrypt(nonce, ciphertext_with_tag, associated_data=None)

    with open(output_path, "wb") as f:
        f.write(plaintext)

    print(f"[+] Decrypted → {output_path}")


# --- CLI ---

def get_password(confirm: bool = False) -> str:
    pwd = getpass.getpass("Password: ")
    if confirm:
        pwd2 = getpass.getpass("Confirm password: ")
        if pwd != pwd2:
            print("Passwords do not match.", file=sys.stderr)
            sys.exit(1)
    return pwd


def main():
    parser = argparse.ArgumentParser(
        description="AES-256-GCM file encryption/decryption",
        formatter_class=argparse.RawDescriptionHelpFormatter,
        epilog="""
Examples:
  python aes_encrypt.py encrypt secret.pdf secret.pdf.enc
  python aes_encrypt.py decrypt secret.pdf.enc secret_out.pdf
  python aes_encrypt.py encrypt photo.jpg photo.jpg.enc --password mypass
        """,
    )
    parser.add_argument("mode", choices=["encrypt", "decrypt"],
                        help="Operation mode")
    parser.add_argument("input",  help="Input file path")
    parser.add_argument("output", help="Output file path")
    parser.add_argument("--password", "-p",
                        help="Password (omit to be prompted securely)")

    args = parser.parse_args()

    if not os.path.isfile(args.input):
        print(f"Error: input file not found: {args.input}", file=sys.stderr)
        sys.exit(1)

    password = args.password or get_password(confirm=(args.mode == "encrypt"))

    try:
        if args.mode == "encrypt":
            encrypt_file(args.input, args.output, password)
        else:
            decrypt_file(args.input, args.output, password)
    except Exception as e:
        print(f"Error: {e}", file=sys.stderr)
        sys.exit(1)


if __name__ == "__main__":
    main()
