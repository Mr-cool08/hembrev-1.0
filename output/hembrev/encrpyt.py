# @hidden_cell
from windtalker import SymmetricCipher
import os
def encrypt():
    c = SymmetricCipher(password="password")
    try:
        c.encrypt_file("password.txt")
        os.remove("password.txt")
    except OSError:
        os.remove("password-encrypted.txt")
    
    
def decrypt():
    c = SymmetricCipher(password="password")
    try:
        c.decrypt_file("password-encrypted.txt")
        os.remove("password-encrypted.txt")
    except:
        pass
    



    
