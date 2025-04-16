import sys
import os
sys.path.insert(0, os.path.abspath(os.path.join(os.path.dirname(__file__), "src")))


from src.interface import iniciar_interface

def main():
    iniciar_interface()

if __name__ == "__main__":
    main()
