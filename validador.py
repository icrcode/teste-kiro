"""
Ponto de entrada da CLI no diretório raiz do projeto.
Delega para arquitetura.validador para que o comando
    python validador.py <caminho_yaml> [--strict]
funcione a partir da raiz do projeto.
"""
from arquitetura.validador import main  # noqa: F401

if __name__ == "__main__":
    main()
