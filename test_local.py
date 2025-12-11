#!/usr/bin/env python3
"""
Test local du VBA MCP Server (sans MCP)
Permet de tester l'extraction VBA directement sans passer par le protocole MCP.
"""

import asyncio
import sys
from pathlib import Path

# Add src directory to path
sys.path.insert(0, str(Path(__file__).parent / "src"))

from lib.office_handler import OfficeHandler
from lib.vba_parser import VBAParser


async def test_extraction():
    """Test l'extraction VBA depuis un fichier Office."""

    # Chemin vers le fichier de test
    test_file = Path("examples/test_simple.xlsm")

    if not test_file.exists():
        print("[ERROR] Fichier de test non trouve:", test_file)
        print()
        print("[INSTRUCTIONS]")
        print("1. Ouvrez Excel")
        print("2. Creez un nouveau classeur")
        print("3. Appuyez sur Alt + F11 pour ouvrir l'editeur VBA")
        print("4. Insertion -> Module")
        print("5. Copiez le code VBA depuis examples/sample_vba_code.txt")
        print("6. Enregistrez le fichier sous 'test_simple.xlsm'")
        print("   (Type: Classeur Excel prenant en charge les macros)")
        print("7. Placez le fichier dans le dossier examples/")
        return

    print("[TEST] Extraction VBA")
    print(f"Fichier: {test_file}")
    print("-" * 70)
    print()

    # Initialiser le handler
    handler = OfficeHandler()

    try:
        # Extraire le projet VBA
        print("[INFO] Extraction du projet VBA...")
        vba_project = handler.extract_vba_project(test_file)

        if not vba_project or not vba_project.get("modules"):
            print("[ERROR] Aucun module VBA trouve dans le fichier")
            print("   Verifiez que le fichier contient bien des macros VBA")
            return

        print(f"[SUCCESS] Trouve {len(vba_project['modules'])} module(s)")
        print()

        # Parser chaque module
        parser = VBAParser()

        for idx, module in enumerate(vba_project["modules"], 1):
            parsed = parser.parse_module(module)

            print(f"[MODULE {idx}] {parsed['name']} ({parsed['type']})")
            print(f"   - Lignes de code: {parsed['line_count']}")
            print(f"   - Procedures: {len(parsed['procedures'])}")

            if parsed['procedures']:
                for proc in parsed['procedures']:
                    proc_type = proc['type'].capitalize()
                    param_count = len(proc.get('parameters', []))
                    params_str = f" ({param_count} param{'s' if param_count > 1 else ''})" if param_count > 0 else ""
                    print(f"      * {proc_type}: {proc['name']}{params_str}")
            else:
                print(f"      * (Aucune procedure detectee)")

            print(f"   - Dependances: {', '.join(parsed.get('dependencies', [])) or 'Aucune'}")
            print()

        print("=" * 70)
        print("[SUCCESS] Test reussi! Le serveur VBA MCP fonctionne correctement.")
        print()
        print("[NEXT STEPS]")
        print("   1. Lancez les tests unitaires : venv/Scripts/pytest tests/ -v")
        print("   2. Configurez le serveur MCP dans Claude Code")
        print("   3. Testez avec Claude : 'Extract VBA from <chemin_fichier>'")

    except Exception as e:
        print(f"[ERROR] Erreur lors du test: {type(e).__name__}")
        print(f"   Message: {str(e)}")
        print()
        print("[DEBUG]")
        import traceback
        traceback.print_exc()


if __name__ == "__main__":
    asyncio.run(test_extraction())
