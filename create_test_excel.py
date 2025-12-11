#!/usr/bin/env python3
"""
Script pour créer automatiquement le fichier Excel de test avec VBA
Utilise win32com pour piloter Excel via COM automation
"""

import sys
from pathlib import Path

try:
    import win32com.client
    HAS_WIN32COM = True
except ImportError:
    HAS_WIN32COM = False
    print("❌ pywin32 n'est pas disponible")
    sys.exit(1)


def create_test_excel():
    """Crée le fichier Excel de test avec VBA"""

    print("Creation du fichier Excel de test avec VBA...")
    print("-" * 70)

    # Lire le code VBA depuis le fichier exemple
    vba_code_file = Path("examples/sample_vba_code.txt")

    if not vba_code_file.exists():
        print(f"Fichier non trouve: {vba_code_file}")
        return False

    print("Lecture du code VBA depuis examples/sample_vba_code.txt...")

    # Extraire le code VBA du fichier texte
    with open(vba_code_file, 'r', encoding='utf-8') as f:
        content = f.read()

    # Trouver le code VBA entre les marqueurs
    start_marker = "CODE VBA À COPIER :"
    end_marker = "FIN DU CODE VBA"

    start_idx = content.find(start_marker)
    end_idx = content.find(end_marker)

    if start_idx == -1 or end_idx == -1:
        print("Impossible de trouver le code VBA dans le fichier")
        return False

    # Extraire le code VBA
    vba_code = content[start_idx + len(start_marker):end_idx].strip()

    # Nettoyer les lignes de séparation
    vba_lines = []
    for line in vba_code.split('\n'):
        if line.strip() and not line.startswith('='):
            vba_lines.append(line)

    vba_code = '\n'.join(vba_lines)

    print(f"Code VBA extrait ({len(vba_lines)} lignes)")
    print()

    # Créer le fichier Excel
    output_file = Path("examples/test_simple.xlsm").absolute()

    print(f"Lancement d'Excel via COM automation...")

    try:
        # Créer une instance Excel
        excel = win32com.client.Dispatch("Excel.Application")
        excel.Visible = False  # Mode invisible
        excel.DisplayAlerts = False  # Pas de popup

        print("Excel lance")

        # Créer un nouveau classeur
        workbook = excel.Workbooks.Add()
        print("Nouveau classeur cree")

        # Accéder au projet VBA
        vb_project = workbook.VBProject

        # Ajouter un module standard
        vb_module = vb_project.VBComponents.Add(1)  # 1 = vbext_ct_StdModule
        vb_module.Name = "Module1"

        print("Module VBA cree")

        # Insérer le code VBA
        vb_module.CodeModule.AddFromString(vba_code)

        print(f"Code VBA insere ({vb_module.CodeModule.CountOfLines} lignes)")

        # Enregistrer en .xlsm (51 = xlOpenXMLWorkbookMacroEnabled)
        workbook.SaveAs(str(output_file), FileFormat=52)

        print(f"Fichier enregistre: {output_file}")

        # Fermer Excel
        workbook.Close(SaveChanges=False)
        excel.Quit()

        print("Excel ferme")
        print()
        print("=" * 70)
        print("Fichier Excel cree avec succes!")
        print(f"Emplacement: {output_file}")
        print()
        print("Vous pouvez maintenant tester:")
        print("   1. python test_local.py")
        print("   2. ./venv/Scripts/pytest tests/ -v")

        return True

    except Exception as e:
        print(f"Erreur lors de la creation du fichier Excel:")
        print(f"   {type(e).__name__}: {e}")
        print()
        print("Causes possibles:")
        print("   1. Excel n'est pas installe sur cette machine")
        print("   2. Les macros VBA sont bloquees par la securite")
        print("   3. Acces VBA Project refuse (voir Fichier > Options > Centre de gestion)")
        print()
        print("Solution manuelle:")
        print("   Creez le fichier manuellement en suivant SETUP_INSTRUCTIONS.md")

        # Nettoyer
        try:
            excel.Quit()
        except:
            pass

        return False


if __name__ == "__main__":
    if not HAS_WIN32COM:
        print("pywin32 requis. Installez avec: pip install pywin32")
        sys.exit(1)

    success = create_test_excel()
    sys.exit(0 if success else 1)
