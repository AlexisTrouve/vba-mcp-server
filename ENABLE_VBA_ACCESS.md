# Activer l'accès programmatique VBA dans Excel

## Problème rencontré

```
Programmatic access to Visual Basic Project is not trusted
```

Cela signifie qu'Excel bloque l'accès au projet VBA par des programmes externes (comme notre script Python).

## Solution : Activer l'accès VBA

### Étape 1 : Ouvrir les options Excel

1. **Ouvrez Microsoft Excel**
2. **Fichier** → **Options**

### Étape 2 : Accéder au Centre de gestion de la confidentialité

3. Dans la fenêtre Options, cliquez sur **Centre de gestion de la confidentialité** (à gauche)
4. Cliquez sur le bouton **Paramètres du Centre de gestion de la confidentialité...**

### Étape 3 : Activer l'accès VBA

5. Dans le Centre de gestion, cliquez sur **Paramètres des macros** (à gauche)
6. **Cochez la case** :
   ```
   ☑ Faire confiance à l'accès au modèle objet du projet VBA
   ```
7. Cliquez sur **OK**
8. Cliquez sur **OK** dans la fenêtre Options
9. **Fermez Excel complètement**

### Étape 4 : Relancer le script

```bash
python create_test_excel.py
```

Le fichier `examples/test_simple.xlsm` devrait être créé automatiquement !

---

## ⚠️ Note de sécurité

Cette option réduit la sécurité d'Excel en permettant aux programmes externes d'accéder au VBA.

**Recommandations** :
- Activez cette option uniquement sur votre machine de développement
- Désactivez-la après avoir créé le fichier de test si vous le souhaitez
- Ne l'activez JAMAIS sur une machine de production

---

## Alternative : Création manuelle

Si vous ne voulez pas modifier les paramètres de sécurité, créez le fichier manuellement :

1. Ouvrez Excel
2. Alt + F11 (éditeur VBA)
3. Insertion → Module
4. Copiez le code depuis `examples/sample_vba_code.txt`
5. Enregistrez sous `examples/test_simple.xlsm` (type: macro-enabled)

---

## Résumé visuel

```
Excel → Fichier → Options → Centre de gestion de la confidentialité
    → Paramètres du Centre de gestion de la confidentialité
    → Paramètres des macros
    → ☑ Faire confiance à l'accès au modèle objet du projet VBA
    → OK → OK
```
