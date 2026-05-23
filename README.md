# LOCVIEW

Visionneuse de documents locale pour Windows — PDF, EPUB, images et plus.  
100 % local, aucune donnée ne quitte la machine.

## Formats supportés

- PDF, EPUB, CBZ, FB2, XPS
- Images : JPG, PNG, GIF, BMP, TIFF, WEBP, SVG
- Texte : DOCX, PPTX, ODP, TXT, MD

## Fonctionnalités

- Scroll continu entre les pages (transition fluide)
- Onglets multiples
- Zoom, rotation, mode plein écran
- Recherche dans le texte
- Marque-pages
- Sélection et copie de texte
- Annotations (dessin, signatures, tampons texte)
- Impression
- Instance unique — double-cliquer un fichier l'ouvre dans un nouvel onglet

## Compilation

Requiert Python 3.12 et les dépendances listées ci-dessous.

```bash
pip install pymupdf pillow tkinterdnd2 pyinstaller
```

Compiler l'exécutable :

```
compile.bat
```

Créer l'installeur :

```
python make_installer.py
```

## Dépendances

- [PyMuPDF](https://github.com/pymupdf/PyMuPDF) — rendu PDF/EPUB
- [Pillow](https://python-pillow.org/) — traitement d'images
- [tkinterdnd2](https://github.com/pmgagne/tkinterdnd2) — glisser-déposer
- [NSIS](https://nsis.sourceforge.io/) — création de l'installeur Windows
