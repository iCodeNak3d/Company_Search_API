#!/bin/bash

# Script pour exécuter facilement enrich_api.py

if [ "$1" = "--help" ] || [ "$1" = "-h" ]; then
  echo "Usage: ./run.sh [input_file] [output_file]"
  echo "  input_file  : Chemin vers le fichier Excel d'entrée (par défaut: data.xlsx)"
  echo "  output_file : Chemin vers le fichier Excel de sortie (par défaut: data_enrichi.xlsx)"
  exit 0
fi

INPUT_FILE="${1:-data.xlsx}"
OUTPUT_FILE="${2:-data_enrichi.xlsx}"

echo "Exécution du script avec :"
echo " - Fichier d'entrée  : $INPUT_FILE"
echo " - Fichier de sortie : $OUTPUT_FILE"
echo ""

python3 enrich_api.py --input "$INPUT_FILE" --output "$OUTPUT_FILE"

if [ $? -eq 0 ]; then
  echo ""
  echo "Traitement terminé avec succès. Résultats disponibles dans :"
  echo " - $OUTPUT_FILE (Excel)"
  echo " - results_*.csv (CSV)"
else
  echo ""
  echo "Une erreur s'est produite lors du traitement."
fi