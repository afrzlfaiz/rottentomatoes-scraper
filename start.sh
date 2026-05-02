#!/usr/bin/env bash
set -e

python scrape_rottentomatoes.py

read -n 1 -s -r -p "Press any key to close..."
echo
