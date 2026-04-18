#!/usr/bin/env bash
set -euo pipefail

RED='\033[0;31m'
GREEN='\033[0;32m'
NC='\033[0m'

if [[ $EUID -ne 0 ]]; then
    echo -e "${RED}[x]${NC} Esegui con sudo:  sudo bash uninstall.sh"
    exit 1
fi

echo -e "${GREEN}[+]${NC} Rimozione Gestione Turni Acustica..."

rm -rf /opt/turni-acustica
rm -f /usr/bin/turni-acustica
rm -f /usr/share/applications/turni-acustica.desktop
rm -f /usr/share/icons/hicolor/scalable/apps/turni-acustica.svg

gtk-update-icon-cache -f /usr/share/icons/hicolor/ 2>/dev/null || true
update-desktop-database /usr/share/applications/ 2>/dev/null || true

echo -e "${GREEN}[+]${NC} Disinstallazione completata."
echo "    I dati utente in ~/.turni_acustica/ non sono stati rimossi."
