#!/usr/bin/env bash
set -euo pipefail

APP_NAME="turni-acustica"
INSTALL_DIR="/opt/${APP_NAME}"
VENV_DIR="${INSTALL_DIR}/venv"
SCRIPT_DIR="$(cd "$(dirname "${BASH_SOURCE[0]}")" && pwd)"

# Colori
RED='\033[0;31m'
GREEN='\033[0;32m'
YELLOW='\033[1;33m'
NC='\033[0m'

info()  { echo -e "${GREEN}[+]${NC} $1"; }
warn()  { echo -e "${YELLOW}[!]${NC} $1"; }
error() { echo -e "${RED}[x]${NC} $1"; exit 1; }

# Controlla root
if [[ $EUID -ne 0 ]]; then
    error "Esegui con sudo:  sudo bash install.sh"
fi

REAL_USER="${SUDO_USER:-$USER}"
info "Installazione Gestione Turni Acustica per EndeavourOS"
echo ""

# Dipendenze di sistema
info "Installazione dipendenze di sistema..."
pacman -S --needed --noconfirm python tk python-pip 2>/dev/null || {
    warn "pacman fallito, assicurati che python e tk siano installati"
}

# Copia applicazione
info "Installazione applicazione in ${INSTALL_DIR}..."
mkdir -p "${INSTALL_DIR}"
cp -r "${SCRIPT_DIR}/turni" "${INSTALL_DIR}/"
cp "${SCRIPT_DIR}/turni_v16.py" "${INSTALL_DIR}/"
cp "${SCRIPT_DIR}/requirements.txt" "${INSTALL_DIR}/"

# Ambiente virtuale
info "Creazione ambiente virtuale e installazione dipendenze Python..."
python3 -m venv "${VENV_DIR}"
"${VENV_DIR}/bin/pip" install --quiet --upgrade pip
"${VENV_DIR}/bin/pip" install --quiet -r "${INSTALL_DIR}/requirements.txt"

# Launcher script
info "Creazione launcher..."
cat > "/usr/bin/${APP_NAME}" << 'LAUNCHER'
#!/usr/bin/env bash
exec /opt/turni-acustica/venv/bin/python /opt/turni-acustica/turni_v16.py "$@"
LAUNCHER
chmod 755 "/usr/bin/${APP_NAME}"

# Desktop entry
info "Installazione desktop entry..."
cp "${SCRIPT_DIR}/packaging/turni-acustica.desktop" \
    /usr/share/applications/turni-acustica.desktop

# Icona
info "Installazione icona..."
mkdir -p /usr/share/icons/hicolor/scalable/apps/
cp "${SCRIPT_DIR}/packaging/turni-acustica.svg" \
    /usr/share/icons/hicolor/scalable/apps/turni-acustica.svg

# Aggiorna cache icone
gtk-update-icon-cache -f /usr/share/icons/hicolor/ 2>/dev/null || true
update-desktop-database /usr/share/applications/ 2>/dev/null || true

# Directory dati utente
USER_HOME="$(eval echo ~"${REAL_USER}")"
mkdir -p "${USER_HOME}/.turni_acustica"
chown -R "${REAL_USER}:${REAL_USER}" "${USER_HOME}/.turni_acustica"

echo ""
info "Installazione completata!"
echo ""
echo "  Avvio da menu:     cerca 'Gestione Turni' nel launcher"
echo "  Avvio da terminale: turni-acustica"
echo "  Disinstallazione:   sudo bash $(dirname "$0")/uninstall.sh"
echo ""
