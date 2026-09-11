#!/usr/bin/env bash
# Abre la interfaz.
#
# La primera vez, ademas, se agrega al menu de aplicaciones. Hace falta
# porque en Linux el doble clic sobre un .sh no lo ejecuta: los gestores de
# archivos lo abren en el editor por seguridad. Con el acceso directo queda
# como cualquier otro programa.

cd "$(dirname "$0")" || exit 1
AQUI="$(pwd)"

# --- acceso directo, solo la primera vez ---
ACCESO="$HOME/.local/share/applications/descarga-fotos.desktop"
if [ ! -f "$ACCESO" ] && [ -n "$HOME" ]; then
    mkdir -p "$(dirname "$ACCESO")"
    cat > "$ACCESO" <<EOF
[Desktop Entry]
Type=Application
Name=Descarga de fotos de producto
Comment=Baja las fotos de los productos de un excel de estructura
Exec=bash "$AQUI/iniciar.sh"
Path=$AQUI
Icon=applications-graphics
Terminal=false
Categories=Utility;
EOF
    chmod +x "$ACCESO" "$AQUI/iniciar.sh" 2>/dev/null
    update-desktop-database "$(dirname "$ACCESO")" 2>/dev/null || true
    echo "Agregado al menu de aplicaciones como 'Descarga de fotos de producto'."
fi

# --- abrir la ventana ---
if [ -x "venv/bin/python" ]; then
    exec venv/bin/python app.py
fi

if command -v python3 >/dev/null 2>&1; then
    exec python3 app.py
fi

echo "No se encontro Python 3. Instalalo con:  sudo apt install python3 python3-tk"
read -r -p "Enter para cerrar..."