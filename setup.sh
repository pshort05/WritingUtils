#!/bin/bash
# setup.sh — Install WritingUtils system-wide on Linux
#
# Installs clean-docx and clean-markdown as globally available commands.
# Uses pip in editable mode so edits to source take effect immediately.
# Requires sudo for system-wide install.

set -e

SCRIPT_DIR="$(cd "$(dirname "${BASH_SOURCE[0]}")" && pwd)"

# ---------------------------------------------------------------------------
# Helpers
# ---------------------------------------------------------------------------

info()    { echo "  $*"; }
success() { echo "✓ $*"; }
fail()    { echo "✗ $*" >&2; exit 1; }

# ---------------------------------------------------------------------------
# Locate pip
# ---------------------------------------------------------------------------

if command -v pip3 &>/dev/null; then
    PIP=pip3
elif command -v pip &>/dev/null; then
    PIP=pip
else
    fail "pip not found. Install Python 3 and pip first (e.g. sudo apt install python3-pip)"
fi

PYTHON=$(${PIP} --version | awk '{print $NF}' | tr -d ')')
info "Using ${PIP} (Python ${PYTHON})"

# ---------------------------------------------------------------------------
# Install the package system-wide (with --break-system-packages for
# modern Debian/Ubuntu systems that protect the global site-packages)
# ---------------------------------------------------------------------------

echo ""
echo "Installing WritingUtils..."

cd "${SCRIPT_DIR}"

# Try a plain install first; if pip refuses due to PEP 668 protection,
# retry with --break-system-packages.
install_cmd="sudo ${PIP} install -e ."

if ! ${install_cmd} 2>/tmp/wu_pip_err; then
    if grep -q "externally-managed" /tmp/wu_pip_err; then
        info "System pip is externally managed — retrying with --break-system-packages"
        sudo ${PIP} install --break-system-packages -e . \
            || fail "Installation failed. See output above."
    else
        cat /tmp/wu_pip_err >&2
        fail "Installation failed."
    fi
fi

# ---------------------------------------------------------------------------
# Verify commands landed in PATH
# ---------------------------------------------------------------------------

echo ""
for cmd in clean-docx clean-markdown; do
    if command -v "${cmd}" &>/dev/null; then
        success "${cmd}  →  $(command -v ${cmd})"
    else
        echo "⚠ ${cmd} not found in PATH after install."
        echo "  The install may have placed scripts in a directory not on your PATH."
        echo "  Try: export PATH=\"\$HOME/.local/bin:\$PATH\""
    fi
done

echo ""
echo "Done. Run 'clean-docx --help' or 'clean-markdown --help' to get started."
