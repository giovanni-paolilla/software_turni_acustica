#!/usr/bin/env bash
# Launcher per Gestione Turni Acustica
export PYTHONPATH="/opt/turni-acustica/vendor${PYTHONPATH:+:$PYTHONPATH}"
exec python3 /opt/turni-acustica/turni_v16.py "$@"
