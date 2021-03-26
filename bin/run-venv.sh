#!/usr/bin/env bash

if [ ! -d bin ]; then
    echo "Run this script from the root directory of the project" >&2
    exit 1
fi

if [ -d "./venv/Scripts" ] && [ -f "./venv/Scripts/activate" ]; then
    source ./venv/Scripts/activate &&
        python main.py "$@" &&
        deactivate
fi
