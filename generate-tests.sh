#!/bin/bash
# Quick launcher for C# Test Generator
# Usage: ./generate-tests.sh [project-path] [--file filename.cs]

echo ""
echo "==================================="
echo " C# Automated Test Generator"
echo "==================================="
echo ""

if [ -z "$1" ]; then
    echo "Usage: ./generate-tests.sh [project-path] [--file filename.cs]"
    echo ""
    echo "Examples:"
    echo "  ./generate-tests.sh FindMissingAppointments"
    echo "  ./generate-tests.sh FindMissingAppointments --file Helper.cs"
    echo "  ./generate-tests.sh ."
    echo ""
    exit 1
fi

node scripts/test-generator/generate-tests.js "$@"

echo ""
echo "==================================="
echo " Generation Complete!"
echo "==================================="
