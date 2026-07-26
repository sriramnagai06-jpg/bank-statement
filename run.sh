#!/bin/bash
echo "Installing required packages..."
pip install -r backend/requirements.txt
echo ""
echo "Starting Bank Statement Analyzer..."
python run.py
