#!/bin/bash
apt-get update && apt-get install -y libreoffice --no-install-recommends
pip install -r requirements.txt
