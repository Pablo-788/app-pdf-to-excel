#!/bin/bash
echo "$(date) - Haciendo curl a la URL..."
curl -s https://pdf-cron-pinger.onrender.com/|| echo "Fallo la petición"