# Sitio estático (HTML/CSS/JS, sin build). El menú lo lee en runtime desde
# Google Sheets vía fetch() del lado del navegador — no hace falta Node ni
# variables de entorno acá, sólo servir los archivos.
FROM nginx:alpine

COPY index.html precios.js /usr/share/nginx/html/
COPY assets/ /usr/share/nginx/html/assets/

EXPOSE 80
