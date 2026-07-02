# Menú web Tina

Menú digital de **Tina — café, té y pastelería**. Página estática (HTML + Tailwind CDN) con los precios cargados desde `precios.js`.

## Estructura

```
menu-web-tina/
├── index.html          Página del menú (única vista)
├── precios.js          Precios de todos los ítems
└── assets/
    ├── logo/           Logo del menú y fondo del hero
    ├── iconos/         Íconos de cada sección + sin TACC + reloj de té
    └── tortas/         Fotos de la pastelería moderna
```

## Ver localmente

Abrir `index.html` en el navegador, o servirlo:

```bash
python3 -m http.server 8000
# luego abrir http://localhost:8000
```

## Actualizar precios

Editar `precios.js`. Cada ítem se enlaza al HTML por su atributo `data-precio`.
