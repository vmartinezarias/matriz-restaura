# Matriz RESTAURA — visualizador

Visualizador interactivo de la matriz de monitoreo de la restauración (SIAT-AC / SINCHI).
Permite navegar la matriz como un árbol (fase → componente → unidad de monitoreo → indicador → variable),
ver la ficha completa de cada elemento y anotar observaciones que quedan guardadas en este mismo repositorio.

**La base es el Excel** `data/Matriz_RESTAURA.xlsx`: la página lo lee directamente cada vez que se abre.

## Estructura

```
index.html                       Página principal
config.js                        Configuración (ruta del Excel, repositorio, rama)
css/styles.css                   Estilos
js/app.js                        Lógica
data/Matriz_RESTAURA.xlsx        La matriz (fuente de verdad)
observaciones/observaciones.json Observaciones del equipo (lo escribe la app)
```

## Publicar en GitHub Pages

1. Crea un repositorio (p. ej. `matriz-restaura`) y sube todo el contenido de esta carpeta.
2. En el repositorio: **Settings → Pages → Build and deployment → Source: Deploy from a branch**,
   rama `main`, carpeta `/ (root)`. Guarda.
3. En uno o dos minutos la página queda en `https://TU_USUARIO.github.io/matriz-restaura/`.

El usuario y el repositorio se detectan solos a partir de esa dirección; no hace falta editar `config.js`.

> En el plan gratuito, GitHub Pages exige que el repositorio sea público, así que las observaciones
> también serán públicas. Si eso no conviene, se puede usar una cuenta u organización con plan de pago,
> o mantener las observaciones solo en los navegadores y compartirlas exportando el JSON.

## Actualizar la matriz

Reemplaza `data/Matriz_RESTAURA.xlsx` por la nueva versión **con el mismo nombre** y haz commit.
La app espera la hoja principal con la fila de encabezados que empieza en «Fase» (las columnas se
reconocen por su nombre, así que se pueden reordenar). También usa, si existen, la hoja de
entidades SIG (`Modelo_entidades_SIG`) y la hoja resumen de indicadores (la que tiene la columna
«Revisiones»).

Las observaciones se asocian a los valores de fase, componente, unidad, código de indicador y código
de variable; si se cambian esos textos en el Excel, las observaciones antiguas siguen en el JSON
y en la exportación, pero ya no aparecerán colgadas del elemento renombrado.

Para revisar otra versión sin subirla, usa **Configuración → Abrir otra versión del Excel**
(solo se ve en tu navegador).

## Cómo se guardan las observaciones

- Siempre se guardan en el navegador de quien las escribe.
- La página **lee** `observaciones/observaciones.json` del repositorio sin credenciales, así que
  todos ven las observaciones del equipo.
- Para **subir** observaciones, cada persona pega una vez un token personal de GitHub en
  **Configuración**. La app hace el commit del JSON (y si dos personas guardan a la vez, combina
  ambos cambios y reintenta).
- Desde **Observaciones** se descargan en Excel, CSV (separado por `;`) o JSON, y se puede
  importar un JSON (por ejemplo, el de alguien que trabajó sin token).

### Crear el token (una vez por persona)

1. GitHub → foto de perfil → **Settings → Developer settings → Personal access tokens →
   Fine-grained tokens → Generate new token**.
2. **Resource owner**: el dueño del repositorio. **Repository access**: *Only select repositories*
   → `matriz-restaura`.
3. **Permissions → Repository permissions → Contents: Read and write**. Nada más.
4. Elige una fecha de vencimiento, genera el token y pégalo en la configuración de la app.

Quien no sea dueño del repositorio debe estar añadido como colaborador (Settings → Collaborators)
o el repositorio debe pertenecer a una organización donde tenga acceso.

El token queda guardado solo en ese navegador (localStorage) y puede borrarse con **Olvidar token**.
No lo uses en computadores compartidos.

### Evitar que cada observación vuelva a desplegar la página (opcional)

Cada commit a `main` reconstruye GitHub Pages (tarda ~1 minuto y no afecta el uso). Si prefieres
evitarlo, crea una rama `observaciones`, y en `config.js` pon `branch: 'observaciones'`.

## Uso rápido

- **Ordenar por**: cambia la jerarquía del árbol (por fase, por indicador, por unidad o por componente).
- **Franja de fases**: salta a cada fase; el ancho de cada franja es proporcional a su número de variables.
- **Filtros**: nivel de evidencia (N1, N2, N3, TD), momento de medición, obligatoriedad y elementos con observaciones.
- **Buscar**: texto libre en todas las columnas (atajo: tecla `/`).
- **Anotar sobre un campo**: pasa el cursor sobre una fila de la ficha y usa el botón de la derecha.
- Los enlaces de la barra de direcciones apuntan al elemento abierto, así que se pueden compartir.

## Probar en local

Abrir `index.html` con doble clic no funciona (el navegador bloquea la lectura del Excel). Usa un
servidor local desde esta carpeta, por ejemplo:

```
python -m http.server 8000
```

y abre `http://localhost:8000`. En local las observaciones quedan en el navegador salvo que
completes usuario, repositorio y token en la configuración.
