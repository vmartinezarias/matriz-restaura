// Configuración del visualizador de la Matriz RESTAURA.
// Normalmente no hace falta tocar nada: si la página está publicada en
// https://USUARIO.github.io/REPOSITORIO/, el usuario y el repositorio se detectan solos.
window.RESTAURA_CONFIG = {
  // Ruta del Excel dentro del repositorio. Para actualizar la matriz, reemplaza ese
  // archivo (con el mismo nombre) y haz commit.
  excel: 'data/Matriz_RESTAURA.xlsx',

  // Repositorio donde se guardan las observaciones (vacío = detección automática).
  owner: '',
  repo: '',
  // Rama donde se guardan las observaciones. Puedes usar una rama aparte
  // (p. ej. 'observaciones') para que cada observación no vuelva a desplegar Pages.
  branch: 'main',
  obsPath: 'observaciones/observaciones.json',
};
