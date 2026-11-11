// Guías de interpretación y glosario del visualizador.
// Este archivo es independiente del Excel: al reemplazar la matriz, las guías se conservan.
// Las guías se enganchan por el CÓDIGO del indicador (p. ej. 'IND-10') y las notas de
// variables por cada código de sistema (p. ej. 'dpc_nucleo_pct'), aunque en el Excel
// aparezcan varios códigos juntos en una misma celda.
// Formato del texto: **negrita** y `código`. Para agregar otra guía, copia el bloque de
// IND-10 con el código del nuevo indicador.
window.RESTAURA_GUIAS = {
  indicadores: {
    'IND-10': {
      titulo: 'Cómo leer la conectividad',
      pregunta: '¿La restauración está ayudando a que los fragmentos de bosque vuelvan a funcionar como una red?',
      resumen: 'La conectividad no se mide en campo: se **modela** cada dos años a partir de la cobertura del SIAT-AC (nivel TD) y se compara siempre con la línea base (T-1/T0). Como en una red de carreteras, no importa solo cuántos parches de bosque hay, sino qué tan fácil es pasar de uno a otro; la restauración construye las vías nuevas.',
      bloques: [
        {
          titulo: 'Corredor de vida: probabilidad de conectividad (PC)',
          texto: 'Probabilidad de que dos puntos tomados al azar en el corredor caigan en hábitat y además estén conectados entre sí (de 0 a 1). Lo que se reporta es su **cambio frente a T0** (`cambio_pc_pct`): «el PC del corredor subió 15 % desde la línea base».',
        },
        {
          titulo: 'Núcleo de reconexión: ECA local',
          texto: 'Área Conectada Equivalente, en hectáreas: el hábitat del núcleo funciona como si fuera **un solo parche continuo de esa área**. Si el ECA crece más que el área de bosque, la restauración está uniendo piezas, no solo sumando área. Por eso no se reporta aparte el parche mayor: el ECA nunca es menor que el área del parche más grande, así que ya lo contiene.',
        },
        {
          titulo: 'Núcleo de reconexión: dPC',
          texto: 'Importancia del núcleo para la red del corredor: cuánto cae el PC del corredor si se retira el hábitat del núcleo. Si el PC es 0,010 con el núcleo y 0,008 sin él, el dPC es 20 %: ese núcleo sostiene una quinta parte de la conectividad del corredor.',
        },
        {
          titulo: 'Mosaico: contexto',
          texto: 'No tiene un índice propio: se reportan la cobertura leñosa de la matriz y la pastura degradada total, y la conectividad se presenta a través de sus corredores y núcleos.',
        },
        {
          titulo: 'Complementos estructurales',
          texto: 'Distancia entre los remanentes que el núcleo busca unir, número y longitud de brechas, ancho mínimo y cobertura leñosa. Son las cifras más fáciles de comunicar («la brecha pasó de 800 m a 300 m»), pero describen geometría y **no prueban por sí solas** que haya conexión funcional.',
        },
      ],
      advertencias: [
        '**Es conectividad potencial, no movimiento observado**: el modelo indica que el paisaje permitiría el desplazamiento bajo los supuestos de hábitat y dispersión, no que los animales se muevan.',
        '**El dPC no es un cambio en el tiempo**: mide importancia relativa en un momento dado. El cambio temporal es `cambio_pc_pct`.',
        '**El ECA local no es el aporte del núcleo al corredor**: ese aporte es el dPC.',
        '**No se suman índices entre unidades**: PC, ECA y dPC no se agregan de núcleo a corredor ni de corredor a mosaico (regla D, emergente).',
        '**La comparabilidad lo es todo**: entre periodos se mantienen el ámbito, la resolución, el criterio de hábitat, las distancias de dispersión y el área de referencia `A_L`. Si algo cambia, la diferencia puede ser del método y no de la restauración.',
      ],
      frase: 'Medimos si la restauración reconecta el bosque: en cada corredor, cuánto aumenta la probabilidad de conectividad frente a la línea base; en cada núcleo, cuántas hectáreas funcionan como bosque conectado y cuánto pesa ese núcleo en la red; y como apoyo, cuántas brechas se cierran y cuánto se acortan.',
      variables: {
        pc_corredor: 'Probabilidad de conectividad de la red del corredor (0 a 1). Depende del área de referencia `A_L`, que debe quedar fija entre periodos.',
        cambio_pc_pct: 'Cambio del PC frente a T0, en %. Es la cifra que muestra si el corredor mejora o empeora en el tiempo; no se define si el PC de T0 es 0.',
        eca_local_nucleo_ha: 'Hectáreas que «valdría» el hábitat del núcleo si fuera un solo parche conectado. No es el aporte del núcleo al corredor.',
        dpc_nucleo_pct: 'Porcentaje del PC del corredor que se pierde al retirar el hábitat del núcleo. Importancia relativa en un momento dado, no cambio en el tiempo.',
        dist_remanentes_m: 'Distancia mínima entre los dos remanentes que el núcleo busca unir. Que estén cerca no demuestra que estén conectados.',
        brechas_n: 'Número de tramos del eje del corredor sin hábitat por encima del umbral.',
        brecha_long_m: 'Longitud de las brechas del eje del corredor. Fácil de comunicar, pero no prueba funcionalidad.',
        ancho_min_m: 'Ancho más angosto de la franja de hábitat del corredor.',
        cobertura_lenosa_pct: 'Porcentaje del corredor o de la matriz del mosaico con cobertura leñosa.',
        pastura_degradada_total_ha: 'Hectáreas de pastura degradada en el mosaico; contexto de la oportunidad de restauración.',
        integridad_vs_lineabase: 'Integridad ecológica frente a T0 (cobertura, estructura y composición calibradas con parcelas permanentes). No asumir causalidad sin un diseño de comparación.',
      },
    },
  },

  glosario: [
    ['PC', 'Índice de probabilidad de conectividad (Saura y Pascual-Hortal, 2007): probabilidad de que dos puntos al azar caigan en hábitat conectado. Adimensional, de 0 a 1.'],
    ['ECA', 'Área Conectada Equivalente: tamaño, en hectáreas, de un único parche que daría la misma conectividad que la red evaluada.'],
    ['dPC', 'Variación del PC al retirar un elemento (aquí, el hábitat de un núcleo); mide su importancia relativa en la red, en %.'],
    ['A_L', 'Área de referencia del paisaje usada como denominador del PC; debe ser la misma en todos los periodos.'],
    ['p*ᵢⱼ', 'Máxima probabilidad de dispersión entre los parches i y j a través de la red.'],
    ['TD', 'Teledetección: dato derivado de sensores remotos (SIAT-AC).'],
    ['N1 · N2 · N3', 'Niveles de evidencia: 100 % de las unidades; submuestra con grupo de control; parcelas permanentes de investigación.'],
    ['T-1 · T0', 'Año previo a la intervención (teledetección) y línea base.'],
    ['VS', 'Vegetación secundaria.'],
    ['IDP · ITP', 'Índice de Degradación de Pastura; Índice de Transición Predial.'],
    ['DiD', 'Diferencias en diferencias: comparación del cambio en fincas intervenidas frente al cambio en fincas control.'],
    ['Regla D', 'Agregación emergente: el valor solo existe en la unidad superior y no se obtiene sumando las inferiores.'],
  ],
};
