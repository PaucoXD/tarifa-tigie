# Tarifa TIGIE · A&C Group

Buscador de la Tarifa TIGIE: fracciones arancelarias, NICO, aranceles, NOM y regulaciones no arancelarias.

**Abrir en línea:** https://paucoxd.github.io/tarifa-tigie/

También se puede descargar `index.html` y abrirlo directo en el navegador.

## Fuentes y vigencia

- Fracciones, aranceles y NICO: SNICE / Secretaría de Economía, archivo del 14/10/2026 (incluye el decreto del DOF del 13/10/2026).
- NOM por fracción: Anexo 2.4.1 del Acuerdo de reglas y criterios de la SE (DOF 25/11/2022) con sus modificaciones del 13/01/2023, 16/08/2023, 22/03/2024, 28/08/2024, 29/05/2026 y 28/09/2026.
- Permisos, SADER, Salud, SEMARNAT, SEDENA, SENER, cuotas compensatorias, PROSEC, IMMEX, frontera y precios estimados: «Base Única» del SNICE del 14/10/2026.

Es información de referencia: confirma en el DOF antes de despachar.

## Cómo actualizar

1. Doble clic en **`Actualizar.cmd`** (necesita Python y haber iniciado sesión en GitHub en esta computadora).
2. Baja del SNICE el archivo de fracciones y la Base Única más recientes, y del DOF los acuerdos del Anexo 2.4.1. Regenera `index.html` y muestra qué cambió: fracciones nuevas o eliminadas, aranceles, NICO, NOM y regulaciones. El reporte queda en `tools/ultimo_reporte.txt`.
3. Pregunta si se sube a GitHub. La página publicada se actualiza en 1 o 2 minutos.

**Cuando avise de un acuerdo nuevo de la SE** («ATENCIÓN: hay acuerdos nuevos sin revisar»): si ese acuerdo modifica el Anexo 2.4.1 (NOM), agrega su número de nota del DOF y su fecha a `modificaciones` en `tools/nom_mods.json`. Agrega la fecha a `revisados` aunque no toque el anexo, para que deje de avisar. Luego vuelve a correr `Actualizar.cmd`. Si el acuerdo quita fracciones o NICO con texto libre («se elimina…», «se suprimen…»), agrégalo también en `eliminar_fracciones` o `suprimir_nicos`.

Si algo no cuadra (por ejemplo, el SNICE cambió el formato de sus archivos), el actualizador se detiene sin tocar la página.
