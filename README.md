# Clinica automatizar

App local liviana para cargar, guardar e imprimir la caratula del paciente.

## Lo que ya hace

- Formulario rapido para secretaria.
- `Fumador` en `No` por defecto.
- `Fecha` en el dia actual por defecto.
- Formato guiado para `DNI` y fechas.
- Calculo automatico de `Edad`.
- Listas editables de `Obra social` y `Deriva`.
- Sugerencias rapidas para escribir obra social y medico.
- Fichas recientes guardadas localmente.
- Vista previa A4 lista para imprimir.
- Descarga en `Word (.docx)` para correcciones manuales si hace falta.
- Guardado local en el navegador para no perder listas ni el ultimo formulario.
- No usa backend ni base de datos: en esta etapa los datos quedan solo en `localStorage` del navegador.

## Ejecutar

```bash
npm install
npm run dev
```

## Deploy en Vercel

La app ya queda lista para desplegar como sitio simple en Vercel.

- Build command: `npm run build`
- Output directory: `dist`
- Persistencia actual: `localStorage` del navegador de cada computadora

Build de produccion:

```bash
npm run build
```

## Stack elegido

- `Vite + TypeScript + HTML/CSS`
- Interfaz sobria y liviana
- Pensada para usarse como app local simple en una PC de recepcion

## Idea de uso final

1. La secretaria abre la app
2. Completa los datos con Enter para avanzar
3. En el ultimo campo, Enter abre impresion
4. Puede guardar la ficha y volver a cargarla despues
5. Si necesita corregir algo afuera del sistema, puede descargar `Word`
