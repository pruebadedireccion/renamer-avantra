# Politica de seguridad

## Alcance

Renamer Avantra es una herramienta de escritorio que procesa archivos
locales. No tiene servidores, conexion a red ni telemetria.

## Reportar una vulnerabilidad

Si encuentras un problema de seguridad, por favor **no abras un issue
publico**. Escribe a traves de:

- [GitHub Security Advisories](https://github.com/pruebadedireccion/renamer-avantra/security/advisories/new)
  (recomendado)

Incluye:

- Version afectada
- Pasos para reproducir
- Impacto esperado
- Entorno (version de Windows, Python si aplica)

## Tiempos de respuesta

| Etapa | Tiempo orientativo |
|---|---|
| Confirmacion de recepcion | 3 dias habiles |
| Evaluacion inicial | 7 dias habiles |
| Correccion (si aplica) | segun severidad |

## Practicas del proyecto

- Procesamiento 100 % local, sin red ni telemetria.
- Operacion no destructiva: los archivos originales nunca se modifican.
- Copias con confirmacion explicita y omision de destinos existentes.
- Datos de ejemplo totalmente ficticios en el repositorio.
- Escaneo de secretos activo en GitHub (push protection).
- Dependencias actualizadas mediante Dependabot.
