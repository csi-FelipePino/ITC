# ITC

# Proyecto de Gestión de Estados de Tránsito Semaforizados

Este proyecto surge como una solución para el equipo de transporte encargado de gestionar los estados de tránsito semaforizados. El objetivo es extraer, transformar y organizar los datos necesarios para el análisis de patrones y estados de funcionamiento de los semáforos, facilitando la administración y seguimiento de los mismos en distintos escenarios de tránsito.

## Descripción General

El proceso comienza con la obtención de un archivo en formato `.PTC2`, generado a partir del software **ITC** (herramienta de control de tráfico), que contiene la información relevante para los semáforos. Este archivo se transforma a formato `CSV` mediante Excel, a partir de lo cual se generan y alimentan diez tablas distintas que agrupan los datos según el tipo de análisis que se está realizando.

Cada tabla cumple una función específica en el análisis y administración de los datos de tránsito y semaforización. Estas tablas organizan la información en diferentes categorías, facilitando su análisis en posteriores revisiones y reportes de gestión.

## Estructura de Tablas

1. **Tablas 1 a 5**: Contienen datos agrupados en función de los distintos estados y eventos del semáforo.
2. **Tabla 6**: Implementa una estructura útil para el análisis de los datos de tránsito. Sin embargo, cabe señalar que su implementación actual no es la más eficiente a nivel de programación, y existen oportunidades para optimizar su rendimiento y organización en futuras versiones.
3. **Tablas 7 a 8**: Organizan información adicional y de soporte para el análisis de secuencias y patrones en los estados de los semáforos.
4. **Tabla 9**: Esta tabla se puede construir completamente en la migración inicial a Excel. No obstante, debido a ciertas limitaciones inherentes a Excel, esta se estructurará parcialmente para asegurar compatibilidad y accesibilidad de los datos.

## Futuras Mejora y Optimización

- **Optimización de la Tabla 6**: Se considera que en versiones futuras se podría implementar una lógica de programación más eficiente para mejorar su velocidad y consumo de recursos.
- **Adaptación de la Tabla 9**: La tabla 9 se presenta actualmente en un formato parcial en Excel, debido a las limitaciones técnicas de este software. En futuras versiones, se explorarán alternativas para su estructura completa y su compatibilidad con Excel.

## Conclusión

Este proyecto proporciona una solución funcional para la gestión de estados de tránsito semaforizados y facilita el trabajo del equipo de transporte. Aunque se han identificado algunas oportunidades de mejora, el sistema actual cumple con los requisitos de análisis y administración de los datos provenientes del software ITC. Este proceso de generación de tablas ofrece al equipo de transporte una herramienta efectiva para tomar decisiones basadas en datos en el ámbito de la gestión semafórica.

## Equipo del Proyecto

Este proyecto ha contado con la colaboración del equipo de transporte, conformado por:
- **Santiago Noboa**
- **Mauro Bruzzone**
- **Gastón Bonfiglio**

