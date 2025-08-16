🧹 Macro de Limpieza para Cajas en Supermercado

Este proyecto es un Excel con macro automatizada para gestionar la limpieza de cajas en supermercados de forma organizada, equitativa y eficiente.

La macro asigna empleados a cajas en las horas menos concurridas, garantiza la rotación justa y mantiene un historial de control para auditoría.

📂 Archivos Necesarios

Archivo Excel habilitado para macros (.xlsm).

La macro calculará automáticamente la semana actual del año, sin necesidad de escribirla manualmente.

📑 Estructura del Archivo Excel
📄 Hoja 1 – Configuración

Lista de empleados AM y PM.

Lista de cajas AM y PM.

Horarios de trabajo de cada turno (AM y PM).

Semana del año calculada automáticamente con la fórmula:

=NUM.DE.SEMANA(HOY();2)


(El 2 indica que la semana inicia en lunes).

📄 Hoja 2 – Registro Histórico

Historial de todas las asignaciones:

Semana	Turno	Caja	Empleado	Hora Asignada
33	AM	Caja 1	Manuel	10:00
33	AM	Caja 2	Pedro	11:00
...	...	...	...	...
📄 Hoja 3 – Asignación Actual

Resultado de la asignación del día:

Caja	Empleado	Hora Asignada
Caja 1	Manuel	10:00
Caja 2	Pedro	11:00
...	...	...
📄 Hoja 4 – Horas de Visita Clientes

Registro de afluencia de clientes:

Hora	Visitantes	Clasificación
08:00	15	Alta
09:00	12	Alta
10:00	5	Baja
11:00	4	Baja

👉 La macro usará únicamente las horas con clasificación “Baja” para programar la limpieza.

⚙️ Lógica de la Macro

Calcular la semana actual automáticamente.

Leer listas de empleados y cajas de cada turno.

Revisar el historial para respetar la rotación.

Buscar horas “Baja” en la hoja de visitas:

✅ Si hay suficientes → asignarlas aleatoriamente.

❌ Si no hay suficientes → mostrar mensaje:

"No hay horas débiles suficientes, favor asignar manualmente"


Guardar asignaciones en el Historial.

Mostrar resultados en la hoja Asignación Actual.

📊 Reglas de Rotación

Se prioriza a los empleados con menos limpiezas acumuladas en ese turno.

En caso de empate, se asigna al que lleve más semanas sin limpiar.

Garantiza que nadie repita hasta que los demás alcancen su mismo conteo.

✅ Ventajas

🔄 Rotación justa entre empleados.

🕒 Limpieza programada en horas menos concurridas.

📆 Semana calculada automáticamente.

📝 Historial completo para auditoría.

✍️ Flexibilidad: permite asignar manualmente si es necesario.

▶️ Uso

Registra los empleados, cajas y horarios en la hoja Configuración.

Ingresa los datos de visitas por hora en Horas de Visita Clientes.

Pulsa el botón de la macro.

Verifica la asignación en Asignación Actual.

El sistema guardará automáticamente la información en Registro Histórico.
