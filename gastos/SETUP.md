# Mis Finanzas — Guía de instalación desde cero

> **Tiempo estimado:** 10–15 minutos  
> **Lo que necesitas:** Una cuenta de Google (Gmail)

---

## Paso 1 — Crea el Google Sheet

1. Abre [sheets.google.com](https://sheets.google.com) en tu navegador
2. Haz clic en el botón **"+"** (hoja en blanco)

   ```
   ┌─────────────────┐
   │  + En blanco    │  ← clic aquí
   └─────────────────┘
   ```

3. En la parte superior donde dice **"Hoja de cálculo sin título"**, haz clic y escribe un nombre:
   ```
   Mis Finanzas 2026
   ```
4. Presiona **Enter**

✅ Ya tienes la hoja creada.

---

## Paso 2 — Abre el editor de Apps Script

Con la hoja abierta:

1. En el menú superior haz clic en **Extensiones**
2. Selecciona **Apps Script**

   ```
   Extensiones  ←  clic aquí
   └── Apps Script  ←  clic aquí
   ```

3. Se abre una pestaña nueva que se ve como un editor de código

---

## Paso 3 — Pega el código principal (Code.gs)

En el editor ya hay algo escrito que dice `function myFunction() {}`.

1. Haz clic dentro del editor
2. Selecciona todo con **Ctrl+A** (Windows) o **Cmd+A** (Mac)
3. Bórralo con **Suprimir** o **Backspace**
4. Abre el archivo **`Code.gs`** de este proyecto y copia **todo** su contenido (**Ctrl+A** → **Ctrl+C**)
5. Pégalo en el editor (**Ctrl+V**)
6. Guarda con **Ctrl+S**

   > En la parte superior izquierda del editor verás el nombre **"Código.gs"** — eso es normal, es el mismo archivo.

---

## Paso 4 — Crea el archivo de la pantalla (index.html)

1. En el panel izquierdo, busca la sección **"Archivos"**
2. Haz clic en el **"+"** que está al lado del título "Archivos"

   ```
   Archivos  [+]  ←  clic aquí
   └── Código.gs
   ```

3. Selecciona **"HTML"**
4. Te pide un nombre — escribe exactamente:
   ```
   index
   ```
   ⚠️ Solo `index`, sin `.html`, sin mayúsculas, sin espacios. Presiona **Enter**.

5. Se crea el archivo con algo de contenido de ejemplo
6. Haz clic dentro del editor
7. Selecciona todo (**Ctrl+A**) y bórralo
8. Abre el archivo **`index.html`** de este proyecto, copia **todo** su contenido (**Ctrl+A** → **Ctrl+C**)
9. Pégalo en el editor (**Ctrl+V**)
10. Guarda con **Ctrl+S**

   Ahora en el panel izquierdo deberías ver:
   ```
   Archivos
   └── Código.gs
   └── index.html  ←  aparece aquí
   ```

---

## Paso 5 — Despliega la aplicación

Este es el paso más importante:

1. En la parte superior derecha haz clic en el botón azul **"Implementar"**
2. Selecciona **"Nueva implementación"**

   ```
   [Implementar ▼]  ←  clic aquí
   └── Nueva implementación  ←  clic aquí
   ```

3. Haz clic en el **ícono de engranaje ⚙️** que aparece junto a "Seleccionar tipo"
4. Elige **"Aplicación web"**
5. Completa los campos así:

   | Campo | Valor |
   |---|---|
   | Descripción | `v1` (o lo que quieras) |
   | Ejecutar como | **Yo (tu email)** |
   | Quién tiene acceso | **Solo yo** |

6. Haz clic en **"Implementar"**

---

## Paso 6 — Autoriza los permisos

La primera vez Google pide permiso para que el script acceda a tu hoja:

1. Haz clic en **"Autorizar el acceso"**
2. Elige tu cuenta de Google
3. ⚠️ Puede aparecer una pantalla que dice **"Google no ha verificado esta aplicación"**
   - Eso es normal — es tu propio código en tu propia cuenta
   - Haz clic en **"Configuración avanzada"** (texto pequeño abajo)
   - Luego en **"Ir a (nombre del proyecto)"**
4. Haz clic en **"Permitir"**
5. Aparece una pantalla con una **URL larga** — **cópiala y guárdala**

   ```
   URL de la aplicación web:
   https://script.google.com/macros/s/XXXXX.../exec
   ```

✅ Esa URL es tu aplicación.

---

## Paso 7 — Abre la app

### En el computador
Pega la URL en cualquier navegador y listo.

### En el celular (recomendado — funciona como app)

**Android (Chrome):**
1. Abre la URL en Chrome
2. Toca los **3 puntos** del menú (arriba a la derecha)
3. Toca **"Agregar a pantalla de inicio"**
4. Confirma — aparece como ícono en tu pantalla

**iPhone (Safari):**
1. Abre la URL en Safari
2. Toca el ícono de **compartir** (cuadrado con flecha hacia arriba)
3. Toca **"Añadir a pantalla de inicio"**
4. Confirma — aparece como ícono en tu pantalla

---

## Lo que crea la app automáticamente en tu Google Sheet

La primera vez que uses la app se crean estas pestañas solas:

| Pestaña | Qué guarda |
|---|---|
| `Transacciones` | Todos los ingresos y egresos |
| `Presupuesto` | El presupuesto mensual por categoría |
| `Me_Deben` | Dinero prestado pendiente de cobro |

> ⚠️ No borres ni renombres estas pestañas.

---

## Categorías disponibles

### Egresos
Vivienda · Alimentos y supermercado · Salud y Bienestar · Transporte · Ocio y Entretenimiento · Comida fuera de casa · Cafés · Compras Personales · Compras Hogar · Viajes personales · Viajes de Trabajo · Ahorro · Deuda · Regalos a otros · Donaciones · Otro

### Ingresos
Salario · Saldo mes anterior · Freelance · Otro ingreso

### Cuentas
Cuenta de Ahorros · Efectivo · AFC

---

## Funciones de la app

| Vista | Qué hace |
|---|---|
| **Resumen** | Balance del mes, alertas de presupuesto, desglose por categoría y por cuenta |
| **Registrar** | Formulario rápido para ingresos, egresos y "me deben" |
| **Historial** | Lista del mes filtrable por tipo, cuenta o categoría |
| **Presupuesto** | Configura montos por categoría y ve el comparativo ejecutado vs presupuestado |
| **Tendencias** | Gráfica anual de ingresos vs egresos y promedios por categoría |

---

## Preguntas frecuentes

**¿Puedo usarla en el celular y el computador al mismo tiempo?**  
Sí. Los datos están en tu Google Sheet, cualquier dispositivo que abra la URL ve la misma información.

**¿Mis datos son privados?**  
Sí. La app vive en tu cuenta de Google y solo tú tienes acceso (a menos que cambies "Quién tiene acceso" al desplegarla).

**¿Tiene algún costo?**  
No. Google Apps Script es gratuito para uso personal.

**Cambié algo y no funciona, ¿qué hago?**  
Después de cambiar el `Code.gs` solo guarda (Ctrl+S) y recarga la URL.  
Después de cambiar el `index.html` debes volver a desplegar: **Implementar → Administrar implementaciones → ✏️ → Nueva versión → Implementar**.

**¿Cómo le comparto la app a otra persona?**  
Cada persona debe seguir esta guía y crear su propia copia — así cada quien tiene sus datos separados en su propia cuenta de Google.
