# Mis Finanzas — Instrucciones de Configuración

## Pasos para poner en marcha la app

### 1. Crea el Google Sheet

1. Ve a [sheets.google.com](https://sheets.google.com) y crea una hoja nueva
2. Ponle un nombre, por ejemplo: **"Mis Finanzas 2026"**

---

### 2. Abre el editor de Apps Script

En el Google Sheet: menú **Extensiones → Apps Script**

---

### 3. Copia los archivos

**Archivo `Code.gs`** (ya existe por defecto):
- Borra todo el contenido que trae por defecto
- Pega todo el contenido del archivo `Code.gs` de este proyecto

**Archivo `index.html`** (debes crearlo):
- Haz clic en el **+** junto a "Archivos" → **HTML**
- Llámalo exactamente `index` (sin extensión, Apps Script agrega `.html` automáticamente)
- Borra el contenido por defecto y pega el contenido del archivo `index.html` de este proyecto

---

### 4. Despliega como aplicación web

1. En Apps Script, haz clic en **"Implementar"** → **"Nueva implementación"**
2. Haz clic en el engranaje ⚙ → selecciona **"Aplicación web"**
3. Configura:
   - **Descripción**: Mis Finanzas
   - **Ejecutar como**: Yo (tu cuenta de Google)
   - **Quién tiene acceso**: Solo yo *(o "Cualquier usuario" si quieres compartirla)*
4. Haz clic en **"Implementar"**
5. Autoriza los permisos que pide (acceso a tu Google Sheet)
6. Copia la **URL de la aplicación web** que aparece

---

### 5. Abre la app

- **En el computador**: pega la URL en el navegador
- **En el celular**: abre la URL en Chrome/Safari → menú → **"Agregar a pantalla de inicio"** (la app funciona como PWA)

---

## Estructura del Google Sheet (se crea automáticamente)

La app crea estas hojas automáticamente al primer uso:

| Hoja | Contenido |
|------|-----------|
| `Transacciones` | Todos los ingresos y egresos registrados |
| `Presupuesto` | Presupuesto mensual por categoría |
| `Me_Deben` | Registro de dinero prestado pendiente de cobro |

> No modificues los encabezados de las hojas manualmente.

---

## Categorías de egresos disponibles

Vivienda · Alimentos y supermercado · Salud y Bienestar · Transporte · Ocio y Entretenimiento · Comida fuera de casa · Cafés · Compras Personales · Compras Hogar · Viajes personales · Viajes de Trabajo · Ahorro · Deuda · Regalos a otros · Donaciones · Otro

## Categorías de ingresos

Salario · Saldo mes anterior · Freelance · Otro ingreso
