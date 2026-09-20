# Dashboard de gastos

Dashboard local para consultar y analizar los datos de gastos almacenados en Google Sheets. El backend solo solicita acceso de lectura a la hoja y sirve la compilación del frontend.

## Requisitos

- Node.js compatible con las versiones fijadas en `package-lock.json`.
- Una cuenta de servicio de Google con acceso de lectura a la hoja matriz.
- La hoja debe contener, como mínimo, las pestañas `Config_App`, `API_Resumen`, `API_Dashboard`, `API_MediasCategoria` y `API_GastosPeriodicos`.

## Configuración

Crea `backend/.env` sin incluirlo en Git:

```dotenv
GOOGLE_APPLICATION_CREDENTIALS=credentials/archivo-cuenta-servicio.json
SPREADSHEET_ID=id-de-la-hoja
PORT=3001
# Opcional: lista separada por comas. Sin valor se permiten los orígenes de la LAN.
CORS_ORIGIN=http://equipo-local:5173
```

## Desarrollo

En dos terminales:

```powershell
cd backend
npm install
npm start
```

```powershell
cd frontend
npm install
npm run dev -- --host 0.0.0.0
```

Vite usa el backend de la misma máquina en el puerto `3001` por defecto. Puede sobrescribirse con `VITE_API_BASE_URL`.

## Producción local

```powershell
cd frontend
npm run build
cd ..\backend
npm start
```

El backend sirve `frontend/dist` y escucha en todas las interfaces para permitir acceso mediante LAN o ZeroTier. Comprueba el servicio en `/api/health`.

## Verificación

```powershell
cd frontend
npm run lint
npm run build
```

```powershell
cd backend
node --check server.js
```

## Datos y copias de seguridad

- Los extractos bancarios y las credenciales son datos locales y no deben añadirse a Git.
- Las operaciones de mantenimiento sobre documentos se hacen en modo lectura.
- Antes de una modificación relevante se guarda un punto de restauración en `.backups/<fecha>-<motivo>/`, con la estructura original y hashes SHA-256.
- `.backups` queda fuera de Git para evitar publicar datos o duplicar archivos históricos.

## Estructura

- `backend/server.js`: API, acceso de solo lectura a Google Sheets y servidor del frontend.
- `frontend/src`: aplicación React.
- `apps-script-auditoria.gs`: automatización asociada a la hoja matriz.
