# ================= FASE 1: CONSTRUCCIÓN =================
FROM mcr.microsoft.com/windows/servercore:ltsc2022 AS builder

WORKDIR C:/app

# 1. Instalar Node.js
ADD https://nodejs.org/dist/v20.11.0/node-v20.11.0-x64.msi node.msi
RUN start /wait msiexec /i node.msi /quiet /qn /norestart INSTALLDIR=C:\\nodejs && \
    del node.msi

# 2. Instalar Python 3.12 (necesario para node-gyp)
ADD https://www.python.org/ftp/python/3.12.3/python-3.12.3-amd64.exe python.exe
RUN start /wait python.exe /quiet InstallAllUsers=1 PrependPath=1 && \
    del python.exe

# 3. Instalar Visual Studio Build Tools (el compilador de C++ para los módulos nativos)
#    Esto es grande y tarda, pero es la forma más segura de tener todo lo necesario.
ADD https://aka.ms/vs/17/release/vs_BuildTools.exe vs_BuildTools.exe
RUN start /wait vs_BuildTools.exe --quiet --wait --add Microsoft.VisualStudio.Workload.VCTools --add Microsoft.VisualStudio.Component.VC.Tools.x86.x64 --add Microsoft.VisualStudio.Component.Windows10SDK.20348 --norestart && \
    del vs_BuildTools.exe

# 4. Configurar las variables de entorno para que npm y node-gyp encuentren las herramientas
ENV PYTHON="C:\Program Files\Python312\python.exe"
ENV PATH="C:\nodejs;C:\Program Files\Python312;C:\Program Files\Python312\Scripts;C:\Program Files (x86)\Microsoft Visual Studio\2022\BuildTools\MSBuild\Current\Bin;%PATH%"

# 5. Copiar, instalar dependencias (aquí se compilarán bcrypt y sqlite3) y copiar el código
COPY package*.json ./
RUN npm install
COPY . .

# ================= FASE 2: EJECUCIÓN =================
FROM mcr.microsoft.com/windows/nanoserver:ltsc2022 AS runtime

WORKDIR C:/app

# Copiar Node.js y los módulos ya compilados desde el builder
COPY --from=builder C:/nodejs C:/nodejs
COPY --from=builder C:/app/node_modules ./node_modules

# Copiar explícitamente todas las carpetas y archivos necesarios
COPY --from=builder C:/app/database ./database
COPY --from=builder C:/app/routes ./routes
COPY --from=builder C:/app/services ./services
COPY --from=builder C:/app/server2.js ./
COPY --from=builder C:/app/.env .
COPY --from=builder C:/app/.env.docker .

ENV PATH="C:\nodejs;%PATH%"

EXPOSE 3000
CMD ["node", "server2.js"]