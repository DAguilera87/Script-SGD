# Script-SGD
Mi primer proyecto en Python

#########################################
# Paso a paso: Agregar .venv a .gitignore
1.-Créalo con este comando (o manualmente):
    echo .venv/ > .gitignore
2.-Si .venv ya estaba siendo rastreado por Git:
    Si creaste el entorno virtual antes de ignorarlo, Git ya lo tiene en el índice, y seguirá queriendo subirlo a GitHub. Para solucionarlo:
    git rm -r --cached .venv
3.-Confirma los cambios en Git
    git add .gitignore
    git commit -m "Agrego .venv al .gitignore para excluir entorno virtual"
    git push
####################################################

####################################################
# Creación de un entorno virtual en VS Code
1.-Crear el entorno virtual
    python -m venv .venv
2.-Activar el entorno virtual
    .venv\Scripts\Activate.ps1
3.-Seleccionar intérprete en VS Code
Presiona Ctrl + Shift + P → escribe Python: Select Interpreter → elige el que diga .venv.
###################################################

###################################################
# Dependencias
1.-Instalar todas las dependencias
    pip install -r requirements.txt
2.-Verificar que todo esté instalado
    pip list
###################################################