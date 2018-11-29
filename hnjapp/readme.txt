#this project depends on below projects(32bit, because sybase driver is 32bit and pyodbc can only access the related-bit version)
    Pillow
    PyPDF2
    SQLAlchemy
    opencv-python
    pylint
    pyodbc
    pytesseract
    pywin32
    xlwings
    yapf
#below is a bat file for create the venv and dependancies
#<bat>
@echo off
if '%1' == '' (
	@set fldr=3dev
) else (
	@set fldr=%1
)
if exist %fldr% (
	echo "virtual environment(%fldr%) already exists"
	exit /B
)
@echo === Begin to create virtual environment(%fldr%) ===
@python -m venv %fldr%
@echo === virtual environment(%fldr%) created ===
cmd /k "%cd%\%fldr%\scripts\activate.bat & @echo === Begin to install the dependancy libraries === & python -m pip install pip --upgrade & pip install Pillow & pip install PyPDF2 & pip install SQLAlchemy & pip install opencv-python & pip install pylint & pip install pyodbc & pip install pytesseract & pip install pywin32 & pip install xlwings & pip install yapf & exit & exit"
@echo === Dependancy installation done ===
echo on
#</bat>
