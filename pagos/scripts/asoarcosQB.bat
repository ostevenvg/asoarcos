set conda_dir="C:\Users\villalta\AppData\Local\miniconda3"
call %conda_dir%\condabin\_conda_activate.bat asoarcos
set path=%conda_dir%\envs\asoarcos\;%path%
python "C:\Users\villalta\Documents\Personal\repos\asoarcos\pagos\scripts\gui.py"
