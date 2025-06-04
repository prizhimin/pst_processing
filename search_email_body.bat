@echo off
if "%1"=="" (
    echo Не указан параметр для body
    exit /b 1
)

python.exe .\main.py --output-dir "R:\Output\%1\Smirnov" --body "%1" Y:\PST\Alexandr.Smirnov.pst
python.exe .\main.py --output-dir "R:\Output\%1\Smirnov" --body "%1" Y:\PST\Alexandr.Smirnov.Vault.pst
python.exe .\main.py --output-dir "R:\Output\%1\Batueva" --body "%1" Y:\PST\Darya.Batueva.pst
python.exe .\main.py --output-dir "R:\Output\%1\Batueva" --body "%1" Y:\PST\Darya.Batueva.Vault.pst
python.exe .\main.py --output-dir "R:\Output\%1\Batueva" --body "%1" Y:\PST\Darya.Batueva.Vault.pst
python.exe .\main.py --output-dir "R:\Output\%1\Ivanova" --body "%1" Y:\PST\Ekaterina.A.Ivanova.pst
python.exe .\main.py --output-dir "R:\Output\%1\Ivanova" --body "%1" Y:\PST\Ekaterina.A.Ivanova.Vault.pst
python.exe .\main.py --output-dir "R:\Output\%1\Andronchik" --body "%1" Y:\PST\Evgeniy.Andronchik.pst
python.exe .\main.py --output-dir "R:\Output\%1\Andronchik" --body "%1" Y:\PST\Evgeniy.Andronchik.Vault.pst
python.exe .\main.py --output-dir "R:\Output\%1\Beynarovich" --body "%1" Y:\PST\Evgeniy.Beynarovich.pst
python.exe .\main.py --output-dir "R:\Output\%1\Beynarovich" --body "%1" Y:\PST\Evgeniy.Beynarovich.Vault.pst
python.exe .\main.py --output-dir "R:\Output\%1\Zhiltsova" --body "%1" Y:\PST\Galina.Zhiltsova.pst
python.exe .\main.py --output-dir "R:\Output\%1\Zhiltsova" --body "%1" Y:\PST\Galina.Zhiltsova.Vault.pst
python.exe .\main.py --output-dir "R:\Output\%1\Polunina" --body "%1" Y:\PST\Inna.Polunina.pst
python.exe .\main.py --output-dir "R:\Output\%1\Polunina" --body "%1" Y:\PST\Inna.Polunina.Vault.pst
python.exe .\main.py --output-dir "R:\Output\%1\Kim" --body "%1" Y:\PST\Irina.Kim.pst
python.exe .\main.py --output-dir "R:\Output\%1\Kim" --body "%1" Y:\PST\Irina.Kim.Vault.pst
python.exe .\main.py --output-dir "R:\Output\%1\Neganov" --body "%1" Y:\PST\Neganov.Evgeniy.pst
python.exe .\main.py --output-dir "R:\Output\%1\Neganov" --body "%1" Y:\PST\Neganov.Evgeniy.Vault.pst
python.exe .\main.py --output-dir "R:\Output\%1\Listopad" --body "%1" Y:\PST\Rodion.Listopad.pst
python.exe .\main.py --output-dir "R:\Output\%1\Listopad" --body "%1" Y:\PST\Rodion.Listopad.Vault.pst
python.exe .\main.py --output-dir "R:\Output\%1\Evsikov" --body "%1" Y:\PST\Sergey.Evsikov.pst
python.exe .\main.py --output-dir "R:\Output\%1\Evsikov" --body "%1" Y:\PST\Sergey.Evsikov.Vault.pst