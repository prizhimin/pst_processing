@echo off
chcp 1251
set KEYWORDS=Ротек Архангельск Норильск Мокулаевск "ТЭЦ-2" котельная ЦБК аэропорт

for %%k in (%KEYWORDS%) do (
    echo Поиск по ключевому слову: "%%k"

    python.exe .\main.py --output-dir "Y:\Output\%%k\Smirnov" --body "%%k" Y:\PST\Alexandr.Smirnov.pst
    python.exe .\main.py --output-dir "Y:\Output\%%k\Smirnov" --body "%%k" Y:\PST\Alexandr.Smirnov.Vault.pst
    python.exe .\main.py --output-dir "Y:\Output\%%k\Batueva" --body "%%k" Y:\PST\Darya.Batueva.pst
    python.exe .\main.py --output-dir "Y:\Output\%%k\Batueva" --body "%%k" Y:\PST\Darya.Batueva.Vault.pst
    python.exe .\main.py --output-dir "Y:\Output\%%k\Batueva" --body "%%k" Y:\PST\Darya.Batueva.Vault.pst
    python.exe .\main.py --output-dir "Y:\Output\%%k\Ivanova" --body "%%k" Y:\PST\Ekaterina.A.Ivanova.pst
    python.exe .\main.py --output-dir "Y:\Output\%%k\Ivanova" --body "%%k" Y:\PST\Ekaterina.A.Ivanova.Vault.pst
    python.exe .\main.py --output-dir "Y:\Output\%%k\Andronchik" --body "%%k" Y:\PST\Evgeniy.Andronchik.pst
    python.exe .\main.py --output-dir "Y:\Output\%%k\Andronchik" --body "%%k" Y:\PST\Evgeniy.Andronchik.Vault.pst
    python.exe .\main.py --output-dir "Y:\Output\%%k\Beynarovich" --body "%%k" Y:\PST\Evgeniy.Beynarovich.pst
    python.exe .\main.py --output-dir "Y:\Output\%%k\Beynarovich" --body "%%k" Y:\PST\Evgeniy.Beynarovich.Vault.pst
    python.exe .\main.py --output-dir "Y:\Output\%%k\Zhiltsova" --body "%%k" Y:\PST\Galina.Zhiltsova.pst
    python.exe .\main.py --output-dir "Y:\Output\%%k\Zhiltsova" --body "%%k" Y:\PST\Galina.Zhiltsova.Vault.pst
    python.exe .\main.py --output-dir "Y:\Output\%%k\Polunina" --body "%%k" Y:\PST\Inna.Polunina.pst
    python.exe .\main.py --output-dir "Y:\Output\%%k\Polunina" --body "%%k" Y:\PST\Inna.Polunina.Vault.pst
    python.exe .\main.py --output-dir "Y:\Output\%%k\Kim" --body "%%k" Y:\PST\Irina.Kim.pst
    python.exe .\main.py --output-dir "Y:\Output\%%k\Kim" --body "%%k" Y:\PST\Irina.Kim.Vault.pst
    python.exe .\main.py --output-dir "Y:\Output\%%k\Neganov" --body "%%k" Y:\PST\Neganov.Evgeniy.pst
    python.exe .\main.py --output-dir "Y:\Output\%%k\Neganov" --body "%%k" Y:\PST\Neganov.Evgeniy.Vault.pst
    python.exe .\main.py --output-dir "Y:\Output\%%k\Listopad" --body "%%k" Y:\PST\Rodion.Listopad.pst
    python.exe .\main.py --output-dir "Y:\Output\%%k\Listopad" --body "%%k" Y:\PST\Rodion.Listopad.Vault.pst
    python.exe .\main.py --output-dir "Y:\Output\%%k\Evsikov" --body "%%k" Y:\PST\Sergey.Evsikov.pst
    python.exe .\main.py --output-dir "Y:\Output\%%k\Evsikov" --body "%%k" Y:\PST\Sergey.Evsikov.Vault.pst
)