Doporučení aktivace EBC

Tento skript slouží k automatickému vyhodnocení, zda mají produkty určené pro aktivaci EBC (Enhanced Brand Content) kompletní data, tedy fotografie a anotace.

Díky tomu snadno zjistíte, které položky mají připravená data a které je potřeba ještě doplnit.

Funkce skriptu

Aktuální verze umí:

✅ Načíst seznam všech produktů z good.xlsx

pracuje s listem, který obsahuje minimálně sloupec:

OID_zbozi (unikátní identifikátor produktu)

pokud chybí sloupce Fotografie nebo Anotace, skript je automaticky vytvoří jako prázdné → nedojde k pádu

✅ Načíst více vstupních souborů se seznamy produktů

pro každý .xlsx soubor vyhodnotí produkty, kde je v některém ze dvou sloupců hodnota ano:

EBC požadovaný stav

EBC 2 požadovaný stav

✅ Spojit data z EBC souborů s informacemi z good.xlsx

✅ Vyhodnotit kompletnost dat pro každý produkt:

Fotografie + Anotace – produkt má obojí

Pouze fotografie – chybí anotace

Pouze anotace – chybí fotografie

Chybí obojí – nemá ani jedno

✅ Bezpečně ignorovat dočasné soubory Excelu ~$*.xlsx, takže se zpracovávají jen skutečné zdrojové soubory

✅ Vytvořit výstupní soubor doporučení_aktivace_EBC.xlsx s kompletním přehledem

✅ Souhrn stavu dat na konci běhu přímo v terminálu

✅ Ošetřit chyby:

soubor otevřený v Excelu (PermissionError)

poškozený soubor (BadZipFile)

prázdný list nebo chybějící sloupce

Požadované soubory

Skript pracuje se soubory ve složce, kde je uložen (doporuceni_aktivace_EBC.py).

1. good.xlsx

Hlavní databáze všech produktů.
Musí obsahovat minimálně:

Sloupec	Povinný	Popis
OID_zbozi	ANO	Unikátní identifikátor produktu
Fotografie	NE (doplní se prázdný)	URL, název souboru, nebo cokoliv, co označuje fotku
Anotace	NE (doplní se prázdný)	Textová anotace / popis produktu

Tip:
Pokud Fotografie nebo Anotace v souboru chybí, skript je automaticky založí jako prázdné, aby nenastal pád.

2. Vstupní soubory se seznamy produktů k aktivaci

Každý soubor musí být ve formátu .xlsx a obsahovat alespoň:

Sloupec	Povinný	Popis
(první sloupec)	ANO	Identifikátor produktu (musí odpovídat OID_zbozi z good.xlsx)
EBC požadovaný stav	ANO	Hodnota ano / ne
EBC 2 požadovaný stav	ANO	Hodnota ano / ne

Produkty, kde je v alespoň jednom z těchto dvou sloupců hodnota ano → jsou vyhodnoceny jako kandidáti pro aktivaci EBC.

Výstup

Po dokončení zpracování se vygeneruje soubor:

doporučení_aktivace_EBC.xlsx


Každý řádek obsahuje:

Sloupec	Popis
Zdrojový soubor	Z jakého vstupního souboru produkt pochází
OID_zbozi	Unikátní identifikátor produktu
Fotografie	Původní hodnota z good.xlsx
Anotace	Původní hodnota z good.xlsx
Stav dat	Výsledek kontroly (viz popisy výše)
... další sloupce z původního vstupního souboru	
Příklad výstupu
Zdrojový soubor	OID_zbozi	Fotografie	Anotace	Stav dat
2025_04_svítidla_RAJL_JIKU.xlsx	1001	foto1.jpg	Popis A	Fotografie + Anotace
2025_04_svítidla_RAJL_JIKU.xlsx	1002	foto2.jpg	(prázdné)	Pouze fotografie
2025_06_svítidla_RAJL_JIKU.xlsx	1003	(prázdné)	Popis B	Pouze anotace
2025_06_svítidla_RAJL_JIKU.xlsx	1004	(prázdné)	(prázdné)	Chybí obojí
Postup spuštění

Umístěte všechny potřebné soubory do jedné složky:

doporuceni_aktivace_EBC.py (skript)

good.xlsx (databáze produktů)

všechny vstupní soubory k aktivaci (např. 2025_04_svítidla_RAJL_JIKU.xlsx)

Ujistěte se, že žádný soubor není otevřený v Excelu.

Otevřete příkazovou řádku v této složce:

cd "C:\Práce na datech\Doporučení aktivace EBC"


Spusťte skript:

python doporuceni_aktivace_EBC.py


Po dokončení se zobrazí souhrn a vytvoří se výstupní soubor doporučení_aktivace_EBC.xlsx.

Poznámky

Skript ignoruje všechny soubory, které začínají na ~$ (dočasné soubory, které vytváří Excel při otevření sešitu).

Pokud je soubor otevřený v Excelu, skript vypíše upozornění a přeskočí ho.

Pokud good.xlsx neobsahuje povinný sloupec OID_zbozi, skript se ukončí s chybou (bez OID není možné párovat produkty).

Pokud Fotografie a/nebo Anotace chybí, automaticky se vytvoří jako prázdné sloupce.

Souhrn výsledků se zobrazuje přímo v terminálu.

Příklad souhrnu v terminálu
📊 Souhrn stavů (všechny soubory):
  - Fotografie + Anotace: 157
  - Pouze fotografie: 21
  - Pouze anotace: 12
  - Chybí obojí: 3

✅ HOTOVO! Výstup uložen jako:
C:\Práce na datech\Doporučení aktivace EBC\doporučení_aktivace_EBC.xlsx

Požadavky na prostředí

Python 3.9+ (doporučená verze: 3.13)

Balíčky:

pandas
openpyxl


Instalace:

pip install pandas openpyxl
