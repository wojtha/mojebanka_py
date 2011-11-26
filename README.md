# Mojebanka konverter transakcí

Konvertuje textový výpis transakcí z TXT do QIF nebo CVS formátu.

Určeno zejména pro ty, kteří mají obyčejné internetové bankovnictví u Komerční
banky, kde je povolen export transakcí pouze do textového formátu.

## Použití

python mojebanka.py --format=qif Tisk_20110910153349.txt
python mojebanka.py --format=cvs *.txt

## Parametry

 * --help         Nápověda - tento text.
 * --format=qif   Cílový formát: qif, cvs

## Licence

Dual licensed under Apache License 2.0 and GNU GPL 2.0.
