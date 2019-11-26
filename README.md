# TETO-Optim
Aplikacija - Optimizacija rada termoelektrane TETO Zagreb

======================== R E A D M E ========================

******************** ime aplikacije *****************************

TETO-Optim

******************** opis aplikacije ****************************

Aplikacija je nastala kao kao dio projekta za HEP d.o.o. Cilj projekta je bio napraviti matematički model i aplikaciju za optimiranje rada termoelektrane TETO u Zagrebu te je sami matematički model napravljen u MATLAB aplikaciji.

Relevantni dijelovi koda iz MATLAB programskog jezika su pomocu Library Compilera konvertirani u odgovarajuce .NET Assembly (MWArray.dll i OptimizacijaTETO.dll) koje su koristene u izradi C# aplikacije.

Da bi se mogla izvrsiti optimizacija pomocu "TETO Zagreb" aplikacije potrebno imati instaliran MATLAB Runtime R2015a 8.5 (moguće besplatno preuzimanje na https://www.mathworks.com/products/compiler/matlab-runtime.html).

Za demonstriranje rada aplikacije izvorni kod je prepravljen tako da pokretanjem optimizacije (kontrola "OPTIMIRAJ") program zapravo simulira optimizaciju. Time je postignuto da korisnik ne mora instalirati navedeni MATLAB Runtime.

Kao primjer dani su rezultati optimizacije u datoteci "rezultati.data" koji se mogu ucitati u aplikaciju.


******************** napomena ***********************************

Library "OptimizacijaTETO.dll" potrebno je otpakirati (npr. pomoću 7-zip). Nalazi se u TETO-Optim/Resources/.

File "rezultati.data" nalazi se u TETO-Optim/bin/Debug/.

Screenshotovi aplikacije dani su u datoteci "Aplikacija.pdf".


******************** verzija ************************************

v1.0 (2016)

******************** autor **************************************

Stjepko Katulic
