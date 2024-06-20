# MaradjTalpon

Ennek a kódnak a lényege az, hogy [PowerPoint](https://en.wikipedia.org/wiki/Microsoft_PowerPoint) prezentációt hozzon létre a *Maradj Talpon!* játékhoz úgy, hogy a kérdéseket véletlenszerű sorrendbe helyezi.

A kérdések a hozzájuk tartozó válaszokkal a ```kerdesek``` könyvtárban találhatóak. Itt a szöveges fájlnak a neve maga a kategória, valamint a kategórián belül egy kérdés egy sorban van, amit követ a hozzá tartozó válasz a következő sorban.

Amennyiben új kategóriákra van szükség, az ahhoz tartozó fájlt létre kell hozni a ```kerdesek``` könyvtárban, a fájlt pedig fel kell tölteni a kérdésekkel+válaszokkal. A program alapértelmezetten figyelembe veszi az összes kategóriát.

## Futtatás

A prezentáció generálásához egyszerűen le kell futtatni a ```MaradjTalponGenerator.py``` kódot.

A generált kérdések számát a kódon belül a ```main()``` függvényen belül lehet módosítani az alábbi sorban:

```Python
max_num_questions = 100
```

Egy másik lehetőség az, hogy a kérdések számát futtatáskor, *Command Line Argument*-ként kapja meg a program.
