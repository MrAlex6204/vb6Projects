#include <cstdlib>
#include <iostream>
#include <fstream>
#include <windows.h>
#include <conio.h>
using namespace std;

//-------------------------------------------------------------------
//Funcion para el color de texto y posicionamiento del mismo
// en pantalla 
void Text(int Color,int x , int y)
{     

// Posicion del Texto
HANDLE hConsoleOutput;
HANDLE hConsoleOut;

COORD CursorPos = {x, y};
hConsoleOutput = GetStdHandle(STD_OUTPUT_HANDLE);
SetConsoleCursorPosition(hConsoleOutput, CursorPos);

// Color de Texto
      
HANDLE hConsole;
hConsole=GetStdHandle(STD_OUTPUT_HANDLE);
SetConsoleTextAttribute(hConsole,Color);
}
