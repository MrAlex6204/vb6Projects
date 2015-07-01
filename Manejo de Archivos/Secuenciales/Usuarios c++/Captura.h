#include <cstdlib>
#include <iostream>
#include <fstream>
#include <windows.h>
#include <conio.h>


using namespace std;
void CapturaFuncion()
{

int MenuIni=71,Tecla;
char Nom[255],Ape[255],Dir[255],Tel[255];
ofstream  Escribe("Usuarios.txt",ios::app);

Text(MenuIni,13,2);cout<<"================================================";
Text(MenuIni,13,3);cout<<"            *CAPTURA DE USUARIOS*               ";
Text(MenuIni,13,4);cout<<"================================================"; 

Text(MenuIni,25,8); cout<<"Nombre(s):"; 
Text(MenuIni,25,9); cout<<"Apellidos:"; 
Text(MenuIni,25,10);cout<<"Direccion:"; 
Text(MenuIni,25,11);cout<<"Telefono :"; 
//Captura los Datos en Las Variables

Text(MenuIni,35,8);cin.getline(Nom,255);
Text(MenuIni,35,9);cin.getline(Ape,255);
Text(MenuIni,35,10);cin.getline(Dir,255);
Text(MenuIni,35,11);cin.getline(Tel,255);

Text(MenuIni,25,14);cout<<"Desea Guardar Los Cambios [S/N]:"; 
Text(MenuIni,57,14); Tecla=getch();

   
    if(Tecla==110|Tecla==78){
  
    }//si presiona "N"
    
    if(Tecla==83|Tecla==115)// si presiona "S"
    {
         Escribe<<Nom<<endl;
         Escribe<<Ape<<endl;
         Escribe<<Dir<<endl;      
         Escribe<<Tel<<endl;      
    }
  
system("cls");


}
   

