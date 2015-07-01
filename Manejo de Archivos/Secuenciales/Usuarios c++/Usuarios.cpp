#include <cstdlib>
#include <iostream>
#include <fstream>
#include <conio.h>
#include <windows.h>
//Librerias Definidas
#include "Funciones.h"
#include "Captura.h"
#include "Reportes.h"
#include "Busqueda.h"
#include "Eliminar.h"

/////////////////////

using namespace std;
int main()
{
    menu:
     int MenuIni=71,Tecla;
    char Cadena[255];
   
  Text(MenuIni,25,12);
  system("cls");
    
Text(MenuIni,13,2);cout<<"================================================";
Text(MenuIni,13,3);cout<<"            *SISTEMA DE USUARIOS*               ";
Text(MenuIni,13,4);cout<<"================================================";

 Text(MenuIni,25,7); cout<<"                         ";
 Text(MenuIni,25,8); cout<<" 1-CAPTURAR USUARIOS     ";
 Text(MenuIni,25,9); cout<<" 2-REPORTE DE USUARIOS   ";
 Text(MenuIni,25,10);cout<<" 3-BUSQUEDA DE USUARIOS  ";
 Text(MenuIni,25,11);cout<<" 4-ELIMINAR USUARIOS     ";
 Text(MenuIni,25,12);cout<<"                         "; 
     Text(116,25,13);cout<<"                         ";
     Text(116,25,14);cout<<"   [ESC. Para Salir]     ";
     Text(116,25,15);cout<<"                         ";
     
   Text(116,25,16);
   Tecla=getch();
   

    //Captura de Usuarios
   if (Tecla==49)
   {
      Text(MenuIni,25,12);
      system("cls");
      CapturaFuncion();          
   }
    //Reportes de Usuarios
   if (Tecla==50)
   {
      Text(MenuIni,25,12);
      system("cls");
      Reportes();          
   }
    //Busqueda  de Usuarios
   if (Tecla==51)
   {
      Text(MenuIni,25,12);
      system("cls");
      Buscar();          
   }
    //Eliminar Archivo
   if (Tecla==52)
   {
   Text(MenuIni,25,12);
      system("cls");
      Eliminar();
          
             
   }
   
   //ESC para salir del sistema
   if(Tecla==27){
            
       return EXIT_SUCCESS;    
        
  }
   
   
   goto menu;
   
  
}



