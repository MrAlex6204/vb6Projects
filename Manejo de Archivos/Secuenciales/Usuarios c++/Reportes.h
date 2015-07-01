#include <cstdlib>
#include <iostream>
#include <fstream>
#include <windows.h>
#include <conio.h>

#include <iomanip>
using namespace std;
void Reportes()
{
int MenuIni=71,i=10;
char Nom[255],Ape[255],Dir[255],Tel[255];
ifstream   Lee("Usuarios.txt");

Text(MenuIni,13,2);cout<<"================================================";
Text(MenuIni,13,3);cout<<"            *REPORTE DE USUARIOS*               ";
Text(MenuIni,13,4);cout<<"================================================";
     
     
Text(MenuIni,1,8);
cout<<setw(20)<<setfill('-')<<left<<"NOMBRE(S)"<<setw(20);
cout<<"APELLIDOS"<<setw(20)<<"DIRECCION"<<setw(15)<<"TEL"<<endl;

Lee.getline(Nom,255);
Lee.getline(Ape,255);
Lee.getline(Dir,255);
Lee.getline(Tel,255);

 
     while(!Lee.eof())
    {
                      
                      
Text(MenuIni,1,i);
 cout<<setw(20)<<setfill(' ')<<left<<Nom<<setw(20);
 cout<<Ape<<setw(20)<<Dir<<setw(15)<<Tel<<endl;    
 Lee.getline(Nom,255);
Lee.getline(Ape,255);
Lee.getline(Dir,255);
Lee.getline(Tel,255);
               i++;        
    }
Lee.close();    

Text(MenuIni,1,i+1);
cout<<setw(75)<<setfill('-')<<left<<""<<endl;

Text(MenuIni,1,i+3);
cout<<"PRESIONE CUALQUIER TECLA.....";
    cin.get();
}



