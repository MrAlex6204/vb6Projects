#include <cstdlib>
#include <iostream>
#include <fstream>
#include <windows.h>
#include <conio.h>
//libreria para caracteres de relleno
#include <iomanip>
//Libreria para conversion de cadenas de caracteres
#include <sstream>
#include <string>


using namespace std;
void Buscar()
{
int MenuIni=71,i=8,Cantidad=0;
bool Encontrado=false;
char Contenido[255];
char Nom[255],Ape[255],Dir[255],Tel[255],Buscar[255];
string busc,nomb;

stringstream Convert;
stringstream Convert1;
ifstream   Lee("Usuarios.txt");

Text(MenuIni,13,2);cout<<"================================================";
Text(MenuIni,13,3);cout<<"            *BUSQUEDA DE USUARIOS*              ";
Text(MenuIni,13,4);cout<<"================================================";

Text(MenuIni,25,6); cout<<"Nombre(s):";
Text(MenuIni,35,6);cin.getline(Buscar,255);

//Conversion de char a string
Convert<<Buscar;
Convert>>busc;
//////////////////////

//Muestra en Pantalla Los Nombres de Los Campos
Text(MenuIni,1,i);
cout<<setw(20)<<setfill('-')<<left<<"NOMBRE(S)"<<setw(20);
cout<<"APELLIDOS"<<setw(20)<<"DIRECCION"<<setw(15)<<"TEL"<<endl;    


//Leee El Primer Registro del Archivo
  Lee.getline(Nom,255);
    Convert1<<Nom;
    Convert1>>nomb;
    
   while(!Lee.eof())   
   {
  
               
                    
     if(Nom==busc)          
     {
     Lee.getline(Ape,255);
     Lee.getline(Dir,255);
     Lee.getline(Tel,255);
                              
       Text(MenuIni,1,i+2);
       Encontrado=true;     
       Cantidad++;      
      
      //Mustra en Pantalla los Datos
      cout<<setw(20)<<setfill(' ')<<left<<Nom<<setw(20);
      cout<<Ape<<setw(20)<<Dir<<setw(15)<<Tel<<endl; 
      
      i++;
      //Se incrementa i para que el Proximo Dato
      //encontrado se brinque un renglon;           
         
         
       
     }
     else{
          
       //cin.get();                       
     //Lee Otro Registro en Caso de Que no 
     //Coinsida Con el Nombre a Buscar
  
      Lee.getline(Ape,255);
      Lee.getline(Dir,255);
      Lee.getline(Tel,255);         
      }
     
     Lee.getline(Nom,255); 
     Convert1<<Nom;
     Convert1>>nomb;

  }

Lee.close();
Text(MenuIni,1,i+3);
cout<<setw(75)<<setfill('-')<<left<<""<<endl;
  
Text(MenuIni,1,i+5);
if (Encontrado==false)
{
cout<<"Resultado de la Busqueda: No Se Encontro Usuario!!"<<endl;
}
else
{
cout<<"Resultado de la Busqueda: "<<Cantidad<<endl<<endl;
}  
cout<<"PRESIONE CUALQUIER TECLA.....";
cin.get();


}
